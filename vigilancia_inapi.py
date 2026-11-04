# -*- coding: utf-8 -*-
"""
Vigilancia de marcas INAPI
--------------------------
Revisa los Estados Diarios de Marcas de INAPI
(https://tramites.inapi.cl/EstadosDiariosMarcas) y avisa (ventana + correo) si
aparece una marca igual, parecida o que comparte una palabra con las marcas
vigiladas. Cada revisión genera un informe con cuántas marcas se revisaron.

Uso:
    pythonw vigilancia_inapi.py            (revisión automática, avisa solo si hay novedades)
    pythonw vigilancia_inapi.py --inicio   (al encender el PC: espera la conexión y revisa)
    pythonw vigilancia_inapi.py --manual   (revisión manual, siempre muestra el resultado)
    pythonw vigilancia_inapi.py --panel    (ventana para administrar marcas, ajustes y correo)
"""
import base64
import csv
import datetime
import difflib
import hashlib
import html
import io
import json
import os
import re
import sys
import time
import traceback
import unicodedata
import zipfile
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

# ---------------------------------------------------------------- rutas
CARPETA = Path(__file__).resolve().parent
ARCH_CLAVES = CARPETA / "palabras_clave.txt"
REPORTES = CARPETA / "informes"
HISTORIAL = CARPETA / "historial_coincidencias.csv"
HIST_REVISIONES = CARPETA / "historial_revisiones.csv"
DATOS = CARPETA / "datos"
ARCHIVOS = DATOS / "archivos_inapi"
ESTADO = DATOS / "estado.json"
REGISTRO = DATOS / "registro.txt"
BLOQUEO = DATOS / "en_ejecucion.lock"
CONFIG = DATOS / "config.json"
ULTIMO_INFORME = REPORTES / "ultimo_informe.html"
NOMBRE_INICIO = "Vigilancia Marcas INAPI"   # nombre en "Aplicaciones de inicio" de Windows
CLAVE_RUN = r"Software\Microsoft\Windows\CurrentVersion\Run"

# ---------------------------------------------------------------- ajustes fijos
VERSION = "1.6"
URL_PAGINA = "https://tramites.inapi.cl/EstadosDiariosMarcas"
TIPOS = {
    "5": "Presentaciones (solicitudes nuevas)",
    "2": "Estado diario (versión extendida)",
}
UMBRAL_PARECIDO = 0.87       # 0 a 1: qué tan parecida debe ser una marca para avisar
MIN_LETRAS_PARECIDO = 7      # solo las palabras de 7+ letras buscan marcas "parecidas"
DIAS_GUARDAR_ARCHIVOS = 45
INFORMES_A_GUARDAR = 90
NAVEGADOR = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")
TIPOS_SIGNO = (r"Mixta|Denominativa|Figurativa|De\s+patr[oó]n|Tridimensional|Sonora|Olfativa|"
               r"Hologr[aá]fica|De\s+movimiento|De\s+posici[oó]n|Etiqueta")
RX_TIPO_SIGNO = re.compile(r"(?<![A-Za-zÁÉÍÓÚáéíóúÑñ])(" + TIPOS_SIGNO + r")(?![A-Za-zÁÉÍÓÚáéíóúÑñ])", re.I)

IGNORADAS_INICIALES = ("de, del, la, las, el, los, lo, y, e, o, u, en, por, para, con, a, al, "
                       "the, of, and, for, pro, spa, ltda, sa, eirl, chile, marca, marcas")

CLAVES_INICIALES = """# Palabras a vigilar en INAPI (una por línea).
# Puedes agregar o borrar líneas y guardar. Las líneas que empiezan con # se ignoran.
# No importan mayúsculas ni tildes.
# Las palabras de 7 letras o más también detectan marcas parecidas
# (por ejemplo "Calagenda" o "Kalagend" para "Calagend").
# Si cambias esta lista, la próxima revisión vuelve a mirar los últimos 30 días.

EOL
EOL Research
APM
Amor por la Marca
Calagend
APM Pro
"""

CONFIG_INICIAL = {
    "avisar_sin_novedades": False,
    "una_palabra": True,
    "palabras_ignoradas": IGNORADAS_INICIALES,
    "correo": {
        "activo": False,
        "destinatarios": "",
        "remitente": "",
        "clave": "",
        "servidor": "smtp.gmail.com",
        "puerto": 465,
        "resumen_sin_novedades": False,
    },
}


# ---------------------------------------------------------------- utilidades
def log(msg):
    linea = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(linea)
    try:
        DATOS.mkdir(parents=True, exist_ok=True)
        with open(REGISTRO, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
    except Exception:
        pass


def aviso(titulo, texto, preguntar=False):
    """Ventana emergente de Windows. Devuelve True si el usuario respondió 'Sí'."""
    try:
        import ctypes
        flags = 0x40 | 0x40000 | (0x4 if preguntar else 0)  # información + siempre visible (+ Sí/No)
        return ctypes.windll.user32.MessageBoxW(0, texto, titulo, flags) == 6
    except Exception:
        print(f"[{titulo}] {texto}")
        return False


def abrir(ruta):
    try:
        os.startfile(str(ruta))  # solo Windows
    except Exception:
        pass


def normalizar(s):
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def preparar_ssl():
    try:
        import truststore  # usa los certificados de Windows
        truststore.inject_into_ssl()
    except Exception:
        pass


def cargar_estado():
    try:
        return json.loads(ESTADO.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_estado(est):
    DATOS.mkdir(parents=True, exist_ok=True)
    tmp = ESTADO.with_suffix(".tmp")
    tmp.write_text(json.dumps(est, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, ESTADO)


def tomar_bloqueo():
    DATOS.mkdir(parents=True, exist_ok=True)
    try:
        if BLOQUEO.exists() and time.time() - BLOQUEO.stat().st_mtime > 7200:
            BLOQUEO.unlink()
        fd = os.open(str(BLOQUEO), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        return False


def soltar_bloqueo():
    try:
        BLOQUEO.unlink()
    except Exception:
        pass


# ---------------------------------------------------------------- configuración
def cargar_config():
    conf = json.loads(json.dumps(CONFIG_INICIAL))
    try:
        guardada = json.loads(CONFIG.read_text(encoding="utf-8"))
        correo = guardada.pop("correo", {})
        conf.update(guardada)
        conf["correo"].update(correo)
    except Exception:
        pass
    return conf


def guardar_config(conf):
    DATOS.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(conf, ensure_ascii=False, indent=1), encoding="utf-8")


def _dpapi(datos, cifrar):
    """Cifra/descifra con la cuenta de Windows del usuario (nadie más puede leerlo)."""
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(datos, len(datos))
    entrada = BLOB(len(datos), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    salida = BLOB()
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    if cifrar:
        ok = crypt32.CryptProtectData(ctypes.byref(entrada), "Vigilancia INAPI", None, None, None, 0,
                                      ctypes.byref(salida))
    else:
        ok = crypt32.CryptUnprotectData(ctypes.byref(entrada), None, None, None, None, 0, ctypes.byref(salida))
    if not ok:
        raise OSError("No se pudo cifrar/descifrar la contraseña")
    resultado = ctypes.string_at(salida.pbData, salida.cbData)
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree(ctypes.cast(salida.pbData, ctypes.c_void_p))
    return resultado


def cifrar_clave(texto):
    if not texto:
        return ""
    try:
        return "dpapi:" + base64.b64encode(_dpapi(texto.encode("utf-8"), True)).decode()
    except Exception:
        return "b64:" + base64.b64encode(texto.encode("utf-8")).decode()


def descifrar_clave(guardada):
    try:
        if guardada.startswith("dpapi:"):
            return _dpapi(base64.b64decode(guardada[6:]), False).decode("utf-8")
        if guardada.startswith("b64:"):
            return base64.b64decode(guardada[4:]).decode("utf-8")
    except Exception:
        pass
    return ""


def ejecutable_sin_ventana():
    exe = Path(sys.executable)
    w = exe.with_name("pythonw.exe")
    return str(w if w.exists() else exe)


def inicio_activado():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CLAVE_RUN) as k:
            winreg.QueryValueEx(k, NOMBRE_INICIO)
            return True
    except Exception:
        return False


def activar_inicio(activar):
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CLAVE_RUN, 0, winreg.KEY_SET_VALUE) as k:
        if activar:
            orden = f'"{ejecutable_sin_ventana()}" "{Path(__file__).resolve()}" --inicio'
            winreg.SetValueEx(k, NOMBRE_INICIO, 0, winreg.REG_SZ, orden)
        else:
            try:
                winreg.DeleteValue(k, NOMBRE_INICIO)
            except FileNotFoundError:
                pass


# ---------------------------------------------------------------- correo
def destinatarios_de(correo):
    return [d.strip() for d in re.split(r"[,;\s]+", correo.get("destinatarios", "")) if "@" in d]


def enviar_correo(correo, asunto, cuerpo_html, cuerpo_texto, adjuntos=()):
    import smtplib
    import ssl
    from email.message import EmailMessage
    from email.utils import formataddr
    dest = destinatarios_de(correo)
    remitente = correo.get("remitente", "").strip()
    if not dest:
        raise ValueError("Falta el correo de destino.")
    if not remitente:
        raise ValueError("Falta la cuenta que envía el correo.")
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = formataddr(("Vigilancia de marcas INAPI", remitente))
    msg["To"] = ", ".join(dest)
    msg.set_content(cuerpo_texto)
    msg.add_alternative(cuerpo_html, subtype="html")
    for ruta in adjuntos:
        ruta = Path(ruta)
        if ruta.exists():
            msg.add_attachment(ruta.read_bytes(), maintype="text", subtype="html", filename=ruta.name)
    servidor = correo.get("servidor", "smtp.gmail.com").strip()
    puerto = int(correo.get("puerto") or 465)
    clave = descifrar_clave(correo.get("clave", "")).replace(" ", "")
    preparar_ssl()
    ctx = ssl.create_default_context()
    if puerto == 465:
        s = smtplib.SMTP_SSL(servidor, puerto, context=ctx, timeout=40)
    else:
        s = smtplib.SMTP(servidor, puerto, timeout=40)
        s.ehlo()
        if s.has_extn("starttls"):
            s.starttls(context=ctx)
            s.ehlo()
    try:
        if clave:
            s.login(remitente, clave)
        s.send_message(msg)
    finally:
        try:
            s.quit()
        except Exception:
            pass


def explicar_error_correo(e):
    t = str(e)
    if "535" in t or "Username and Password not accepted" in t or "BadCredentials" in t:
        return ("Gmail rechazó la cuenta o la contraseña. Usa una «contraseña de aplicación» de 16 letras "
                "(no tu contraseña normal) y revisa que la cuenta que envía esté bien escrita.")
    if "getaddrinfo" in t or "timed out" in t.lower() or "10060" in t:
        return "No pude conectarme al servidor de correo. Revisa tu internet."
    return t


# ---------------------------------------------------------------- palabras clave
def leer_lista_claves():
    if not ARCH_CLAVES.exists():
        ARCH_CLAVES.write_text(CLAVES_INICIALES, encoding="utf-8")
    lista, vistas = [], set()
    for linea in ARCH_CLAVES.read_text(encoding="utf-8-sig").splitlines():
        s = linea.strip()
        if s and not s.startswith("#") and normalizar(s) and normalizar(s) not in vistas:
            vistas.add(normalizar(s))
            lista.append(s)
    return lista


def guardar_lista_claves(lista):
    encabezado = "\n".join(l for l in CLAVES_INICIALES.splitlines() if l.startswith("#"))
    ARCH_CLAVES.write_text(encabezado + "\n\n" + "\n".join(lista) + "\n", encoding="utf-8")


def palabras_ignoradas(conf):
    return set(normalizar(conf.get("palabras_ignoradas", "")).split())


class Clave:
    def __init__(self, texto, ignoradas=frozenset(), una_palabra=True):
        self.texto = texto.strip()
        self.norm = normalizar(self.texto)
        self.toks = self.norm.split()
        self.letras = self.norm.replace(" ", "")
        if len(self.toks) == 1 and len(self.letras) <= 4:
            # siglas cortas (EOL, APM): palabra completa, acepta E.O.L. o A P M
            patron = r" ?".join(map(re.escape, self.letras))
        else:
            # acepta las palabras separadas o juntas: "APM Pro" / "APMPRO"
            patron = r" ?".join(map(re.escape, self.toks))
        self.rx = re.compile(r"(?<![a-z0-9])" + patron + r"(?![a-z0-9])")
        self.parecido = len(self.letras) >= MIN_LETRAS_PARECIDO
        # palabras sueltas significativas (solo en marcas de varias palabras)
        self.sueltas = []
        if una_palabra and len(self.toks) > 1:
            self.sueltas = [t for t in self.toks if t not in ignoradas and (len(t) >= 3 or t.isdigit())]


def cargar_claves(conf=None):
    conf = conf or cargar_config()
    ign = palabras_ignoradas(conf)
    return [Clave(s, ign, conf.get("una_palabra", True)) for s in leer_lista_claves()]


def huella_busqueda(claves, conf):
    partes = sorted(c.norm for c in claves) + [f"una_palabra={conf.get('una_palabra')}",
                                               "ign=" + " ".join(sorted(palabras_ignoradas(conf)))]
    return hashlib.sha1("\n".join(partes).encode()).hexdigest()


def mejor_parecido(tokens, clave):
    objetivo, n, largo = clave.letras, len(clave.toks), len(clave.letras)
    sm = difflib.SequenceMatcher(autojunk=False)
    sm.set_seq2(objetivo)
    mejor = 0.0
    for k in range(max(1, n - 1), n + 2):
        for i in range(0, len(tokens) - k + 1):
            cand = "".join(tokens[i:i + k])
            if abs(len(cand) - largo) > 3:
                continue
            sm.set_seq1(cand)
            if sm.real_quick_ratio() < UMBRAL_PARECIDO or sm.quick_ratio() < UMBRAL_PARECIDO:
                continue
            mejor = max(mejor, sm.ratio())
    return mejor


NIVELES = {"exacta": "Exacta", "parecida": "Parecida", "palabra": "Una palabra en común"}


def buscar(texto, claves, marca=""):
    """Devuelve [(palabra_vigilada, nivel, detalle)] encontradas en un registro."""
    t = normalizar(texto)
    # versión con siglas juntas: "e o l research" -> "eol research"
    t2 = re.sub(r"(?<![a-z0-9])[a-z0-9](?: [a-z0-9](?![a-z0-9]))+",
                lambda m: m.group(0).replace(" ", ""), t)
    tokens = t.split()
    tokens_marca = set(normalizar(marca).split())  # "una palabra" solo mira el nombre de la marca
    hallazgos, exactas = [], set()
    for c in claves:
        if c.rx.search(t) or c.rx.search(t2):
            hallazgos.append((c.texto, "exacta", "Exacta"))
            exactas.add(c.norm)
            continue
        if c.parecido:
            p = mejor_parecido(tokens, c)
            if p >= UMBRAL_PARECIDO:
                hallazgos.append((c.texto, "parecida", f"Parecida ({p:.0%})"))
                continue
        comunes = [w for w in c.sueltas if w in tokens_marca]
        if comunes:
            hallazgos.append((c.texto, "palabra", "Una palabra en común: " + ", ".join(f"«{w}»" for w in comunes)))
    # si la palabra en común ya es una marca vigilada encontrada exacta, no repetir
    return [h for h in hallazgos
            if not (h[1] == "palabra" and all(w.strip("«»") in exactas
                                                for w in h[2].split(": ", 1)[1].split(", ")))]


# ---------------------------------------------------------------- lectura de archivos
def extraer_texto(datos):
    if datos[:4] == b"%PDF":
        from pypdf import PdfReader
        lector = PdfReader(io.BytesIO(datos))
        return "\n".join((p.extract_text() or "") for p in lector.pages)
    if datos[:2] == b"PK":
        try:
            import openpyxl
            libro = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
            lineas = []
            for hoja in libro.worksheets:
                for fila in hoja.iter_rows(values_only=True):
                    vals = [str(v).strip() for v in fila if v not in (None, "")]
                    if vals:
                        lineas.append("  ".join(vals))
            return "\n".join(lineas)
        except Exception:
            partes = []
            with zipfile.ZipFile(io.BytesIO(datos)) as z:
                for nombre in z.namelist():
                    if not nombre.endswith("/"):
                        partes.append(extraer_texto(z.read(nombre)))
            return "\n".join(partes)
    texto = None
    for cod in ("utf-8", "latin-1"):
        try:
            texto = datos.decode(cod)
            break
        except UnicodeDecodeError:
            pass
    if "<html" in texto[:3000].lower():
        texto = html.unescape(re.sub(r"<[^>]+>", "\n", texto))
    return texto


def extension(datos):
    if datos[:4] == b"%PDF":
        return ".pdf"
    if datos[:2] == b"PK":
        return ".xlsx"
    return ".txt"


RX_SOLICITUD = re.compile(r"^\s*(?:\d{1,4}\s+)?(\d{6,8})\b")  # en Presentaciones va un correlativo antes
RX_SECCION = re.compile(r"Secci[oó]n\s+(M\s?\d+)\s*:?\s*(.*)", re.I)
RUIDO = [re.compile(p, re.I) for p in (
    r"^estado diario", r"^solicitud\s+representante", r"^\d{1,4}$", r"^p[aá]gina\s+\d+")]


def separar_registros(texto):
    """Agrupa las líneas del PDF en registros: cada uno empieza con un N° de solicitud."""
    regs, actual, seccion, esperar_sub = [], None, "", False
    for linea in texto.splitlines():
        l = linea.strip()
        if not l:
            continue
        m = RX_SECCION.search(l)
        if m and not RX_SOLICITUD.match(l):
            resto = m.group(2).strip()
            seccion = f"Sección {m.group(1).replace(' ', '')}" + (f": {resto}" if resto else "")
            esperar_sub = not resto
            if actual:
                regs.append(actual)
                actual = None
            continue
        ruido = any(r.search(l) for r in RUIDO)
        if esperar_sub and not ruido and not RX_SOLICITUD.match(l):
            seccion += f": {l}"
            esperar_sub = False
            continue
        esperar_sub = False
        if ruido:
            continue
        m = RX_SOLICITUD.match(l)
        if m:
            if actual:
                regs.append(actual)
            actual = {"solicitud": m.group(1), "seccion": seccion, "texto": l}
        elif actual:
            actual["texto"] += " " + l
    if actual:
        regs.append(actual)
    return regs


# Donde empiezan las observaciones de INAPI (texto legal que no es parte de la marca)
RX_FIN_MARCA = re.compile(
    r"\s(?:Vistos?[,:]|Que,|(?:Atendid[oa]s?|Habi[eé]ndose|En particular|CONSIDERANDO|Considerando|Por cumplido|"
    r"Con protecci[oó]n|Sin protecci[oó]n|No se ha|N[uú]mero de Registro|Se (?:observa|solicita|rechaza|acoge|"
    r"tiene|deja|ordena|concede)|Respecto|Conforme a|De acuerdo|Por resoluci[oó]n|Teniendo presente|"
    r"El solicitante|La solicitud|Notif[ií]quese|Resoluci[oó]n|Atendiendo|Previo a|Se le)\b)")


def extraer_marca(texto):
    """Nombre de la marca: lo que viene después del tipo de signo (Mixta, Denominativa...),
    sin las observaciones legales que INAPI agrega a continuación."""
    m = RX_TIPO_SIGNO.search(texto)
    if not m:
        return ""
    marca = " " + re.sub(r"\s+", " ", texto[m.end():]).strip()
    fin = RX_FIN_MARCA.search(marca)
    if fin:
        marca = marca[:fin.start()]
    palabras = marca.split()
    return " ".join(palabras[:14])


def registros_de(texto):
    """Devuelve (registros, cantidad_de_marcas)."""
    regs = separar_registros(texto)
    if len(regs) >= 3:
        return regs, len(regs)
    # Formato desconocido: se revisa de a dos líneas seguidas
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    salida, numeros = [], set()
    for i in range(len(lineas)):
        t = " ".join(lineas[i:i + 2])
        m = re.search(r"\b\d{6,8}\b", lineas[i])
        if m:
            numeros.add(m.group(0))
        salida.append({"solicitud": m.group(0) if m else "", "seccion": "", "texto": t, "ventana": True})
    return salida, (len(numeros) or len(lineas))


def analizar(doc, texto, claves):
    """Devuelve (coincidencias, cantidad_de_marcas_revisadas, solicitudes)."""
    regs, cantidad = registros_de(texto)
    solicitudes = {r["solicitud"] for r in regs if r["solicitud"]}
    resultados, anterior = [], None
    for reg in regs:
        marca = extraer_marca(reg["texto"])
        hall = buscar(reg["texto"], claves, marca)
        if not hall:
            anterior = None
            continue
        palabras = ", ".join(h[0] for h in hall)
        if reg.get("ventana") and anterior == palabras:
            continue  # misma coincidencia que la ventana anterior
        anterior = palabras
        orden = {"exacta": 0, "parecida": 1, "palabra": 2}
        nivel = min((h[1] for h in hall), key=orden.get)
        resultados.append({
            "fecha": doc["fecha"],
            "documento": TIPOS.get(doc["tipo"], doc["tipo"]),
            "seccion": reg["seccion"],
            "solicitud": reg["solicitud"],
            "marca": marca,
            "palabras": palabras,
            "nivel": nivel,
            "tipo": "; ".join(f"{h[0]}: {h[2]}" for h in hall),
            "texto": re.sub(r"\s+", " ", reg["texto"]).strip(),
            "url": doc["url"],
            "archivo": doc.get("archivo", ""),
            "claves": ["|".join([doc["stream_id"], reg["solicitud"] or normalizar(reg["texto"])[:80], h[0]])
                       for h in hall],
        })
    return resultados, cantidad, solicitudes


# ---------------------------------------------------------------- INAPI
# ojo: INAPI pone la fecha con espacios dentro del enlace ("04/11/2026 12:00:00 a. m."),
# así que el enlace termina en la comilla, no en el primer espacio
RX_ENLACE = re.compile(r"(?:https?://tramites\.inapi\.cl)?/Trademark/TrademarkDailyStatus/DownloadFile\?[^\"'<>\n]+",
                       re.I)
DIAS_A_REVISAR = 30   # la página trae todo el historial; solo se revisan los últimos días
UA_EDGE = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
           "Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0")
SIN_VENTANA = 0x08000000 if os.name == "nt" else 0


def extraer_documentos(texto):
    """Busca los enlaces DownloadFile (en HTML o en JSON) y arma la lista de documentos."""
    pagina = html.unescape(texto or "").replace("\\u0026", "&").replace("\\/", "/")
    docs = {}
    limite = (datetime.date.today() - datetime.timedelta(days=DIAS_A_REVISAR)).isoformat()
    for m in RX_ENLACE.finditer(pagina):
        url = urljoin(URL_PAGINA, m.group(0).strip().replace(" ", "%20"))
        q = parse_qs(urlparse(url).query)
        tipo = q.get("dailyStatesTypeId", [""])[0]
        sid = q.get("stream_id", [""])[0]
        fd = q.get("filedate", [""])[0]
        if tipo not in TIPOS or not sid:
            continue
        try:
            fecha = datetime.datetime.strptime(fd[:10], "%d/%m/%Y").date().isoformat()
        except ValueError:
            fecha = fd[:10]
        if len(fecha) == 10 and fecha < limite:
            continue
        docs[sid] = {"url": url, "tipo": tipo, "fecha": fecha, "stream_id": sid}
    return sorted(docs.values(), key=lambda d: (d["fecha"], d["tipo"]))


def parece_documento(datos):
    return datos[:4] == b"%PDF" or datos[:2] == b"PK" or b"<html" not in datos[:3000].lower()


def guardar_texto(ruta, texto):
    try:
        DATOS.mkdir(parents=True, exist_ok=True)
        Path(ruta).write_text(texto or "", encoding="utf-8")
    except Exception:
        pass


def instalar_chromium():
    import subprocess
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], timeout=1200,
                   creationflags=SIN_VENTANA, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


class FuenteINAPI:
    """Lee INAPI directo; si la página arma el listado con JavaScript, usa un navegador invisible (Edge)."""

    def __init__(self, requests):
        self.requests = requests
        self.sesion = requests.Session()
        self.sesion.headers.update({"User-Agent": NAVEGADOR, "Accept-Language": "es-CL,es;q=0.9"})
        self.pw = self.nav = self.ctx = None
        self.metodo = "directo"
        self.notas = []

    def listar(self):
        r = self.sesion.get(URL_PAGINA, timeout=60)  # los errores de conexión suben (para reintentar)
        docs = extraer_documentos(r.text) if r.ok else []
        if docs:
            self.notas.append(f"Vía directa OK: {len(docs)} documentos")
            return docs
        self.notas.append(f"Vía directa sin enlaces: HTTP {r.status_code}, {len(r.text)} caracteres")
        guardar_texto(DATOS / "ultima_pagina_inapi.html", r.text)
        log("La página de INAPI arma el listado con JavaScript: la leo con el navegador (invisible)...")
        contenido, extras = self._leer_con_navegador()
        guardar_texto(DATOS / "ultima_pagina_inapi_navegador.html", contenido)
        docs = extraer_documentos(contenido + "\n" + "\n".join(extras))
        if not docs:
            raise RuntimeError("No encontré los enlaces de los Estados Diarios en la página de INAPI, ni "
                               "leyéndola directo ni con el navegador. Puede que el sitio haya cambiado.")
        self.metodo = "navegador"
        self.notas.append(f"Vía navegador OK: {len(docs)} documentos")
        return docs

    def _abrir_navegador(self):
        if self.ctx is not None:
            return
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise RuntimeError("Falta el componente 'playwright' para leer la página con el navegador. "
                               "Vuelve a ejecutar el instalador.")
        self.pw = sync_playwright().start()
        for canal in ("msedge", "chrome", None):
            try:
                self.nav = (self.pw.chromium.launch(channel=canal, headless=True) if canal
                            else self.pw.chromium.launch(headless=True))
                self.notas.append(f"Navegador usado: {canal or 'chromium descargado'}")
                break
            except Exception as e:
                self.notas.append(f"No se pudo abrir {canal or 'chromium'}: {str(e).splitlines()[0][:200]}")
        if self.nav is None:
            log("No encontré Edge ni Chrome: descargo un navegador (solo la primera vez, unos minutos)...")
            instalar_chromium()
            self.nav = self.pw.chromium.launch(headless=True)
            self.notas.append("Navegador usado: chromium recién descargado")
        self.ctx = self.nav.new_context(user_agent=UA_EDGE, locale="es-CL")

    def _leer_con_navegador(self):
        self._abrir_navegador()
        pagina = self.ctx.new_page()
        respuestas = []
        pagina.on("response", lambda resp: respuestas.append(resp)
                  if resp.request.resource_type in ("xhr", "fetch") else None)
        pagina.goto(URL_PAGINA, wait_until="domcontentloaded", timeout=90000)
        try:
            pagina.wait_for_selector("a[href*='DownloadFile']", timeout=45000)
        except Exception:
            self.notas.append("El navegador no vio enlaces DownloadFile tras 45 s")
            try:
                pagina.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass
        contenido = pagina.content()
        extras = []
        carpeta = DATOS / "respuestas_inapi"
        carpeta.mkdir(parents=True, exist_ok=True)
        for f in carpeta.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
        for i, resp in enumerate(respuestas[:40]):
            try:
                t = resp.text()
                extras.append(t)
                self.notas.append(f"Petición {resp.status} {resp.url[:160]} ({len(t)} car.)")
                guardar_texto(carpeta / f"respuesta_{i:02d}.txt", resp.url + "\n\n" + t[:200000])
            except Exception:
                pass
        pagina.close()
        return contenido, extras

    def descargar(self, url):
        ultimo = None
        for intento in range(3):
            try:
                if self.metodo == "directo":
                    r = self.sesion.get(url, timeout=180)
                    if r.ok and parece_documento(r.content):
                        return r.content
                    self.notas.append(f"Descarga directa no válida (HTTP {r.status_code}): cambio al navegador")
                    self.metodo = "navegador"
                self._abrir_navegador()
                r = self.ctx.request.get(url, timeout=180000)
                datos = r.body()
                if r.ok and parece_documento(datos):
                    return datos
                raise RuntimeError(f"INAPI no entregó el archivo (HTTP {r.status})")
            except Exception as e:
                ultimo = e
                time.sleep(5 * (intento + 1))
        raise ultimo

    def cerrar(self):
        for accion in (lambda: self.ctx and self.ctx.close(), lambda: self.nav and self.nav.close(),
                       lambda: self.pw and self.pw.stop()):
            try:
                accion()
            except Exception:
                pass
        guardar_texto(DATOS / "notas_conexion.txt",
                      f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n" + "\n".join(self.notas))


# ---------------------------------------------------------------- diagnóstico
def mostrar_en_explorador(ruta):
    try:
        import subprocess
        subprocess.Popen(f'explorer /select,"{ruta}"')
    except Exception:
        abrir(Path(ruta).parent)


def generar_diagnostico(error_texto=""):
    """Arma un ZIP con todo lo necesario para corregir el programa."""
    import platform
    from importlib import metadata
    carpeta = CARPETA / "diagnostico"
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f"Diagnostico_Vigilancia_INAPI_{datetime.datetime.now():%Y-%m-%d_%H%M}.zip"

    def version(p):
        try:
            return metadata.version(p)
        except Exception:
            return "no instalado"

    conf = cargar_config()
    if conf["correo"].get("clave"):
        conf["correo"]["clave"] = "(guardada, oculta)"
    est = cargar_estado()
    resumen_est = {k: est.get(k) for k in ("ultima_revision", "fallos_seguidos", "ultimo_resumen", "acumulado")}
    resumen_est["documentos_procesados"] = len(est.get("procesados", {}))
    try:
        marcas = leer_lista_claves()
    except Exception as e:
        marcas = [f"(error al leer: {e})"]
    lineas = [
        "INFORME DE DIAGNÓSTICO - Vigilancia de marcas INAPI",
        f"Generado: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Versión del programa: {VERSION}",
        "",
        "== Error ==", error_texto or "(sin error indicado)",
        "",
        "== Equipo ==",
        f"Sistema: {platform.platform()}  ({platform.machine()})",
        f"Python: {sys.version}",
        f"Ejecutable: {sys.executable}",
        f"Carpeta del programa: {CARPETA}",
        f"Revisión al encender activada: {inicio_activado()}",
        "",
        "== Componentes ==",
    ] + [f"{p}: {version(p)}" for p in ("requests", "pypdf", "openpyxl", "truststore", "playwright")] + [
        "",
        "== Marcas vigiladas ==", ", ".join(marcas),
        "",
        "== Ajustes ==", json.dumps(conf, ensure_ascii=False, indent=1),
        "",
        "== Estado ==", json.dumps(resumen_est, ensure_ascii=False, indent=1),
    ]
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("diagnostico.txt", "\n".join(lineas))
        if REGISTRO.exists():
            z.writestr("registro.txt", "\n".join(REGISTRO.read_text(encoding="utf-8", errors="replace")
                                                  .splitlines()[-500:]))
        for nombre in ("instalacion.log", "ultima_pagina_inapi.html", "ultima_pagina_inapi_navegador.html",
                       "notas_conexion.txt"):
            f = DATOS / nombre
            if f.exists():
                z.write(f, nombre)
        for f in sorted((DATOS / "respuestas_inapi").glob("*.txt"))[:40]:
            z.write(f, "respuestas_inapi/" + f.name)
    for viejo in sorted(carpeta.glob("Diagnostico_*.zip"))[:-5]:
        try:
            viejo.unlink()
        except Exception:
            pass
    log(f"Diagnóstico generado: {ruta}")
    return ruta


# ---------------------------------------------------------------- informe
COLOR_NIVEL = {"exacta": ("#fde2e1", "#b42318"), "parecida": ("#fff1d6", "#a15c07"),
               "palabra": ("#e6efff", "#2457d6")}


def fmt(n):
    return f"{n:,}".replace(",", ".")


def construir_informe(r):
    """r: dict con el resumen de la revisión. Devuelve el HTML del informe."""
    e = html.escape
    nuevas = r["nuevas"]
    por_nivel = {k: sum(1 for c in nuevas if c["nivel"] == k) for k in NIVELES}

    def tarjeta(valor, titulo, nota=""):
        return (f"<td class='kpi'><div class='v'>{valor}</div><div class='t'>{e(titulo)}</div>"
                f"<div class='nt'>{e(nota)}</div></td>")

    filas_c = []
    for c in nuevas:
        fondo, color = COLOR_NIVEL[c["nivel"]]
        enlace_local = Path(c["archivo"]).as_uri() if c.get("archivo") else ""
        filas_c.append(
            "<tr>"
            f"<td><span class='chip' style='background:{fondo};color:{color}'>{e(NIVELES[c['nivel']])}</span></td>"
            f"<td class='n'>{e(c['solicitud'] or '—')}</td>"
            f"<td><b>{e(c['marca'] or '(sin nombre / figurativa)')}</b><br>"
            f"<span class='s'>{e(c['texto'][:260])}</span></td>"
            f"<td>{e(c['tipo'])}</td>"
            f"<td>{e(c['fecha'])}<br><span class='s'>{e(c['documento'])}<br>{e(c['seccion'])}</span></td>"
            f"<td><a href='{e(c['url'])}'>INAPI</a>"
            + (f"<br><a href='{e(enlace_local)}'>PDF local</a>" if enlace_local else "") +
            "</td></tr>")
    tabla_c = ("<table><tr><th>Nivel</th><th>N° solicitud</th><th>Marca</th><th>Coincide con</th>"
               "<th>Publicado</th><th>Ver</th></tr>" + "".join(filas_c) + "</table>") if nuevas else \
        "<p class='ok'>✓ No se encontraron coincidencias nuevas en esta revisión.</p>"

    filas_d = "".join(
        f"<tr><td>{e(d['fecha'])}</td><td>{e(TIPOS.get(d['tipo'], d['tipo']))}</td>"
        f"<td class='r'>{fmt(d['marcas'])}</td><td class='r'>{d['coincidencias']}</td>"
        f"<td>{e(d.get('error', ''))}</td></tr>" for d in r["documentos"])
    tabla_d = ("<table><tr><th>Fecha</th><th>Documento</th><th class='r'>Marcas revisadas</th>"
               "<th class='r'>Coincidencias</th><th>Observación</th></tr>" + filas_d + "</table>") \
        if r["documentos"] else "<p>No había documentos nuevos de INAPI desde la revisión anterior.</p>"

    return f"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<title>Informe de vigilancia INAPI · {e(r['cuando'])}</title>
<style>
 body{{font-family:Segoe UI,Arial,sans-serif;margin:0;padding:28px;color:#1d2433;background:#f4f6fb}}
 .caja{{max-width:1100px;margin:auto}} h1{{font-size:22px;margin:0}} h2{{font-size:16px;margin:28px 0 10px}}
 .sub{{color:#5b6475;margin:4px 0 18px}} table{{border-collapse:collapse;width:100%;background:#fff}}
 th,td{{padding:9px 11px;border-bottom:1px solid #e5e7ef;vertical-align:top;font-size:13px;text-align:left}}
 th{{background:#1d2433;color:#fff;font-weight:600}} .s{{color:#6b7280;font-size:12px}}
 .n{{font-family:Consolas,monospace;white-space:nowrap}} .r{{text-align:right}} a{{color:#2457d6;white-space:nowrap}}
 .kpis{{border-collapse:separate;border-spacing:10px 0;margin:0 -10px;background:none}}
 .kpi{{background:#fff;border:none;padding:16px;width:25%}} .kpi .v{{font-size:28px;font-weight:700}}
 .kpi .t{{font-size:13px;color:#1d2433}} .kpi .nt{{font-size:11px;color:#6b7280}}
 .chip{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap}}
 .ok{{background:#e8f6ee;color:#16794a;padding:12px 14px}} .pie{{color:#6b7280;font-size:12px;margin-top:24px}}
</style></head><body><div class="caja">
<h1>Informe de vigilancia de marcas INAPI</h1>
<div class="sub">Revisión del {e(r['cuando'])} · Documentos de INAPI disponibles: {e(r['rango'])}</div>
<table class="kpis"><tr>
{tarjeta(fmt(r['marcas']), 'Marcas revisadas', 'registros leídos en los documentos')}
{tarjeta(fmt(r['solicitudes']), 'Solicitudes distintas', 'una solicitud puede aparecer en varios días')}
{tarjeta(len(r['documentos']), 'Documentos revisados', 'Presentaciones + Estado diario')}
{tarjeta(len(nuevas), 'Coincidencias nuevas',
         f"{por_nivel['exacta']} exactas · {por_nivel['parecida']} parecidas · {por_nivel['palabra']} por una palabra")}
</tr></table>
<h2>Coincidencias nuevas</h2>
{tabla_c}
<h2>Marcas vigiladas</h2>
<p>{e(', '.join(r['vigiladas']))}{' · También avisa si comparte una palabra (se ignoran: ' + e(r['ignoradas']) + ')' if r['una_palabra'] else ''}</p>
<h2>Documentos revisados</h2>
{tabla_d}
<p class="pie">Acumulado desde la instalación: {fmt(r['acum_marcas'])} marcas revisadas en {r['acum_docs']} documentos ·
{r['acum_coinc']} coincidencias registradas en historial_coincidencias.csv.<br>
Para ver el detalle de una solicitud, búscala por su número en <a href="https://tramites.inapi.cl/">tramites.inapi.cl</a>.
Si una marca ajena se parece a la tuya, revisa los plazos de oposición con tu abogado de marcas.</p>
</div></body></html>"""


def guardar_informe(contenido):
    REPORTES.mkdir(parents=True, exist_ok=True)
    ruta = REPORTES / f"informe_{datetime.datetime.now():%Y-%m-%d_%H%M%S}.html"
    ruta.write_text(contenido, encoding="utf-8")
    ULTIMO_INFORME.write_text(contenido, encoding="utf-8")
    viejos = sorted(REPORTES.glob("informe_*.html"))[:-INFORMES_A_GUARDAR]
    for v in viejos:
        try:
            v.unlink()
        except Exception:
            pass
    return ruta


ENCABEZADO_HIST = ["Detectado el", "Fecha estado diario", "Documento", "Sección", "N° solicitud", "Marca",
                   "Nivel", "Coincide con", "Tipo de coincidencia", "Texto del registro", "Enlace INAPI"]


def agregar_historial(nuevas):
    filas_previas = []
    if HISTORIAL.exists():
        with open(HISTORIAL, encoding="utf-8-sig", newline="") as f:
            lector = csv.DictReader(f, delimiter=";")
            if lector.fieldnames != ENCABEZADO_HIST:  # versión anterior: se migra
                filas_previas = list(lector)
        if filas_previas:
            HISTORIAL.unlink()
    nuevo = not HISTORIAL.exists()
    with open(HISTORIAL, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if nuevo:
            w.writerow(ENCABEZADO_HIST)
        for p in filas_previas:
            w.writerow([p.get(k, "") for k in ENCABEZADO_HIST])
        hoy = f"{datetime.datetime.now():%Y-%m-%d %H:%M}"
        for c in nuevas:
            w.writerow([hoy, c["fecha"], c["documento"], c["seccion"], c["solicitud"], c["marca"],
                        NIVELES[c["nivel"]], c["palabras"], c["tipo"], c["texto"], c["url"]])


def agregar_revision(r):
    nuevo = not HIST_REVISIONES.exists()
    with open(HIST_REVISIONES, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if nuevo:
            w.writerow(["Fecha revisión", "Documentos revisados", "Marcas revisadas", "Solicitudes distintas",
                        "Coincidencias nuevas", "Correo enviado"])
        w.writerow([r["cuando"], len(r["documentos"]), r["marcas"], r["solicitudes"], len(r["nuevas"]),
                    r.get("correo", "")])


def limpiar_antiguos(est):
    limite = time.time() - DIAS_GUARDAR_ARCHIVOS * 86400
    for f in ARCHIVOS.glob("*"):
        try:
            if f.stat().st_mtime < limite:
                f.unlink()
        except Exception:
            pass
    corte = (datetime.date.today() - datetime.timedelta(days=90)).isoformat()
    est["procesados"] = {k: v for k, v in est.get("procesados", {}).items() if v.get("fecha", "9") >= corte}


def texto_plano(r):
    lineas = [f"Vigilancia de marcas INAPI - revisión del {r['cuando']}",
              f"Marcas revisadas: {fmt(r['marcas'])} en {len(r['documentos'])} documentos "
              f"({fmt(r['solicitudes'])} solicitudes distintas).",
              f"Coincidencias nuevas: {len(r['nuevas'])}", ""]
    for c in r["nuevas"]:
        lineas.append(f"- [{NIVELES[c['nivel']]}] N° {c['solicitud'] or '—'} · {c['marca'] or '(sin nombre)'} · "
                      f"{c['tipo']} · {c['fecha']} · {c['url']}")
    lineas.append("\nEl informe completo va adjunto.")
    return "\n".join(lineas)


# ---------------------------------------------------------------- principal
def revisar(manual, inicio=False):
    preparar_ssl()
    import requests
    fuente = FuenteINAPI(requests)
    try:
        _revisar(fuente, requests, manual, inicio)
    finally:
        fuente.cerrar()


def _revisar(fuente, requests, manual, inicio):
    conf = cargar_config()
    claves = cargar_claves(conf)
    if not claves:
        aviso("Vigilancia de marcas INAPI", "No hay marcas para vigilar. Agrégalas en el panel.")
        return
    est = cargar_estado()
    huella = huella_busqueda(claves, conf)
    rehacer = est.get("huella_claves") != huella
    procesados = {} if rehacer else est.get("procesados", {})
    vistos = set(est.get("vistos", []))

    log("Leyendo la página de Estados Diarios de INAPI...")
    # Al encender el PC la conexión puede tardar: se reintenta hasta ~10 minutos
    intentos = 8 if inicio else 1
    for n in range(intentos):
        try:
            docs = fuente.listar()
            break
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
            if n == intentos - 1:
                raise RuntimeError("No hay conexión con INAPI (revisa tu internet).") from e
            log("   Sin conexión todavía, reintento en 75 segundos...")
            time.sleep(75)
    pendientes = [d for d in docs if d["stream_id"] not in procesados]
    log(f"{len(docs)} documentos disponibles, {len(pendientes)} por revisar. "
        f"Marcas vigiladas: {', '.join(c.texto for c in claves)}")

    ARCHIVOS.mkdir(parents=True, exist_ok=True)
    nuevas, revisados, solicitudes, total_marcas = [], [], set(), 0
    for i, d in enumerate(pendientes, 1):
        log(f"[{i}/{len(pendientes)}] {d['fecha']} · {TIPOS[d['tipo']]}")
        try:
            datos = fuente.descargar(d["url"])
            nombre = f"{d['fecha']}_{'presentaciones' if d['tipo'] == '5' else 'estado_diario'}_{d['stream_id'][:8]}"
            ruta = ARCHIVOS / (nombre + extension(datos))
            ruta.write_bytes(datos)
            d["archivo"] = str(ruta)
            texto = extraer_texto(datos)
            # muestra del texto leído (sirve para diagnosticar cambios de formato de INAPI)
            guardar_texto(DATOS / f"muestra_{'presentaciones' if d['tipo'] == '5' else 'estado_diario'}.txt",
                          "\n".join(texto.splitlines()[:200]))
            encontrados, cantidad, sols = analizar(d, texto, claves)
        except Exception as e:
            log(f"   No pude revisar este documento: {e}")
            revisados.append({**d, "marcas": 0, "coincidencias": 0, "error": f"No se pudo leer: {e}"[:120]})
            continue
        for c in encontrados:
            if any(k not in vistos for k in c["claves"]):
                vistos.update(c["claves"])
                nuevas.append(c)
        total_marcas += cantidad
        solicitudes |= sols
        revisados.append({**d, "marcas": cantidad, "coincidencias": len(encontrados)})
        log(f"   {fmt(cantidad)} marcas revisadas, {len(encontrados)} coincidencia(s)")
        procesados[d["stream_id"]] = {"fecha": d["fecha"], "tipo": d["tipo"], "marcas": cantidad,
                                      "coincidencias": len(encontrados)}
        est.update({"procesados": procesados, "vistos": sorted(vistos), "huella_claves": huella})
        guardar_estado(est)
        time.sleep(1)

    ahora = f"{datetime.datetime.now():%Y-%m-%d %H:%M}"
    acum = est.get("acumulado", {"marcas": 0, "docs": 0, "coincidencias": 0})
    acum["marcas"] += total_marcas
    acum["docs"] += sum(1 for d in revisados if not d.get("error"))
    acum["coincidencias"] += len(nuevas)
    fechas = sorted({d["fecha"] for d in docs})
    resumen = {
        "cuando": ahora, "rango": f"{fechas[0]} al {fechas[-1]}" if fechas else "—",
        "marcas": total_marcas, "solicitudes": len(solicitudes), "documentos": revisados, "nuevas": nuevas,
        "vigiladas": [c.texto for c in claves], "una_palabra": conf.get("una_palabra", True),
        "ignoradas": conf.get("palabras_ignoradas", ""),
        "acum_marcas": acum["marcas"], "acum_docs": acum["docs"], "acum_coinc": acum["coincidencias"],
    }
    est.update({"procesados": procesados, "vistos": sorted(vistos), "huella_claves": huella,
                "ultima_revision": ahora, "fallos_seguidos": 0, "acumulado": acum})
    if revisados:
        est["ultimo_resumen"] = {"cuando": ahora, "documentos": len(revisados), "marcas": total_marcas,
                                 "solicitudes": len(solicitudes), "nuevas": len(nuevas)}
    limpiar_antiguos(est)
    guardar_estado(est)

    informe = None
    if revisados:
        informe = guardar_informe(construir_informe(resumen))
        log(f"Informe: {informe}")
    if nuevas:
        agregar_historial(nuevas)

    # correo
    correo = conf["correo"]
    estado_correo = ""
    if correo.get("activo") and informe and (nuevas or correo.get("resumen_sin_novedades")):
        if nuevas:
            exactas = sum(1 for c in nuevas if c["nivel"] == "exacta")
            asunto = (f"⚠ INAPI: {len(nuevas)} marca(s) en conflicto o similares"
                      + (f" ({exactas} exacta/s)" if exactas else ""))
        else:
            asunto = f"INAPI: sin coincidencias nuevas ({fmt(total_marcas)} marcas revisadas)"
        try:
            enviar_correo(correo, asunto, informe.read_text(encoding="utf-8"), texto_plano(resumen), [informe])
            estado_correo = "Sí"
            log(f"Correo enviado a {', '.join(destinatarios_de(correo))}")
        except Exception as e:
            estado_correo = "Error"
            log(f"No pude enviar el correo: {e}")
            aviso("Vigilancia de marcas INAPI", "No pude enviar el correo de aviso:\n\n" + explicar_error_correo(e)
                  + "\n\nRevisa los datos en el panel → «Ajustes y correo».")
    if revisados:
        resumen["correo"] = estado_correo
        agregar_revision(resumen)

    # ventana de aviso
    cifras = (f"{fmt(total_marcas)} marcas revisadas en {len(revisados)} documento(s)." if revisados else
              "No había documentos nuevos de INAPI desde la revisión anterior.")
    if nuevas:
        lista = "\n".join(f"• [{NIVELES[c['nivel']]}] N° {c['solicitud'] or '—'}  {c['marca'][:40] or '(sin nombre)'}"
                          f"  ← {c['palabras']}" for c in nuevas[:8])
        mas = f"\n…y {len(nuevas) - 8} más." if len(nuevas) > 8 else ""
        extra = "\nTe envié el aviso por correo." if estado_correo == "Sí" else ""
        if aviso("Vigilancia de marcas INAPI",
                 f"Encontré {len(nuevas)} coincidencia(s) nueva(s) en INAPI.\n{cifras}\n\n{lista}{mas}\n{extra}\n"
                 f"¿Abrir el informe?", preguntar=True):
            abrir(informe)
    else:
        log("Sin coincidencias nuevas.")
        texto = (f"Revisión lista: no hay coincidencias nuevas.\n{cifras}\n\n"
                 f"Documentos de INAPI disponibles: {resumen['rango']}\n"
                 f"Marcas vigiladas: {', '.join(c.texto for c in claves)}")
        if manual:
            mostrar = informe or (ULTIMO_INFORME if ULTIMO_INFORME.exists() else None)
            if mostrar and aviso("Vigilancia de marcas INAPI", texto + "\n\n¿Abrir el " +
                                 ("informe?" if informe else "último informe?"), preguntar=True):
                abrir(mostrar)
            elif not mostrar:
                aviso("Vigilancia de marcas INAPI", texto)
        elif conf.get("avisar_sin_novedades") and revisados:
            aviso("Vigilancia de marcas INAPI", texto)


# ---------------------------------------------------------------- panel (ventana)
AZUL_OSCURO, AZUL, FONDO, BLANCO, GRIS = "#1d2433", "#2457d6", "#f4f6fb", "#ffffff", "#6b7280"


class Panel:
    def __init__(self):
        import tkinter as tk
        from tkinter import ttk, messagebox
        self.tk, self.ttk, self.mb = tk, ttk, messagebox
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # letras nítidas en pantallas con zoom
        except Exception:
            pass
        self.proceso = None
        self.enlaces = {}

        r = self.raiz = tk.Tk()
        r.title(f"Vigilancia de marcas INAPI  ·  versión {VERSION}")
        r.geometry("1080x660")
        r.minsize(900, 560)
        r.configure(bg=FONDO)
        try:
            r.iconbitmap(str(Path(os.environ.get("LOCALAPPDATA", "")) / "VigilanciaMarcasINAPI" / "icono.ico"))
        except Exception:
            pass

        st = ttk.Style(r)
        st.theme_use("clam")
        base = ("Segoe UI", 10)
        st.configure(".", font=base, background=FONDO)
        st.configure("Tarjeta.TFrame", background=BLANCO)
        st.configure("Tarjeta.TLabel", background=BLANCO)
        st.configure("Tarjeta.TCheckbutton", background=BLANCO)
        st.configure("Titulo.TLabel", background=BLANCO, font=("Segoe UI Semibold", 12))
        st.configure("Nota.TLabel", background=BLANCO, foreground=GRIS, font=("Segoe UI", 9))
        st.configure("Pie.TLabel", background=FONDO, foreground=GRIS)
        st.configure("TButton", padding=(12, 6))
        st.configure("Principal.TButton", background=AZUL, foreground=BLANCO, font=("Segoe UI Semibold", 10),
                     padding=(18, 8), borderwidth=0)
        st.map("Principal.TButton", background=[("disabled", "#9db3e8"), ("active", "#1b46b3")],
               foreground=[("disabled", BLANCO)])
        st.configure("TCheckbutton", background=FONDO)
        st.configure("Treeview", rowheight=26, font=base, background=BLANCO, fieldbackground=BLANCO)
        st.configure("Treeview.Heading", font=("Segoe UI Semibold", 10), background="#e8ebf3")

        # encabezado
        cab = tk.Frame(r, bg=AZUL_OSCURO)
        cab.pack(fill="x")
        tk.Label(cab, text="Vigilancia de marcas INAPI", bg=AZUL_OSCURO, fg=BLANCO,
                 font=("Segoe UI Semibold", 17)).pack(anchor="w", padx=24, pady=(14, 0))
        self.lbl_estado = tk.Label(cab, text="", bg=AZUL_OSCURO, fg="#c7cde0", font=("Segoe UI", 10),
                                   justify="left")
        self.lbl_estado.pack(anchor="w", padx=24, pady=(2, 12))

        cuerpo = ttk.Frame(r)
        cuerpo.pack(fill="both", expand=True, padx=20, pady=14)
        cuerpo.columnconfigure(1, weight=1)
        cuerpo.rowconfigure(0, weight=1)

        # tarjeta izquierda: marcas
        izq = ttk.Frame(cuerpo, style="Tarjeta.TFrame", padding=16)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        ttk.Label(izq, text="Marcas vigiladas", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(izq, text="No importan mayúsculas ni tildes.", style="Nota.TLabel").pack(anchor="w", pady=(0, 8))
        self.lista = tk.Listbox(izq, font=("Segoe UI", 11), width=26, height=10, activestyle="none",
                                relief="flat", highlightthickness=1, highlightcolor=AZUL,
                                highlightbackground="#d7dbe7", selectbackground=AZUL, selectforeground=BLANCO)
        self.lista.pack(fill="both", expand=True)
        self.lista.bind("<Delete>", lambda e: self.quitar())
        fila = ttk.Frame(izq, style="Tarjeta.TFrame")
        fila.pack(fill="x", pady=(10, 4))
        self.entrada = ttk.Entry(fila, font=("Segoe UI", 11))
        self.entrada.pack(side="left", fill="x", expand=True, ipady=3)
        self.entrada.bind("<Return>", lambda e: self.agregar())
        ttk.Button(fila, text="Agregar", command=self.agregar).pack(side="left", padx=(8, 0))
        ttk.Button(izq, text="Quitar la marca seleccionada", command=self.quitar).pack(fill="x", pady=(4, 8))
        ttk.Label(izq, text="Avisa si una marca es igual, parecida\n(7+ letras) o comparte una palabra.",
                  style="Nota.TLabel", justify="left").pack(anchor="w")

        # tarjeta derecha: coincidencias
        der = ttk.Frame(cuerpo, style="Tarjeta.TFrame", padding=16)
        der.grid(row=0, column=1, sticky="nsew")
        der.rowconfigure(2, weight=1)
        der.columnconfigure(0, weight=1)
        ttk.Label(der, text="Coincidencias encontradas", style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        self.lbl_total = ttk.Label(der, text="", style="Nota.TLabel")
        self.lbl_total.grid(row=1, column=0, sticky="w", pady=(0, 8))
        cols = ("nivel", "fecha", "solicitud", "marca", "coincide")
        self.tabla = ttk.Treeview(der, columns=cols, show="headings", selectmode="browse")
        for c, t, w, s in (("nivel", "Nivel", 178, False), ("fecha", "Fecha", 90, False),
                           ("solicitud", "N° solicitud", 95, False), ("marca", "Marca", 200, True),
                           ("coincide", "Coincide con", 150, True)):
            self.tabla.heading(c, text=t, anchor="w")
            self.tabla.column(c, width=w, minwidth=60, stretch=s, anchor="w")
        self.tabla.tag_configure("Exacta", foreground="#b42318")
        self.tabla.tag_configure("Parecida", foreground="#a15c07")
        barra = ttk.Scrollbar(der, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.grid(row=2, column=0, sticky="nsew")
        barra.grid(row=2, column=1, sticky="ns")
        self.tabla.bind("<Double-1>", self.abrir_fila)
        bot = ttk.Frame(der, style="Tarjeta.TFrame")
        bot.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(bot, text="Último informe", command=lambda: self.abrir_si_existe(ULTIMO_INFORME)
                   ).pack(side="right")
        ttk.Button(bot, text="Historial (Excel)", command=lambda: self.abrir_si_existe(HISTORIAL)
                   ).pack(side="right", padx=8)
        ttk.Label(bot, text="Doble clic en una fila para abrir el documento de INAPI.",
                  style="Nota.TLabel").pack(side="left", fill="x", expand=True)

        # pie
        pie = ttk.Frame(r)
        pie.pack(fill="x", padx=20, pady=(0, 8))
        self.btn_revisar = ttk.Button(pie, text="Revisar ahora", style="Principal.TButton", command=self.revisar_ahora)
        self.btn_revisar.pack(side="left")
        self.var_inicio = tk.BooleanVar(value=inicio_activado())
        ttk.Checkbutton(pie, text="Revisar al encender el computador", variable=self.var_inicio,
                        command=self.cambiar_inicio).pack(side="left", padx=(18, 0))
        ttk.Button(pie, text="Abrir carpeta", command=lambda: abrir(CARPETA)).pack(side="right")
        ttk.Button(pie, text="Ajustes y correo…", command=self.abrir_ajustes).pack(side="right", padx=8)
        self.lbl_pie = ttk.Label(r, text="", style="Pie.TLabel")
        self.lbl_pie.pack(fill="x", padx=22, pady=(0, 10))

        self.refrescar_todo()
        self.entrada.focus_set()

    # --- datos
    def refrescar_todo(self):
        self.refrescar_marcas()
        self.refrescar_coincidencias()
        est = cargar_estado()
        conf = cargar_config()
        u = est.get("ultimo_resumen")
        acum = est.get("acumulado", {})
        if u:
            linea1 = (f"Última revisión con documentos nuevos: {u['cuando']}  ·  {fmt(u['marcas'])} marcas revisadas"
                      f" en {u['documentos']} documento(s)  ·  {u['nuevas']} coincidencia(s) nueva(s)")
            if est.get("ultima_revision") and est["ultima_revision"] != u["cuando"]:
                linea1 += f"  ·  Última comprobación: {est['ultima_revision']}"
        elif est.get("ultima_revision"):
            linea1 = f"Última revisión: {est['ultima_revision']}"
        else:
            linea1 = "Aún no se ha hecho ninguna revisión"
        auto = "al encender el PC y " if inicio_activado() else ""
        mail = ("correo a " + ", ".join(destinatarios_de(conf["correo"]))) if conf["correo"].get("activo") \
            else "correo desactivado"
        linea2 = (f"Total: {fmt(acum.get('marcas', 0))} marcas revisadas  ·  Revisión automática: {auto}"
                  f"todos los días a las 11:00  ·  Avisos: ventana + {mail}")
        self.lbl_estado.config(text=linea1 + "\n" + linea2)

    def refrescar_marcas(self, seleccionar=None):
        self.marcas = leer_lista_claves()
        self.lista.delete(0, "end")
        for m in self.marcas:
            self.lista.insert("end", "  " + m)
        if seleccionar in self.marcas:
            i = self.marcas.index(seleccionar)
            self.lista.selection_set(i)
            self.lista.see(i)

    def refrescar_coincidencias(self):
        self.tabla.delete(*self.tabla.get_children())
        self.enlaces = {}
        filas = []
        if HISTORIAL.exists():
            try:
                with open(HISTORIAL, encoding="utf-8-sig", newline="") as f:
                    filas = list(csv.DictReader(f, delimiter=";"))
            except Exception:
                filas = []
        filas.reverse()
        for fl in filas[:500]:
            nivel = fl.get("Nivel") or (fl.get("Tipo de coincidencia", "").split(":")[-1].split("(")[0].strip())
            marca = fl.get("Marca") or (fl.get("Texto del registro", "") or "")[:80]
            iid = self.tabla.insert("", "end", tags=(nivel.split()[0] if nivel else "",), values=(
                nivel, fl.get("Fecha estado diario", ""), fl.get("N° solicitud", ""), marca,
                fl.get("Coincide con", "")))
            self.enlaces[iid] = fl.get("Enlace INAPI", "")
        self.lbl_total.config(text=f"{len(filas)} en total, las más recientes primero" if filas
                              else "Todavía no hay coincidencias.")

    # --- acciones
    def agregar(self):
        texto = " ".join(self.entrada.get().split())
        if not texto:
            return
        n = normalizar(texto)
        if len(n.replace(" ", "")) < 2:
            self.mb.showwarning("Marca muy corta", "Escribe al menos 2 letras o números.", parent=self.raiz)
            return
        if n in {normalizar(m) for m in self.marcas}:
            self.mb.showinfo("Ya está", f"«{texto}» ya está en la lista.", parent=self.raiz)
            self.entrada.delete(0, "end")
            return
        if len(n.replace(" ", "")) <= 2 and not self.mb.askyesno(
                "¿Seguro?", f"«{texto}» es muy corta y puede generar muchos avisos. ¿Agregarla igual?",
                parent=self.raiz):
            return
        guardar_lista_claves(self.marcas + [texto])
        self.entrada.delete(0, "end")
        self.refrescar_marcas(seleccionar=texto)
        self.lbl_pie.config(text=f"«{texto}» agregada. La próxima revisión la buscará en los últimos 30 días "
                                 f"(o pulsa «Revisar ahora»).")

    def quitar(self):
        sel = self.lista.curselection()
        if not sel:
            self.mb.showinfo("Quitar marca", "Primero haz clic en la marca que quieres quitar.", parent=self.raiz)
            return
        marca = self.marcas[sel[0]]
        if not self.mb.askyesno("Quitar marca", f"¿Dejar de vigilar «{marca}»?", parent=self.raiz):
            return
        guardar_lista_claves([m for m in self.marcas if m != marca])
        self.refrescar_marcas()
        self.lbl_pie.config(text=f"«{marca}» ya no se vigila." +
                            ("  Ojo: la lista quedó vacía." if not self.marcas else ""))

    def revisar_ahora(self):
        if self.proceso and self.proceso.poll() is None:
            return
        import subprocess
        flags = 0x08000000 if os.name == "nt" else 0  # sin ventana negra
        self.proceso = subprocess.Popen([ejecutable_sin_ventana(), str(Path(__file__).resolve()), "--manual"],
                                        cwd=str(CARPETA), creationflags=flags)
        self.btn_revisar.state(["disabled"])
        self.btn_revisar.config(text="Revisando…")
        self.lbl_pie.config(text="Revisando INAPI… la primera vez o tras cambiar las marcas puede tardar varios "
                                 "minutos. Puedes cerrar esta ventana: el aviso llegará igual.")
        self.esperar()

    def esperar(self):
        if self.proceso and self.proceso.poll() is None:
            self.raiz.after(1500, self.esperar)
            return
        self.btn_revisar.state(["!disabled"])
        self.btn_revisar.config(text="Revisar ahora")
        self.lbl_pie.config(text="Revisión terminada.")
        self.refrescar_todo()

    def cambiar_inicio(self):
        try:
            activar_inicio(self.var_inicio.get())
        except Exception as e:
            self.var_inicio.set(inicio_activado())
            self.mb.showerror("No se pudo cambiar", str(e), parent=self.raiz)
        self.refrescar_todo()

    def abrir_fila(self, _evento):
        sel = self.tabla.selection()
        if sel and self.enlaces.get(sel[0]):
            import webbrowser
            webbrowser.open(self.enlaces[sel[0]])

    def abrir_si_existe(self, ruta):
        if Path(ruta).exists():
            abrir(ruta)
        else:
            self.mb.showinfo("Aún no existe", "Todavía no hay nada guardado. Pulsa «Revisar ahora».",
                             parent=self.raiz)

    # --- ventana de ajustes
    def abrir_ajustes(self):
        tk, ttk = self.tk, self.ttk
        conf = cargar_config()
        c = conf["correo"]
        v = tk.Toplevel(self.raiz)
        v.title("Ajustes y correo de aviso")
        v.configure(bg=BLANCO)
        v.transient(self.raiz)
        v.resizable(False, False)
        marco = ttk.Frame(v, style="Tarjeta.TFrame", padding=22)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(1, weight=1)

        def titulo(texto, fila):
            ttk.Label(marco, text=texto, style="Titulo.TLabel").grid(row=fila, column=0, columnspan=2,
                                                                     sticky="w", pady=(0, 6))

        def campo(etiqueta, fila, valor, oculto=False):
            ttk.Label(marco, text=etiqueta, style="Tarjeta.TLabel").grid(row=fila, column=0, sticky="w",
                                                                        padx=(0, 12), pady=3)
            e = ttk.Entry(marco, width=46, show="•" if oculto else "")
            e.insert(0, valor)
            e.grid(row=fila, column=1, sticky="ew", pady=3)
            return e

        titulo("Búsqueda", 0)
        var_palabra = tk.BooleanVar(value=conf.get("una_palabra", True))
        ttk.Checkbutton(marco, text="Avisar también si una marca comparte al menos una palabra con las mías",
                        variable=var_palabra, style="Tarjeta.TCheckbutton").grid(row=1, column=0, columnspan=2,
                                                                                 sticky="w")
        e_ign = campo("Palabras que no cuentan:", 2, conf.get("palabras_ignoradas", ""))
        ttk.Label(marco, text="Palabras comunes que no deben gatillar avisos (separadas por coma).",
                  style="Nota.TLabel").grid(row=3, column=1, sticky="w")
        var_sin = tk.BooleanVar(value=conf.get("avisar_sin_novedades", False))
        ttk.Checkbutton(marco, text="Mostrar ventana aunque no haya novedades (revisiones automáticas)",
                        variable=var_sin, style="Tarjeta.TCheckbutton").grid(row=4, column=0, columnspan=2,
                                                                             sticky="w", pady=(4, 16))

        titulo("Correo de aviso (Gmail)", 5)
        var_activo = tk.BooleanVar(value=c.get("activo", False))
        ttk.Checkbutton(marco, text="Enviarme un correo cuando haya coincidencias", variable=var_activo,
                        style="Tarjeta.TCheckbutton").grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 4))
        e_dest = campo("Enviar a:", 7, c.get("destinatarios", ""))
        ttk.Label(marco, text="Puedes poner varios correos separados por coma.",
                  style="Nota.TLabel").grid(row=8, column=1, sticky="w")
        e_rem = campo("Cuenta de Gmail que envía:", 9, c.get("remitente", ""))
        e_clave = campo("Contraseña de aplicación:", 10, "", oculto=True)
        nota_clave = ("Ya hay una guardada. Déjalo vacío para mantenerla." if c.get("clave")
                      else "16 letras que entrega Google en myaccount.google.com/apppasswords")
        ttk.Label(marco, text=nota_clave, style="Nota.TLabel").grid(row=11, column=1, sticky="w")
        var_res = tk.BooleanVar(value=c.get("resumen_sin_novedades", False))
        ttk.Checkbutton(marco, text="Enviar también el informe cuando no haya coincidencias", variable=var_res,
                        style="Tarjeta.TCheckbutton").grid(row=12, column=0, columnspan=2, sticky="w", pady=(6, 6))
        e_srv = campo("Servidor (avanzado):", 13, c.get("servidor", "smtp.gmail.com"))
        e_pto = campo("Puerto (avanzado):", 14, str(c.get("puerto", 465)))
        ayuda = ttk.Label(marco, style="Nota.TLabel", justify="left", text=(
            "Cómo obtener la contraseña de aplicación de Gmail:\n"
            "1) Activa la verificación en dos pasos en tu cuenta de Google.\n"
            "2) Entra a myaccount.google.com/apppasswords, escribe «Vigilancia INAPI» y pulsa Crear.\n"
            "3) Copia las 16 letras y pégalas arriba."))
        ayuda.grid(row=15, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(marco, text="Abrir la página de Google", command=lambda: __import__("webbrowser").open(
            "https://myaccount.google.com/apppasswords")).grid(row=16, column=0, sticky="w", pady=(6, 0))

        def leer_formulario():
            nuevo = cargar_config()
            nuevo["una_palabra"] = bool(var_palabra.get())
            nuevo["palabras_ignoradas"] = e_ign.get().strip()
            nuevo["avisar_sin_novedades"] = bool(var_sin.get())
            co = nuevo["correo"]
            co["activo"] = bool(var_activo.get())
            co["destinatarios"] = e_dest.get().strip()
            co["remitente"] = e_rem.get().strip()
            if e_clave.get().strip():
                co["clave"] = cifrar_clave(e_clave.get().strip().replace(" ", ""))
            co["resumen_sin_novedades"] = bool(var_res.get())
            co["servidor"] = e_srv.get().strip() or "smtp.gmail.com"
            try:
                co["puerto"] = int(e_pto.get().strip() or 465)
            except ValueError:
                co["puerto"] = 465
            return nuevo

        def validar(nuevo, para_prueba=False):
            co = nuevo["correo"]
            if co["activo"] or para_prueba:
                if not destinatarios_de(co):
                    return "Escribe al menos un correo en «Enviar a»."
                if "@" not in co["remitente"]:
                    return "Escribe la cuenta de Gmail que envía (ej.: tucuenta@gmail.com)."
                if not co.get("clave"):
                    return "Falta la contraseña de aplicación."
            return ""

        def probar():
            nuevo = leer_formulario()
            error = validar(nuevo, para_prueba=True)
            if error:
                self.mb.showwarning("Faltan datos", error, parent=v)
                return
            v.config(cursor="watch")
            v.update()
            try:
                cuerpo = ("<div style='font-family:Segoe UI,Arial'><h2>✓ Correo de prueba</h2>"
                          "<p>Si recibes esto, los avisos de <b>Vigilancia de marcas INAPI</b> están "
                          "bien configurados.</p></div>")
                enviar_correo(nuevo["correo"], "Prueba · Vigilancia de marcas INAPI", cuerpo,
                              "Si recibes esto, los avisos de Vigilancia de marcas INAPI están bien configurados.")
                guardar_config(nuevo)
                self.mb.showinfo("Correo enviado", "Listo: revisa tu bandeja de entrada (y la carpeta Spam). "
                                 "Los datos quedaron guardados.", parent=v)
            except Exception as e:
                self.mb.showerror("No se pudo enviar", explicar_error_correo(e), parent=v)
            finally:
                v.config(cursor="")

        def guardar():
            nuevo = leer_formulario()
            error = validar(nuevo)
            if error:
                self.mb.showwarning("Faltan datos", error, parent=v)
                return
            guardar_config(nuevo)
            v.destroy()
            self.refrescar_todo()
            self.lbl_pie.config(text="Ajustes guardados.")

        def diagnostico():
            try:
                ruta = generar_diagnostico("Generado a mano desde el panel")
                self.mb.showinfo("Diagnóstico listo", f"Quedó en:\n{ruta}\n\nEnvíaselo a Claude (arrastra el "
                                 "archivo al chat) si algo no funciona.", parent=v)
                mostrar_en_explorador(ruta)
            except Exception as e:
                self.mb.showerror("No se pudo generar", str(e), parent=v)

        ttk.Button(marco, text="Generar informe de diagnóstico", command=diagnostico).grid(
            row=16, column=1, sticky="e", pady=(6, 0))
        botones = ttk.Frame(marco, style="Tarjeta.TFrame")
        botones.grid(row=17, column=0, columnspan=2, sticky="e", pady=(18, 0))
        ttk.Button(botones, text="Enviar correo de prueba", command=probar).pack(side="left")
        ttk.Button(botones, text="Cancelar", command=v.destroy).pack(side="left", padx=8)
        ttk.Button(botones, text="Guardar", style="Principal.TButton", command=guardar).pack(side="left")
        v.grab_set()
        e_dest.focus_set()

    def correr(self):
        self.raiz.mainloop()


def main():
    if "--panel" in sys.argv:
        Panel().correr()
        return 0
    if "--diagnostico" in sys.argv:
        i = sys.argv.index("--diagnostico")
        mensaje = " ".join(sys.argv[i + 1:])
        ruta = generar_diagnostico(mensaje)
        mostrar_en_explorador(ruta)
        return 0
    manual = "--manual" in sys.argv
    inicio = "--inicio" in sys.argv
    if inicio:
        time.sleep(40)  # deja que Windows termine de arrancar
    if not tomar_bloqueo():
        if manual:
            aviso("Vigilancia de marcas INAPI", "Ya hay una revisión en curso. Espera unos minutos.")
        return 0
    codigo = 0
    try:
        revisar(manual, inicio)
    except Exception as e:
        detalle = "".join(traceback.format_exception(type(e), e, e.__traceback__))
        log("ERROR: " + detalle)
        codigo = 1
        est = cargar_estado()
        est["fallos_seguidos"] = est.get("fallos_seguidos", 0) + 1
        guardar_estado(est)
        if manual or est["fallos_seguidos"] in (3, 10):
            diag = None
            try:
                diag = generar_diagnostico(detalle)
            except Exception as e2:
                log(f"No pude generar el diagnóstico: {e2}")
            texto = f"No pude completar la revisión:\n\n{e}\n\n"
            if diag:
                texto += (f"Generé un informe de diagnóstico para corregirlo:\n{diag.name}\n\n"
                          "Envíaselo a Claude (arrastra el archivo al chat).\n\n¿Abrir la carpeta donde quedó?")
                if aviso("Vigilancia de marcas INAPI", texto, preguntar=True):
                    mostrar_en_explorador(diag)
            else:
                aviso("Vigilancia de marcas INAPI", texto + f"Detalle en: {REGISTRO}")
    finally:
        soltar_bloqueo()
    return codigo


if __name__ == "__main__":
    sys.exit(main())
