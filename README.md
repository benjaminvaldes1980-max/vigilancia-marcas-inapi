# Vigilancia de marcas INAPI

Revisa **todos los días, sola y en la nube** los Estados Diarios de Marcas de INAPI
(https://tramites.inapi.cl/EstadosDiariosMarcas) y avisa por **correo a Gmail** si alguien
inscribe una marca **igual**, **parecida** o que **comparte una palabra** con las marcas vigiladas.
No necesita que tu computador esté encendido.

## Qué hace cada día (~09:15 hora de Chile)
1. Abre la página de INAPI con un navegador invisible y baja los PDF nuevos
   (Presentaciones y Estado diario versión extendida).
2. Lee todas las marcas de esos PDF y las compara con `palabras_clave.txt`.
3. Si encuentra algo nuevo, te envía un correo con la lista y el informe adjunto.
4. Publica el último informe en la página del proyecto (GitHub Pages) y guarda el historial.

## Agregar o quitar marcas
Abre `palabras_clave.txt` → ícono del **lápiz** → escribe o borra líneas → **Commit changes**.
Al guardar, se hace una revisión de inmediato con las marcas nuevas (últimos 30 días).

## Revisar ahora
Pestaña **Actions** → **Vigilancia de marcas INAPI** → **Run workflow**.

## Correo de aviso (una sola vez)
**Settings → Secrets and variables → Actions → New repository secret**, y crea:

| Nombre | Valor |
|---|---|
| `GMAIL_USUARIO` | tu cuenta de Gmail (ej.: tucuenta@gmail.com) |
| `GMAIL_CLAVE_APP` | la contraseña de aplicación de 16 letras (myaccount.google.com/apppasswords) |
| `AVISO_PARA` | *(opcional)* a quién avisar; varios separados por coma |

## Archivos
- `informes/` – un informe por revisión (cuántas marcas se revisaron y qué se encontró).
- `historial_coincidencias.csv` – todo lo encontrado (se abre en Excel).
- `historial_revisiones.csv` – una fila por revisión.
- `docs/index.html` – el último informe (página pública si activas GitHub Pages).
