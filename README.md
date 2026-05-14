# PRIZE PRO - Actualización de Datos Persistente

## Qué incluye esta versión

- Proyecto listo para Render + GitHub.
- Login administrador/operador.
- CRUD de usuarios.
- Carga incremental de base de trabajadores por DNI.
- Registro de correo, celular, nivel educativo, procedencia, indumentaria, tiempo, carnet CONADIS, teléfono de emergencia y observación.
- Soporte para PostgreSQL mediante `DATABASE_URL`.
- Exportación a Excel desde el sistema.
- Nueva pestaña **Respaldo** para descargar:
  - respaldo total,
  - base de trabajadores,
  - datos actualizados.
- Corrección para migraciones en PostgreSQL.

## Usuario inicial

Usuario: `admin`  
Clave: `admin123`

Cambia la clave luego de ingresar.

## Arquitectura recomendada

```text
Flask App
   ↓
PostgreSQL (datos principales)
   ↓
Excel (reportes / exportaciones / respaldo descargable)
```

## Importante sobre Render Free

Render permite Web Service gratis, pero los archivos locales se pueden perder al reiniciar o redesplegar. Por eso **no debes depender de Excel local, SQLite local ni carpetas locales como almacenamiento principal**.

Para persistencia real usa PostgreSQL y configura la variable:

```text
DATABASE_URL=postgresql://usuario:clave@host:puerto/base
```

Render indica que las bases PostgreSQL Free tienen 1 GB y expiran después de 30 días. Para datos de producción conviene una base pagada de Render o una base externa tipo Neon/Supabase con plan gratuito y límites.

## Cómo proceder en Render

### Opción recomendada para cuenta gratuita

1. Sube este proyecto a GitHub.
2. Crea o actualiza tu Web Service en Render.
3. En Render, entra a tu Web Service.
4. Ve a **Environment**.
5. Agrega estas variables:

```text
SECRET_KEY=una_clave_larga_segura
APP_TIMEZONE=America/Lima
DATABASE_URL=pega_aqui_tu_url_postgresql
```

6. Haz **Manual Deploy > Clear build cache & deploy**.
7. Entra al sistema con `admin / admin123`.
8. Carga tu base de trabajadores desde **Cargar base**.
9. Descarga respaldos desde la pestaña **Respaldo**.

### Sobre PostgreSQL en Render

Si creas PostgreSQL gratis dentro de Render, úsalo solo para pruebas porque expira después de 30 días. Si lo usarás para datos reales, usa una base pagada o una externa gratuita con mejor permanencia.

## Descarga de información

Dentro del sistema tienes:

- **Registros > Exportar Excel**: descarga datos actualizados.
- **Cargar base > Descargar base actual**: descarga trabajadores.
- **Respaldo > Descargar respaldo total Excel**: descarga todo en un solo archivo.

## Comandos locales

```bash
pip install -r requirements.txt
python app.py
```

Luego abre:

```text
http://localhost:5000
```

## Build y Start Command para Render

Build Command:

```bash
python -m pip install --upgrade pip && pip install -r requirements.txt
```

Start Command:

```bash
gunicorn app:app
```
