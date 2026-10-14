# Reset 2026-3 · Dashboard ejecutivo

Dashboard público (sin login) de inscritos Reset, Equipo Principal y Equipo Next.
Lee y guarda en los Excel de Google Drive:
- Equipo Principal: `Reporte de Avance Inscritos Equipo Principal - Reset.xlsx` (1OrJNlLTA5MMtvDtWeQeunafC0e3WWQGJ)
- Equipo Next: `Reporte de Avance Inscritos Equipo Next - Reset.xlsx` (1b01_tVCwx5lLpkDuXvtyWUwLqAKd1foA)

## Regla de carga (PDF)
- Campo Fecha en formato AAAA-MM-DD (ej. 2026-10-10): 09 → SEP, 10 → OCT, 11 → NOV, 12 → DIC.
- El código del grupo se asigna al líder por prefijo: C186D → C186; C18J3D → C18J3 (Next) y C18J (Principal).
- Solo cuentan códigos de la lista de líderes (Sede Surco). Otros meses u otras sedes se ignoran.
- Se reescriben SEP, OCT, NOV, DIC y TOTAL de los líderes que aparecen en el PDF. LOGRO (2026-2) no se toca.
- El detalle de inscritos solo se muestra en pantalla; no se guarda en Drive ni se envían DNI/celulares.

## Publicar en Vercel (reset-dashboard.vercel.app)
1. Google Cloud Console → crear proyecto → habilitar **Google Drive API**.
2. Crear una **cuenta de servicio** → Claves → Agregar clave JSON (se descarga un archivo).
3. En Drive, compartir los 2 Excel (o la carpeta) con el correo de la cuenta de servicio como **Editor**.
4. Subir esta carpeta a GitHub (o usar `npx vercel` desde la carpeta).
5. En Vercel → Add New Project → nombre del proyecto: **reset-dashboard** → Framework: Other.
6. Settings → Environment Variables:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` = contenido completo del JSON de la cuenta de servicio
   - (opcional) `UPLOAD_KEY` = clave para permitir cargas; si no se define, cualquiera con el link puede subir
   - (opcional) `FILE_ID_PRINCIPAL`, `FILE_ID_NEXT` si cambian los archivos
7. Deploy. Si ya usaste la misma cuenta de servicio para Escuela de Líderes, reutiliza el mismo JSON.
