# Seguimiento_licencias

⚙️ INSTALACIÓN Y USO

1️⃣ Instala dependencias (una sola vez):

py -m pip install pandas openpyxl python-dateutil pywin32


2️⃣ Ajusta:

EXCEL_PATH → ruta a tu archivo Control_Licencias.xlsx

EMAIL_TO → destinatarios (coma si varios)

Opcionalmente, cambia mail.Display() por mail.Send() para enviar sin confirmación.

--------------

Estructura del archivo de excel 

| Nº | Tipo     | Producto / Servicio | Fabricante | Nº Serie / Clave | Usuario / Área | Fecha Inicio | Fecha Fin  | Días Restantes           | Estado                                                                      | Aviso                                                   |
| -- | -------- | ------------------- | ---------- | ---------------- | -------------- | ------------ | ---------- | ------------------------ | --------------------------------------------------------------------------- | ------------------------------------------------------- |
| 1  | Software | Microsoft 365       | Microsoft  | XXXXX-XXXXX      | IT             | 01/01/2024   | 31/12/2025 | `=SI(H2="";"";H2-HOY())` | `=SI(I2="";"";SI(I2<=0;"Vencido";SI(I2<=120;"Próximo a vencer";"Activo")))` | `=SI(J2="Próximo a vencer";"⚠️ Revisar renovación";"")` |




3️⃣ Ejecuta el script:

py .\enviar_alertas_outlook.py


Outlook abrirá un nuevo correo con la tabla de alertas (o lo enviará directamente si activas .Send()).

🧠 Ventajas de esta versión

✅ No usa contraseñas ni configuración SMTP.
✅ Funciona en entornos corporativos con Outlook / Microsoft 365.
✅ Permite revisión manual antes del envío.
✅ 100 % compatible con Windows.

-----------------

🧩 Instalación y uso en Windows

Instala Python (si no lo tienes) y luego:

py -m pip install --upgrade pip
py -m pip install pandas openpyxl python-dateutil


Configura variables de entorno (ejemplos para Office 365):

setx SMTP_SERVER "smtp.office365.com"
setx SMTP_PORT "587"
setx SMTP_USER "tu_correo@tu_dominio.com"
setx SMTP_PASS "tu_contraseña_o_app_password"
setx EMAIL_FROM "tu_correo@tu_dominio.com"
setx EMAIL_TO "destinatario1@dominio.com,destinatario2@dominio.com"


Cierra y vuelve a abrir la consola para que tomen efecto, o usa $env:SMTP_USER="..." en la sesión actual.

Ajusta la ruta del Excel en la variable EXCEL_PATH del script.

Ejecuta:

py .\enviar_alertas_licencias.py

📌 Notas y opciones

El script no necesita que el Excel calcule fórmulas: recalcula fechas y estados en Python.

Formato de fecha esperado en la columna “Fecha Fin”: dd/mm/yyyy.

Cambia el umbral con THRESHOLD_DAYS = 120.

Si quieres que siempre envíe correo, incluso sin alertas, pon SEND_IF_EMPTY = True.

Si tu servidor SMTP requiere otra configuración (por ejemplo, servidor interno), ajusta SMTP_SERVER/PORT.
