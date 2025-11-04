# -*- coding: utf-8 -*-
"""
Script: enviar_alertas_outlook.py
Descripción:
  - Lee Control_Licencias.xlsx
  - Calcula licencias próximas a vencer o vencidas
  - Envía un email con Outlook (sin contraseña)
Requiere:
  pip install pandas openpyxl python-dateutil pywin32
"""

import pandas as pd
from datetime import date, datetime
from dateutil import tz
from pathlib import Path
import win32com.client as win32

# ================== CONFIGURACIÓN ==================
EXCEL_PATH = Path(r"C:\Ruta\A\Control_Licencias.xlsx")  # <-- Ajusta ruta
SHEET_NAME = "Control_Licencias"
THRESHOLD_DAYS = 120   # 4 meses aprox
LOCAL_TZ = tz.gettz("Europe/Madrid")

# Dirección o direcciones destino (coma si varias)
EMAIL_TO = "destinatario@empresa.com"

# Asunto del correo
EMAIL_SUBJECT = "⚠️ Aviso: Licencias próximas a vencer o vencidas"

# ¿Enviar correo si no hay alertas?
SEND_IF_EMPTY = False
# ====================================================


def leer_excel(path: Path) -> pd.DataFrame:
    """Lee y procesa el Excel de licencias"""
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    df["Fecha Fin"] = pd.to_datetime(df["Fecha Fin"], dayfirst=True, errors="coerce")

    df = df.dropna(subset=["Fecha Fin"]).copy()
    hoy = pd.Timestamp(date.today())
    df["Días Restantes"] = (df["Fecha Fin"] - hoy).dt.days

    def estado(d):
        if pd.isna(d): return ""
        if d <= 0: return "Vencido"
        elif d <= THRESHOLD_DAYS: return "Próximo a vencer"
        return "Activo"

    df["Estado"] = df["Días Restantes"].apply(estado)
    return df


def construir_html(df_alertas: pd.DataFrame) -> str:
    """Construye la tabla HTML para el correo"""
    style = """
    <style>
      table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
      th, td { border: 1px solid #ddd; padding: 8px; font-size: 13px; }
      th { background: #f4f4f4; text-align: left; }
      .vencido { background: #ffebee; color: #b71c1c; font-weight: 600; }
      .proximo { background: #fff8e1; color: #9e7700; font-weight: 600; }
    </style>
    """

    rows = []
    for _, r in df_alertas.iterrows():
        estado = r["Estado"]
        css = "vencido" if estado == "Vencido" else "proximo"
        producto = str(r["Producto / Servicio"])
        fabricante = str(r.get("Fabricante", ""))
        fecha = r["Fecha Fin"].strftime("%d/%m/%Y")
        dias = int(r["Días Restantes"])
        rows.append(f"""
            <tr class="{css}">
              <td>{producto}</td>
              <td>{fabricante}</td>
              <td>{fecha}</td>
              <td style="text-align:right;">{dias}</td>
              <td>{estado}</td>
            </tr>
        """)

    ahora = datetime.now(LOCAL_TZ).strftime("%d/%m/%Y %H:%M")
    html = f"""
    <html>
    <head>{style}</head>
    <body>
      <p><strong>Aviso de licencias próximas a vencer o vencidas</strong></p>
      <p>Generado: {ahora} (Europe/Madrid)</p>
      <table>
        <thead>
          <tr>
            <th>Producto / Servicio</th>
            <th>Fabricante</th>
            <th>Fecha Fin</th>
            <th>Días Restantes</th>
            <th>Estado</th>
          </tr>
        </thead>
        <tbody>
          {''.join(rows)}
        </tbody>
      </table>
      <p style="font-size:12px;color:#666;">Umbral configurado: {THRESHOLD_DAYS} días.</p>
    </body>
    </html>
    """
    return html


def enviar_por_outlook(subject: str, html_body: str, to: str):
    """Envía el correo usando la sesión de Outlook activa"""
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.To = to
    mail.HTMLBody = html_body
    mail.Display()  # Abre el correo en pantalla (útil para revisar antes de enviar)
    # mail.Send()   # <-- Descomenta si quieres que se envíe automáticamente
    print(f"📧 Correo preparado para: {to}")


def main():
    print(f"📂 Leyendo Excel: {EXCEL_PATH}")
    df = leer_excel(EXCEL_PATH)
    alertas = df[df["Estado"].isin(["Vencido", "Próximo a vencer"])]

    print(f"Total registros: {len(df)}")
    print(f"Alertas encontradas: {len(alertas)}")

    if alertas.empty and not SEND_IF_EMPTY:
        print("✅ No hay alertas, no se enviará correo.")
        return

    if alertas.empty and SEND_IF_EMPTY:
        html = "<p>No hay licencias próximas a vencer ni vencidas.</p>"
    else:
        html = construir_html(alertas)

    enviar_por_outlook(EMAIL_SUBJECT, html, EMAIL_TO)
    print("✅ Proceso completado.")


if __name__ == "__main__":
    main()
