import win32com.client as win32
import os
import sys
import boto3
import webbrowser
import time
import pyperclip

from tkinter import Tk, messagebox
from datetime import datetime

# =========================
# CONFIG
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PDF_DIR = os.path.join(BASE_DIR, "pdf")

os.makedirs(PDF_DIR, exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

PDF_NAME = f"OC-{timestamp}.pdf"

PDF_PATH = os.path.join(PDF_DIR, PDF_NAME)

# =========================
# EXCEL
# =========================
print("Buscando Excel abierto...")

excel = win32.GetActiveObject("Excel.Application")

wb = excel.ActiveWorkbook

if wb is None:
    print("No hay Excel activo")
    sys.exit()

# =========================
# GENERAR PDF
# =========================
print("Generando PDF...")

sheet = wb.Worksheets("OrdendeCompra")

sheet.ExportAsFixedFormat(0, PDF_PATH)

# esperar creación real
for _ in range(20):

    if os.path.exists(PDF_PATH):
        break

    time.sleep(0.5)

# validar
if not os.path.exists(PDF_PATH):

    messagebox.showerror(
        "Error",
        "No fue posible generar el PDF."
    )

    sys.exit()

print("PDF generado:")
print(PDF_PATH)

# =========================
# ABRIR PDF
# =========================
webbrowser.open(PDF_PATH)

time.sleep(2)

# =========================
# CONFIRMACION
# =========================
root = Tk()

root.withdraw()

respuesta = messagebox.askyesno(
    "Confirmar envío",
    "La Orden de Compra está lista.\n\n"
    "¿Deseas continuar con el envío?"
)

# =========================
# CANCELAR
# =========================
if not respuesta:

    try:
        os.remove(PDF_PATH)
    except:
        pass

    wb.Worksheets("OrdendeCompra").Activate()

    root.destroy()

    sys.exit()

# =========================
# R2 CONFIG
# =========================
ACCOUNT_ID = "d3180e3d4fa713047894083e0dc23652"

ACCESS_KEY_ID = "e7f3fe418ba1528356802d5bc9feab81"

SECRET_ACCESS_KEY = "75f945ee46faace9ac3b8b2f4c4be2b49c1c6bbba4ebb5998ef03786022f5738"

BUCKET_NAME = "ordenes-compra"

PUBLIC_URL = "https://pub-a75717827ac2456ab5dd4b555345bfbe.r2.dev"

# =========================
# SUBIR A R2
# =========================
print("Subiendo a R2...")

s3 = boto3.client(
    service_name="s3",
    endpoint_url=f"https://{ACCOUNT_ID}.r2.cloudflarestorage.com",
    aws_access_key_id=ACCESS_KEY_ID,
    aws_secret_access_key=SECRET_ACCESS_KEY,
    region_name="auto"
)

with open(PDF_PATH, "rb") as f:

    s3.upload_fileobj(
        f,
        BUCKET_NAME,
        PDF_NAME,
        ExtraArgs={
            "ContentType": "application/pdf"
        }
    )

# =========================
# URL FINAL
# =========================
public_url = f"{PUBLIC_URL}/{PDF_NAME}"

print("PDF SUBIDO:")
print(public_url)

# =========================
# COPIAR LINK
# =========================
pyperclip.copy(public_url)

# =========================
# LIMPIAR PEDIDO
# =========================
pedido_sheet = wb.Worksheets("Pedido")

last_row = pedido_sheet.Cells(
    pedido_sheet.Rows.Count,
    6
).End(-4162).Row

pedido_sheet.Range(
    f"F2:F{last_row}"
).ClearContents()

# =========================
# VOLVER A OC
# =========================
wb.Worksheets("OrdendeCompra").Activate()

# =========================
# GUARDAR
# =========================
wb.Save()

# =========================
# MENSAJE FINAL
# =========================
messagebox.showinfo(
    "Orden enviada",
    f"La Orden de Compra fue procesada correctamente.\n\n"
    f"LINK PARA COMPARTIR:\n\n"
    f"{public_url}\n\n"
    f"El enlace fue copiado automáticamente."
)

root.destroy()