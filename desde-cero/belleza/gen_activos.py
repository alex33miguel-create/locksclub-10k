import openpyxl
from openpyxl.styles import Font
import qrcode, os
base = r"C:\Users\alex3\OneDrive\Escritorio\Proyecto OPencode 1\desde-cero\belleza"
os.makedirs(base, exist_ok=True)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Citas"
headers = ["Fecha","Hora","Cliente","Servicio","Precio","Anticipo","Resta","Estado","IG_Contacto"]
ws.append(headers)
for c in ws[1]:
    c.font = Font(bold=True)
servs = [("Uñas acrílico",499),("Corte+barba",499),("Facial",599),("Pedicure",449),("Membresía x4",1199)]
from datetime import date, timedelta
r=2
for i in range(20):
    s,p = servs[i % len(servs)]
    ws.append([(date.today()+timedelta(days=i%7)).isoformat(),"10:00" if i%2==0 else "17:00",f"Cliente {i+1}",s,p,250,f"=E{r}-F{r}","Apartado",f"@cliente{i+1}"])
    r+=1
ws2 = wb.create_sheet("Resumen")
ws2.append(["Meta 7 días",10000])
ws2.append(["Venta esperada 20 bonos x499","=20*499"])
ws2.append(["Venta 9 membresías x1199","=9*1199"])
ws2.append(["Utilidad día = Ventas-Gastos","=SUMA(Citas!E:E)-0"])
xlsx = os.path.join(base,"control-citas.xlsx")
wb.save(xlsx)
print("XLSX OK", xlsx)
for name,url in [("qr-pago.png","https://link.mercadopago.com.mx/locksclub"),("qr-tienda.png","https://alex33miguel-create.github.io/locksclub-10k/desde-cero/servicios/")]:
    img = qrcode.make(url)
    p = os.path.join(base,name)
    img.save(p)
    print("QR OK", p)
