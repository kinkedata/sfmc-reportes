# 📊 Sistema Automatizado de Reportes Telcel Empresas v1.4 FINAL

Genera automáticamente reportes Excel encriptados por región desde datos de SFMC.

---

## 🚀 Quick Start (2 minutos)

### 1. **Instalar requisitos**

```bash
pip install openpyxl msoffcrypto-tool
```

> `msoffcrypto-tool` es opcional pero recomendado: habilita encriptación real de los Excel. Sin él, los Excel se generan sin contraseña.

### 2. **Exportar CSV de SFMC**

- Accede a SFMC → Contacts → Data Extensions
- Busca `de-sol-forms`
- Haz clic en "Download" → CSV
- Guarda en: `C:\Reportes\Telcel\leads_diarios.csv`

### 3. **Ejecutar el script**

```bash
python generar_reportes.py
```

Output esperado:
```
============================================================
🚀 GENERADOR DE REPORTES TELCEL EMPRESAS v1.4
============================================================
Fecha: 12-03-2026 09:00:00

📖 Leyendo CSV: C:\Reportes\Telcel\leads_diarios.csv
✅ 52 registros cargados

📅 Leads del 11-03-2026: 35

📊 Distribución de Leads:
  R1: 8 leads
  R2: 6 leads
  R3: 7 leads
  ...

📝 Generando archivos Excel con protección...

✅ R1: R1-leads-calificados-x-dia-r1-2026-03-12.xlsx (8 leads) 🔒 PROTEGIDO R1TELABC
✅ R2: R2-leads-calificados-x-dia-r2-2026-03-12.xlsx (6 leads) 🔒 PROTEGIDO R2TELX4F
...

🔐 Contraseñas guardadas en: CONTRASEÑAS_12032026.txt

============================================================
✅ PROCESO COMPLETADO
============================================================
```

> Si aparece `✅ CREADO` en lugar de `🔒 PROTEGIDO`, instala msoffcrypto-tool: `pip install msoffcrypto-tool`

### 4. **Enviar mails (manual)**

1. Abre `CONTRASEÑAS_DDMMYYYY.txt` para ver las contraseñas del día
2. Consulta `config.json` para saber los destinatarios de cada región
3. Por cada región con leads: redacta email, adjunta Excel, incluye contraseña en el cuerpo

---

## ⚙️ Automatización (Ejecución diaria a las 9 AM)

### **Windows (Task Scheduler)**

1. **Abre CMD como Administrador**
   - Windows + R → `cmd` → Ctrl+Shift+Enter

2. **Navega a la carpeta del script**
   ```bash
   cd C:\Reportes\Telcel\Scripts
   ```

3. **Ejecuta el instalador**
   ```bash
   instalar_scheduler_windows.bat
   ```

4. **Verifica que se creó la tarea**
   ```bash
   schtasks /query /tn TelcelReportesAuto_9AM
   ```

✅ Listo. Se ejecutará automáticamente mañana a las 9 AM.

---

### **Linux / Mac (cron)**

1. **Abre terminal**

2. **Navega a la carpeta del script**
   ```bash
   cd /home/usuario/Reportes/Telcel
   ```

3. **Dale permisos al script**
   ```bash
   chmod +x instalar_scheduler_linux.sh
   ```

4. **Ejecuta el instalador**
   ```bash
   sudo ./instalar_scheduler_linux.sh
   ```

5. **Verifica**
   ```bash
   crontab -l
   ```

✅ Listo. Se ejecutará automáticamente a las 9 AM.

---

## 📁 Estructura de Archivos

```
C:\Reportes\Telcel\
├── leads_diarios.csv                    ← Exporta aquí desde SFMC (ruta en config.json)
├── Respaldo\                            ← Respaldo automático
└── Scripts\
    ├── config.json                      ← Configuración (regiones, ejecutivos, rutas)
    ├── generar_reportes.py              ← Script principal v1.4 FINAL
    ├── instalar_scheduler_windows.bat   ← Para Windows
    ├── instalar_scheduler_linux.sh      ← Para Linux/Mac
    ├── README.md                        ← Este archivo
    ├── GUIA_RAPIDA.txt                  ← Guía rápida de uso
    ├── ejemplo_leads_diarios.csv        ← Para testing sin SFMC
    │
    └── Excel_Generados\                 ← Se crea automáticamente
        ├── R1-leads-calificados-x-dia-r1-2026-03-12.xlsx  🔒
        ├── R2-leads-calificados-x-dia-r2-2026-03-12.xlsx  🔒
        ├── ... (hasta R9 si tienen leads)
        ├── MM-tel-leads-mobile-marketing-2026-03-12.xlsx   🔒
        ├── BI-tel-leads-business-intellig-2026-03-12.xlsx  🔒
        └── CONTRASEÑAS_12032026.txt     ← Contraseñas aleatorias del día
```

---

## 🔐 Contraseñas Aleatorias

**Formato:** `{REGION}TEL{3 caracteres aleatorios de A-Z y 0-9}`

Ejemplo de una ejecución:
```
R1  :  R1TELABC
R2  :  R2TELX4F
R3  :  R3TEL9ZK
...
MOM :  MOMTELQ7R
BUI :  BUITELW2P
```

**Las contraseñas son completamente aleatorias en cada ejecución.**
Se guardan en `CONTRASEÑAS_DDMMYYYY.txt` (en la carpeta de salida).

Los Excel se encriptan con `msoffcrypto-tool`. Al abrir el archivo,
Excel solicita la contraseña antes de mostrar el contenido.

---

## 📋 Columnas en los Reportes

```
Correo
Nombre
Apellido
Telefono
Estado
Municipio
RazonSocial
Cargo
OtroCargo
NoEmpleados
SolucionInteres
ServicioInteres
Mensaje
CreatedAt
Region
CorreoComercial
```

**Excluidas:** AvisoPrivacidad, Newsletter, Tipo, utm_source, utm_medium, utm_campaign, utm_content

---

## 📊 Distribución de Leads

El sistema separa automáticamente:

### **Por Solución (prioritario)**
- **Mobile Marketing (MOM)**:
  - Si `SolucionInteres` contiene "Mobile Marketing" o es exactamente "MOM"
  - → Se clasifica en el archivo MOM

- **Business Intelligence (BUI)**:
  - Si `SolucionInteres` contiene "Business Intelligence" o es exactamente "BUI"
  - → Se clasifica en el archivo BUI

### **Por Región (R1-R9)**
- Si no es MOM ni BUI, se clasifica según el campo `Region`
- Solo se genera archivo si hay al menos un lead en esa región

### **Lógica de Clasificación**
```
1. ¿SolucionInteres contiene "Mobile Marketing" o es "MOM"? → Archivo MOM
2. ¿SolucionInteres contiene "Business Intelligence" o es "BUI"? → Archivo BUI
3. Si no → Se clasifica por campo Region (R1-R9)
4. Leads sin región válida → Ignorados
```

Los destinatarios de cada región están en `config.json` → sección `"regiones"`.

---

## 🛠️ Personalizar Configuración

Edita `config.json` para cambiar:

### **Cambiar hora de ejecución**
```json
"hora_ejecucion": "10:00"  // En lugar de 09:00
```

### **Agregar/modificar ejecutivos**
```json
"R1": {
  "nombre": "Región 1",
  "ejecutivos": {
    "junior": "nuevo@email.com",
    "senior": "otro@email.com"
  }
}
```

### **Cambiar ruta de carpetas**
```json
"rutas": {
  "csv_entrada": "C:\\NuevaRuta\\leads_diarios.csv",
  "carpeta_salida": "C:\\NuevaRuta\\Excel_Generados",
  "carpeta_respaldo": "C:\\NuevaRuta\\Respaldo"
}
```

### **Agregar/quitar columnas del reporte**
```json
"columnas_reporte": [
  "Correo",
  "Nombre",
  "TuNuevaColumna"
]
```
Las columnas deben existir en el CSV de SFMC.

**IMPORTANTE:** Después de editar `config.json`, reinicia la tarea programada.

---

## 🐛 Troubleshooting

### **"No encontrado: leads_diarios.csv"**
- ✅ Verifica la ruta `csv_entrada` en `config.json`
- ✅ El CSV debe estar exactamente en esa ruta con ese nombre

### **"FileNotFoundError: config.json"**
- ✅ `config.json` debe estar en el mismo directorio que `generar_reportes.py`
- ✅ Verifica sintaxis JSON correcta (jsonlint.com)

### **La tarea programada no se ejecuta**
- ✅ Abre Task Scheduler (Windows) y verifica que está **Enabled**
- ✅ Ejecuta manualmente para ver errores: `python generar_reportes.py`
- ✅ Verifica que Python esté en el PATH: `python --version`

### **Los Excel no se generan / "No hay leads"**
- ✅ El CSV debe tener leads del DÍA ANTERIOR (no de hoy)
- ✅ Verifica el campo `CreatedAt` — el script soporta múltiples formatos
- ✅ Abre el CSV con Excel y verifica que tenga datos con fecha de ayer

### **Los Excel no están encriptados (muestra ✅ CREADO)**
- ✅ Instala msoffcrypto-tool: `pip install msoffcrypto-tool`
- ✅ Verifica que `msoffcrypto-tool` esté en el PATH: ejecuta `msoffcrypto-tool --help`
- ✅ Después de instalarlo, vuelve a ejecutar el script

---

## 🔄 Flujo Diario Completo

```
09:00 AM → Script se ejecuta automáticamente
    ↓
Lee CSV de SFMC (leads del día anterior)
    ↓
Clasifica leads por Solución (MOM/BUI) y Región (R1-R9)
    ↓
Genera Excel encriptados (solo regiones con leads)
    ↓
Crea CONTRASEÑAS_DDMMYYYY.txt con contraseñas aleatorias
    ↓
Tu intervención: Abres CONTRASEÑAS_DDMMYYYY.txt
    ↓
Consultas config.json para los destinatarios de cada región
    ↓
Por cada región: redactas email, adjuntas Excel, incluyes contraseña
    ↓
Envías el email
```

---

## 📧 Template de Email

Para facilitar, puedes usar este template (adaptar por región):

```
TO: ignacio.sanchez@telcel.com, melpadsa@telcel.com
CC: gisela.sosa@demo.reseller.telcel.com, enrique.ramirez@courtavenue.com

ASUNTO: Reporte de leads R1 - 11 de Marzo 2026

CUERPO:
Buenos días equipo.
Comparto el reporte de leads generados para la R1 correspondiente al período del 11 de Marzo.
Contraseña del archivo: R1TELABC
(La contraseña se pide al abrir el Excel)
Quedo pendiente.
Saludos.

ADJUNTAR: R1-leads-calificados-x-dia-r1-2026-03-12.xlsx
```

Los destinatarios de cada región están en `config.json`.

---

## ✅ Checklist de Instalación

- [ ] Python 3.8+ instalado (`python --version`)
- [ ] openpyxl instalado (`pip install openpyxl`)
- [ ] msoffcrypto-tool instalado (`pip install msoffcrypto-tool`)
- [ ] `config.json` editado con tus rutas y datos
- [ ] CSV de SFMC guardado en la ruta de `csv_entrada` en config.json
- [ ] Script `generar_reportes.py` probado — Excel con `🔒 PROTEGIDO`
- [ ] Tarea programada instalada (Windows/Linux)
- [ ] Probaste con un email de test adjuntando Excel + contraseña

---

## 📞 Soporte

**Si algo no funciona:**

1. Ejecuta manualmente: `python generar_reportes.py` (para ver el error exacto)
2. Verifica la sintaxis del `config.json` en jsonlint.com
3. Asegúrate que Python puede acceder a las carpetas
4. Verifica que openpyxl y msoffcrypto-tool están instalados

---

**Última actualización:** 12-03-2026
**Autor:** Sistema Automatizado Telcel
**Versión:** 1.4 FINAL
