#!/bin/bash

# Scheduler automático para Reportes Telcel - cron (Linux/Mac)

echo ""
echo "========================================"
echo "Configurador de Automatización - Telcel"
echo "========================================"
echo ""

# Verificar si se ejecuta como root (recomendado para cron)
if [[ $EUID -ne 0 ]]; then
   echo "[*] No se ejecuta como root. Se usará el cron del usuario actual."
   echo ""
fi

# Obtener directorio actual
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
SCRIPT_PATH="$SCRIPT_DIR/generar_reportes.py"

echo "[*] Ruta del script: $SCRIPT_PATH"
echo "[*] Directorio: $SCRIPT_DIR"
echo ""

# Verificar que el script existe
if [ ! -f "$SCRIPT_PATH" ]; then
    echo "[ERROR] No se encontró el script en: $SCRIPT_PATH"
    exit 1
fi

# Verificar que Python esté instalado
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 no está instalado"
    exit 1
fi

echo "[OK] Python encontrado: $(python3 --version)"
echo ""

# Crear archivo de log
LOG_FILE="$SCRIPT_DIR/reportes_telcel.log"
echo "[*] Archivo de log: $LOG_FILE"
echo ""

# Crear entrada cron
CRON_ENTRY="0 9 * * * cd $SCRIPT_DIR && python3 $SCRIPT_PATH >> $LOG_FILE 2>&1"

echo "[*] Instalando entrada cron..."
echo ""
echo "   Hora: 09:00 (9 AM)"
echo "   Frecuencia: Diaria"
echo "   Comando: $CRON_ENTRY"
echo ""

# Agregar a crontab (sin duplicar si ya existe)
(crontab -l 2>/dev/null | grep -v "$SCRIPT_PATH"; echo "$CRON_ENTRY") | crontab -

if [ $? -eq 0 ]; then
    echo "[OK] Entrada cron instalada exitosamente!"
    echo ""
    echo "[*] Crontab actual:"
    crontab -l | grep -A1 -B1 "$SCRIPT_PATH"
    echo ""
    echo "========================================"
    echo "Para desinstalar:"
    echo "  crontab -e"
    echo "  (y elimina la línea con generar_reportes.py)"
    echo "========================================"
    echo ""
else
    echo "[ERROR] No se pudo instalar la entrada cron"
    exit 1
fi
