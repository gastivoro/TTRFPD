# update_cafci.ps1 — Genera cafci_data.js con datos de Toronto Trust RFP Clase B
# Fuente: api.pub.cafci.org.ar/pb_get (Planilla Diaria CAFCI)
# Correr después de las 18:00 hs (CAFCI publica tarde).
# Requiere Python con openpyxl: pip install openpyxl
#
# Uso:  powershell -ExecutionPolicy Bypass -File update_cafci.ps1

$ErrorActionPreference = "Stop"

$dir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp    = Join-Path $env:TEMP "cafci_pb.xlsx"
$output = Join-Path $dir "cafci_data.js"

Write-Host "[CAFCI] Descargando planilla diaria..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://api.pub.cafci.org.ar/pb_get" -OutFile $tmp -UseBasicParsing

$pyFile = Join-Path $env:TEMP "cafci_parser.py"
@'
import openpyxl, json, sys, datetime

tmp = sys.argv[1]
wb  = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
ws  = wb.active
T   = "toronto trust renta fija plus - clase d"

# Columnas de variación: 7=día, 9=MTD, 10=YTD, 11=12M
# Las fechas de referencia están en la sub-fila de encabezado (col 6=ayer, 9=MTD, 10=YTD, 11=12M)
ref_dates = {}
data_row  = None

for row in ws.iter_rows(min_row=1, max_row=1500, values_only=True):
    # Sub-header: la fila que tiene fecha en col[10] (ej: "30/12/25") — puede tener row[0]=None
    if not ref_dates and isinstance(row[10], str) and len(str(row[10])) == 8 and '/' in str(row[10]):
        ref_dates = {
            'fecha_ayer': str(row[6])  if row[6]  else '',
            'fecha_mtd':  str(row[9])  if row[9]  else '',
            'fecha_ytd':  str(row[10]) if row[10] else '',
            'fecha_12m':  str(row[11]) if row[11] else '',
        }
    if row[0] and isinstance(row[0], str) and T in row[0].lower():
        data_row = row
        break

if not data_row:
    print("ERROR: fondo no encontrado en la planilla", file=sys.stderr)
    sys.exit(1)

d = dict(
    nombre    = data_row[0],
    fecha     = str(data_row[4])         if data_row[4]  else '',
    vcp       = round(float(data_row[5]), 6) if data_row[5]  else 0,
    dia       = round(float(data_row[7]), 4) if data_row[7]  else 0,
    mtd       = round(float(data_row[9]), 4) if data_row[9]  else 0,
    ytd       = round(float(data_row[10]),4) if data_row[10] else 0,
    m12       = round(float(data_row[11]),4) if data_row[11] else 0,
    ts        = datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
)
d.update(ref_dates)
print(json.dumps(d))
sys.exit(0)
'@ | Set-Content -Path $pyFile -Encoding utf8

$result = python3 $pyFile $tmp 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "[CAFCI] Fallo: $result"
    exit 1
}

$json = ($result | Where-Object { $_ -match '^\{' }) | Select-Object -First 1
$data = $json | ConvertFrom-Json

$ts = Get-Date -Format "yyyy-MM-dd HH:mm"
$js = "// Auto-generado — $ts`n// $($data.nombre)`nwindow.CAFCI_DATA = $json;"

Set-Content -Path $output -Value $js -Encoding utf8

Write-Host "[CAFCI] OK" -ForegroundColor Green
Write-Host "  Fondo : $($data.nombre)"
Write-Host "  Fecha : $($data.fecha)"
Write-Host "  VCP   : $($data.vcp)"
Write-Host "  Dia   : $($data.dia)%  (vs $($data.fecha_ayer))"
Write-Host "  MTD   : $($data.mtd)%  (desde $($data.fecha_mtd))"
Write-Host "  YTD   : $($data.ytd)%  (desde $($data.fecha_ytd))"
Write-Host "  12M   : $($data.m12)%  (desde $($data.fecha_12m))"
Write-Host "  Arch. : $output"
