# update_cafci.ps1 — Genera cafci_data.js con cuotaparte de Toronto Trust Renta Fija Clase B
# Correr después de las 18:00 hs (CAFCI publica tarde).
# Requiere Python con openpyxl instalado: pip install openpyxl
#
# Uso:  powershell -ExecutionPolicy Bypass -File update_cafci.ps1

$ErrorActionPreference = "Stop"

$dir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp    = Join-Path $env:TEMP "cafci_pb.xlsx"
$output = Join-Path $dir "cafci_data.js"

Write-Host "[CAFCI] Descargando planilla diaria..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://api.pub.cafci.org.ar/pb_get" -OutFile $tmp -UseBasicParsing

# Script Python guardado en archivo temporal para evitar problemas de comillas
$pyFile = Join-Path $env:TEMP "cafci_parser.py"
@'
import openpyxl, json, sys, datetime

tmp   = sys.argv[1]
wb    = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
ws    = wb.active
T     = "toronto trust renta fija - clase b"

for row in ws.iter_rows(min_row=1, max_row=1500, values_only=True):
    if not row[0] or not isinstance(row[0], str):
        continue
    if T in row[0].lower():
        d = dict(
            nombre = row[0],
            fecha  = str(row[4]) if row[4] else "",
            vcp    = round(float(row[5]),  6) if row[5]  else 0,
            dia    = round(float(row[7]),  4) if row[7]  else 0,
            ytd    = round(float(row[10]), 4) if row[10] else 0,
            m12    = round(float(row[11]), 4) if row[11] else 0,
            ts     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        )
        print(json.dumps(d))
        sys.exit(0)

print("ERROR: fondo no encontrado en la planilla", file=sys.stderr)
sys.exit(1)
'@ | Set-Content -Path $pyFile -Encoding utf8

$result = python3 $pyFile $tmp 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "[CAFCI] Fallo: $result"
    exit 1
}

$json = ($result | Where-Object { $_ -match '^\{' }) | Select-Object -First 1
$data = $json | ConvertFrom-Json

$ts = Get-Date -Format "yyyy-MM-dd HH:mm"
$js = "// Auto-generado por update_cafci.ps1 — $ts`n// $($data.nombre)`nwindow.CAFCI_DATA = $json;"

Set-Content -Path $output -Value $js -Encoding utf8

Write-Host "[CAFCI] OK" -ForegroundColor Green
Write-Host "  Fondo : $($data.nombre)"
Write-Host "  Fecha : $($data.fecha)"
Write-Host "  VCP   : $($data.vcp)"
Write-Host "  Dia   : $($data.dia)%"
Write-Host "  YTD   : $($data.ytd)%"
Write-Host "  12M   : $($data.m12)%"
Write-Host "  Arch. : $output"
