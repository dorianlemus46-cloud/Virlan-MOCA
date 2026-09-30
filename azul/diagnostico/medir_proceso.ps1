# Mide el proceso jp2launcher (Azul) desde fuera, sin tocar el puente ni la pantalla, para ver
# si la memoria sube de forma sostenida durante una corrida o si se mantiene plana mientras la
# CPU se dispara. Analisis de rendimiento pedido por Dorian, 29/09/2026 (ver el plan guardado
# ese dia): antes de tocar el heap de Azul hace falta saber si la degradacion documentada en
# ESTADO.md y en lote.ps1 (Azul murio a los 32 min con 1.7 GB, 17/09/2026) es de memoria o de
# otra cosa.
#
# Solo llama a Get-Process -Name jp2launcher, el mismo patron que ya usan captura.ps1, lote.ps1
# y buscar_cuenta.ps1. No abre el puente de accesibilidad, no hace clics, no lee datos de
# cliente: no hay ningun riesgo para Azul ni para REGLAS.md.
#
# Uso, en paralelo con una corrida normal de lote.ps1, en OTRA ventana de PowerShell:
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\medir_proceso.ps1"
#   ...\medir_proceso.ps1 -IntervaloSeg 10 -MinutosTope 60
#   ...\medir_proceso.ps1 -Salida "C:\ruta\propia.csv"
#
# Termina solo si jp2launcher desaparece (Azul se cerro o se cayo) o al llegar a -MinutosTope.
# Si no, hay que pararlo con Ctrl+C -- el CSV ya escrito hasta ese punto queda intacto.
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

param(
  [int]$IntervaloSeg = 15,   # cada cuanto se toma una muestra
  [int]$MinutosTope = 0,     # 0 = sin tope, corre hasta que Azul desaparezca o se interrumpa a mano
  [string]$Salida = ""       # CSV de salida; vacio = azul\salidas\medir_proceso_<fecha_hora>.csv
)
$ErrorActionPreference = 'Stop'

$raiz = Split-Path $PSScriptRoot -Parent
if ($Salida -eq "") {
  $Salida = Join-Path $raiz ("salidas\medir_proceso_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv")
}
$dir = Split-Path $Salida -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

'hora,segundos_desde_inicio,working_set_mb,memoria_privada_mb,cpu_acumulada_seg,hilos,manejadores' |
  Out-File -FilePath $Salida -Encoding utf8

Write-Host "Midiendo jp2launcher cada $IntervaloSeg s."
Write-Host "Salida: $Salida"
Write-Host "Ctrl+C para parar a mano. Si Azul se cierra o se cae, este script lo nota solo y termina."
Write-Host ""

$inicio = Get-Date
while ($true) {
  $p = Get-Process -Name jp2launcher -ErrorAction SilentlyContinue |
       Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if (-not $p) {
    Write-Host "jp2launcher no esta (cerrado o caido). Fin de la medicion."
    break
  }

  $seg   = [Math]::Round(((Get-Date) - $inicio).TotalSeconds, 1)
  $wsMb  = [Math]::Round($p.WorkingSet64 / 1MB, 1)
  $pmMb  = [Math]::Round($p.PrivateMemorySize64 / 1MB, 1)
  $cpu   = [Math]::Round($p.CPU, 1)
  $hilos = $p.Threads.Count
  $mane  = $p.HandleCount
  $hora  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

  "$hora,$seg,$wsMb,$pmMb,$cpu,$hilos,$mane" | Out-File -FilePath $Salida -Encoding utf8 -Append
  Write-Host ("  {0}  +{1,7}s  WS={2,7} MB  Privada={3,7} MB  CPU={4,7} s  hilos={5,4}  manejadores={6,5}" -f `
    $hora, $seg, $wsMb, $pmMb, $cpu, $hilos, $mane)

  if ($MinutosTope -gt 0 -and $seg -ge ($MinutosTope * 60)) {
    Write-Host "Tope de $MinutosTope minutos alcanzado. Fin de la medicion."
    break
  }
  Start-Sleep -Seconds $IntervaloSeg
}

Write-Host ""
Write-Host "CSV: $Salida"
