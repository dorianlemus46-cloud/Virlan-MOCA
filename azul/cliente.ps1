# Fusiona buscar_cuenta.ps1 y azul_fast.ps1 en UN SOLO proceso de PowerShell de 32 bits por
# cliente, en vez de los dos que lanza lote.ps1 hoy. Paso 3 del analisis de rendimiento pedido
# por Dorian el 29/09/2026 (ver el plan de esa fecha): cada proceso paga su propio arranque de
# PowerShell y su propia compilacion de C# con Add-Type, y captura.ps1 se compilaba DOS veces,
# una por proceso, aunque ya se protege sola con "-as [type]" -- asi que el ahorro real es un
# arranque de proceso entero, no la compilacion de captura.ps1 en si.
#
# COMPROBADO antes de escribir esto, en un proceso de prueba aparte (no aqui, no contra Azul):
#   - Los tres bloques de C# (AzulShot de captura.ps1, Azul de azul_fast.ps1, Busca de
#     buscar_cuenta.ps1) compilan juntos en el mismo proceso sin chocar: buscar_cuenta.ps1 ya
#     usa ACI2/ATI2/MSG2 en vez de ACI/ATI/MSG, asi que no hay tipos con el mismo nombre.
#   - Los dos scripts no comparten ningun nombre de funcion.
# Add-Type con el MISMO codigo dos veces (captura.ps1, cargado por los dos) no falla: PowerShell
# reconoce que ya esta compilado. NO SE HA CORRIDO NADA DE ESTO CONTRA AZUL REAL.
#
# NO TOCA nada de lo que se lee, se pulsa o se escribe en Azul -- es orquestacion pura, los dos
# scripts siguen siendo los mismos, sin editar una sola linea de su logica.
#
# La UNICA diferencia observable de fuera: buscar_cuenta.ps1 y azul_fast.ps1 usan los dos
# "exit 1" para cualquier error, y aqui corren en el MISMO proceso, asi que el codigo de salida
# por si solo ya no basta para saber si fallo la busqueda o la lectura (antes se sabia por CUAL
# de los dos procesos habia fallado). Por eso este script dejar una marca justo despues de que
# la busqueda termine bien y antes de empezar a leer: si el proceso termina en fallo y la marca
# NO esta, fallo la busqueda; si SI esta, fallo la lectura. lote.ps1 la mira con -FusionarProcesos.
#
# Uso -- SOLO con el PowerShell de 32 bits, igual que los dos scripts que fusiona:
#   $ps32 = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
#   & $ps32 -NoProfile -ExecutionPolicy Bypass -File ".\azul\cliente.ps1" -Cuenta 595000001 -Razon "..."
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

param(
  [string]$Cuenta = "",
  [string]$Razon = "",
  [int]$PausaEscribir = 1,      # buscar_cuenta.ps1: entre escribir la cuenta y pulsar "Buscar Ahora"
  [int]$PausaResultado = 0,     # buscar_cuenta.ps1: entre pulsar "Buscar Ahora" y leer el resultado
  [double]$PausaLineas = 0,     # azul_fast.ps1: entre cerrar el detalle de una linea y la siguiente
  [double]$PausaLectura = 0     # azul_fast.ps1: entre abrir el detalle de una linea y leerlo
)
$ErrorActionPreference = 'Stop'
$raiz = $PSScriptRoot

# La marca que distingue FALLO BUSQUEDA de FALLO LECTURA para quien llama (lote.ps1). Se borra
# al empezar cada cliente para no arrastrar la del cliente anterior.
$dirSalidas = Join-Path $raiz 'salidas'
if (-not (Test-Path $dirSalidas)) { New-Item -ItemType Directory -Path $dirSalidas -Force | Out-Null }
$marca = Join-Path $dirSalidas '.busqueda_ok'
if (Test-Path -LiteralPath $marca) { Remove-Item -LiteralPath $marca -Force }

# ---- 1: buscar la cuenta, entrar al cliente, abrir Suscripciones --------------
# Si esto falla, buscar_cuenta.ps1 ya hizo "exit 1" (o el que sea) mas arriba: las lineas de
# abajo, incluidas azul_fast.ps1 y la marca, NUNCA se alcanzan -- el proceso entero termina
# aqui, con el mismo codigo de salida que daba antes el proceso separado.
. (Join-Path $raiz 'buscar_cuenta.ps1') -Cuenta $Cuenta -Razon $Razon -Restaurar `
    -PausaEscribir $PausaEscribir -PausaResultado $PausaResultado

New-Item -ItemType File -Path $marca -Force | Out-Null

# ---- 2: leer las lineas, clasificar y dejar el CSV -----------------------------
# -SinWord: igual que hoy, quien decide si el cliente entra al documento es lote.ps1, no el
# lector. El resto de los parametros y el contrato entre los dos scripts (el CSV, el cierre de
# la interaccion al final de la lectura) quedan exactamente como estan; aqui no se toca nada
# de eso, solo donde corren.
. (Join-Path $raiz 'azul_fast.ps1') -SinWord -Cuenta $Cuenta -Razon $Razon `
    -PausaLineas $PausaLineas -PausaLectura $PausaLectura
