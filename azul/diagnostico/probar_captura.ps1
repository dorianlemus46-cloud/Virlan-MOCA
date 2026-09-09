# Prueba del C# de captura.ps1. NO toca Azul: trabaja sobre bitmaps sinteticos, asi que se
# puede correr siempre, en cualquier momento, sin la aplicacion abierta.
#
# Existe por la TRAMPA 14 (ver docs/trampas.md): PowerShell REDONDEA al castear a [int] y C#
# TRUNCA. [int]9.6 da 10; (int)9.6 da 9. Al portar el muestreo de Test-CapturasDistintas
# (PowerShell, GetPixel) a [AzulShot]::Diferencia (C#, LockBits), traducir el cast movio la
# rejilla a otros puntos -- en silencio, y dando resultados PARECIDOS, que es lo peor.
#
# La comprobacion que de verdad importa es la 2: los dos muestreos, el viejo y el nuevo,
# tienen que dar el MISMO numero sobre el mismo par de imagenes. Si alguien vuelve a tocar
# Diferencia y la rejilla se desplaza, esa comprobacion lo caza.
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\probar_captura.ps1"
#
# NOTA: ASCII puro, igual que el resto de los .ps1 del proyecto.

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'captura.ps1')
"captura.ps1 cargado y su C# compilado: OK"

# El muestreo VIEJO, copiado tal cual de Test-CapturasDistintas, para comparar contra el nuevo.
# No se borra aunque Wait-PantallaEstable ya no lo use: es el patron de referencia.
function Get-FraccionVieja {
  param($A, $B, [int]$Rejilla = 40)
  if ($A.Width -ne $B.Width -or $A.Height -ne $B.Height) { return 1.0 }
  $dif = 0; $total = $Rejilla * $Rejilla
  for ($i = 0; $i -lt $Rejilla; $i++) {
    for ($j = 0; $j -lt $Rejilla; $j++) {
      $x = [int](($i + 0.5) * $A.Width  / $Rejilla)
      $y = [int](($j + 0.5) * $A.Height / $Rejilla)
      if ($x -ge $A.Width)  { $x = $A.Width  - 1 }
      if ($y -ge $A.Height) { $y = $A.Height - 1 }
      if ($A.GetPixel($x,$y).ToArgb() -ne $B.GetPixel($x,$y).ToArgb()) { $dif++ }
    }
  }
  return ($dif / $total)
}

# Bloques de colores al azar, para que la rejilla caiga sobre pixeles variados y no sobre un
# liso donde cualquier desplazamiento daria el mismo resultado por casualidad.
function New-ImagenPrueba {
  param([int]$Ancho = 1366, [int]$Alto = 768, [int]$Semilla = 1)
  $bmp = New-Object System.Drawing.Bitmap $Ancho, $Alto
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::FromArgb(255, 240, 240, 235))
  $rnd = New-Object System.Random $Semilla
  for ($k = 0; $k -lt 400; $k++) {
    $col = [System.Drawing.Color]::FromArgb(255, $rnd.Next(256), $rnd.Next(256), $rnd.Next(256))
    $br = New-Object System.Drawing.SolidBrush $col
    $g.FillRectangle($br, $rnd.Next($Ancho), $rnd.Next($Alto), $rnd.Next(80) + 5, $rnd.Next(60) + 5)
    $br.Dispose()
  }
  $g.Dispose()
  return $bmp
}

$script:fallos = 0
function Test-Igual {
  param([string]$Que, $Esperado, $Obtenido)
  if ([Math]::Abs([double]$Esperado - [double]$Obtenido) -lt 1e-9) {
    "  OK    $Que  (=$Obtenido)"
  } else {
    "  FALLA $Que  esperado=$Esperado obtenido=$Obtenido"
    $script:fallos++
  }
}

$a = New-ImagenPrueba -Semilla 7

""
"--- 1. dos imagenes identicas dan 0 ---"
$b = New-Object System.Drawing.Bitmap $a
Test-Igual "identicas" 0.0 ([AzulShot]::Diferencia($a, $b, 40))
$b.Dispose()

""
"--- 2. LA QUE IMPORTA: el muestreo nuevo da lo mismo que el viejo ---"
$b = New-ImagenPrueba -Semilla 99
$vieja = Get-FraccionVieja $a $b
$nueva = [AzulShot]::Diferencia($a, $b, 40)
"  viejo (GetPixel) = $vieja"
"  nuevo (LockBits) = $nueva"
Test-Igual "misma rejilla, mismo resultado" $vieja $nueva
$b.Dispose()

""
"--- 3. un solo punto de la rejilla cambiado: el caso del umbral 0.0006 ---"
# El primer punto que muestrea la rejilla. Se calcula con la regla de PowerShell -- [int]
# redondea -- que es la que Diferencia tiene que estar replicando.
$b = New-Object System.Drawing.Bitmap $a
$x = [int]((0 + 0.5) * $a.Width  / 40)
$y = [int]((0 + 0.5) * $a.Height / 40)
"  punto muestreado por la rejilla: ($x,$y)   -- si Diferencia truncara, miraria ($x,$([int][Math]::Truncate((0 + 0.5) * $a.Height / 40)))"
$viejoColor = $b.GetPixel($x, $y)
$nuevoColor = [System.Drawing.Color]::FromArgb(255, (255 - $viejoColor.R), (255 - $viejoColor.G), (255 - $viejoColor.B))
$b.SetPixel($x, $y, $nuevoColor)
$nueva = [AzulShot]::Diferencia($a, $b, 40)
Test-Igual "un punto de 1600" (1/1600) $nueva
Test-Igual "coincide con el viejo" (Get-FraccionVieja $a $b) $nueva
if ($nueva -ge 0.0006) { "  OK    cruza el umbral 0.0006: cuenta como 'sigue pintando'" }
else { "  FALLA no cruza el umbral 0.0006"; $script:fallos++ }
$b.Dispose()

""
"--- 4. tamanos distintos = 1.0 (el 'cambio todo' que antes era return true) ---"
$b = New-ImagenPrueba -Ancho 1000 -Alto 700 -Semilla 3
Test-Igual "tamanos distintos" 1.0 ([AzulShot]::Diferencia($a, $b, 40))
$b.Dispose()

""
"--- 5. nulos y rejilla invalida no revientan ---"
Test-Igual "segunda imagen nula" 1.0 ([AzulShot]::Diferencia($a, $null, 40))
Test-Igual "rejilla 0" 0.0 ([AzulShot]::Diferencia($a, $a, 0))

""
"--- 6. los bitmaps siguen usables despues de LockBits ---"
# Si Diferencia se saltara el UnlockBits, el bitmap quedaria bloqueado y la vuelta siguiente
# del bucle de Wait-PantallaEstable fallaria.
$b = New-Object System.Drawing.Bitmap $a
[void][AzulShot]::Diferencia($a, $b, 40)
try {
  [void]$a.GetPixel(5,5); [void]$b.GetPixel(5,5)
  "  OK    ambos siguen legibles tras UnlockBits"
} catch {
  "  FALLA quedaron bloqueados: $($_.Exception.Message)"; $script:fallos++
}
$b.Dispose()
$a.Dispose()

""
"--- 7. Wait-PantallaEstable: firma y garantia de estabilidad ---"
$cmd = Get-Command Wait-PantallaEstable
$comunes = [System.Management.Automation.PSCmdlet]::CommonParameters
$ps = $cmd.Parameters.Keys | Where-Object { $_ -notin $comunes }
"  parametros: $($ps -join ', ')"
if ($ps -contains 'Temporal') { "  FALLA sigue pidiendo -Temporal (ya no usa disco)"; $script:fallos++ }
else { "  OK    ya no pide -Temporal" }
# La garantia que importa no es el intervalo ni las confirmaciones por separado, sino su
# producto: cuanto tiempo seguido tiene que estar quieta la pantalla para darla por estable.
$vIntervalo = 250; $vConfirmaciones = 6
$src = Get-Content (Join-Path (Split-Path $PSScriptRoot -Parent) 'captura.ps1') -Raw
if ($src -match '\[int\]\$Intervalo\s*=\s*(\d+)')      { $vIntervalo = [int]$Matches[1] }
if ($src -match '\[int\]\$Confirmaciones\s*=\s*(\d+)') { $vConfirmaciones = [int]$Matches[1] }
$garantia = $vIntervalo * $vConfirmaciones
"  intervalo=$vIntervalo ms x confirmaciones=$vConfirmaciones = $garantia ms de pantalla quieta"
if ($garantia -ge 1400) { "  OK    la garantia no baja de los 1400 ms que daba 700x2" }
else { "  FALLA la garantia bajo a $garantia ms; con 700x2 eran 1400"; $script:fallos++ }

""
if ($script:fallos -eq 0) { "TODO OK" } else { "$($script:fallos) comprobacion(es) fallaron"; exit 1 }
