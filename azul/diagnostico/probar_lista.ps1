# Pruebas del lector de la lista de ordenes (azul\lista.ps1).
# NO toca Azul, ni Excel, ni Word: escribe CSV sinteticos en la carpeta temporal y los lee.
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\probar_lista.ps1"
#
# Existe porque este lector es la unica puerta por la que entran las cuentas, y de el sale la
# razon social, que es lo que sostiene la guarda contra entrar al cliente equivocado. Un fallo
# aqui no se ve: se ve tres pasos mas adelante, en el documento.
#
# LOS ACENTOS SE CONSTRUYEN DESDE SU CODIGO DE CARACTER, no se escriben. Este archivo es ASCII
# puro (trampa 1) y una prueba de codificacion escrita con literales acentuados no vale nada:
# PowerShell 5.1 lee el .ps1 como ANSI y rompe POR IGUAL el valor esperado y el escrito, asi
# que la comparacion pasa comparando basura contra la misma basura. Paso el 05/09/2026 al
# escribir la primera version de esta prueba, y dio un OK falso.
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'lista.ps1')
"lista.ps1 cargado: OK"

$tmp    = $env:TEMP
$fallos = 0
$mudo   = { param($m) }          # para las pruebas donde los avisos serian ruido

$I = [char]0xCD                  # I con tilde
$o = [char]0xF3                  # o con tilde
$RAZON = "DISTRIBUIDORA R" + $I + "O"
$ENCAB = "Raz" + $o + "n Social"

function Escribe {
  param([string]$Nombre, [string]$Texto, $Codificacion)
  $ruta = Join-Path $tmp $Nombre
  [IO.File]::WriteAllText($ruta, $Texto, $Codificacion)
  return $ruta
}
function Ok   { param([string]$M) "  OK    $M" }
function Falla{ param([string]$M) "  FALLA $M"; $script:fallos++ }

# Nueve columnas como las de resultData, con la cuenta y la razon social en medio: el lector
# tiene que encontrarlas por encabezado y no por posicion.
$CUERPO = @(
  "Folio,Consecutivo,Cuenta,$ENCAB,Tipo,Estatus,Fecha,Creado por,Tecnologia"
  "F1,1,595000001,$RAZON,REN,PEND,01/09/2026,DL,4G"
  "F2,2,508000001,ACME SA DE CV,REN,PEND,01/09/2026,DL,4G"
  "F3,3,595000001,$RAZON,REN,PEND,02/09/2026,DL,4G"
  "F4,4,5.95E+08,CUENTA ROTA SA,REN,PEND,02/09/2026,DL,4G"
  "F5,5,,SIN CUENTA SA DE CV,REN,PEND,02/09/2026,DL,4G"
  "F6,6,507000002,CONSULTORES EN SERVICIOS,REN,PEND,02/09/2026,DL,4G"
  ",,,,,,,,"
) -join "`r`n"

$ANSI    = [Text.Encoding]::GetEncoding(1252)
$CONBOM  = New-Object Text.UTF8Encoding($true)
$SINBOM  = New-Object Text.UTF8Encoding($false)

# ---- 1. lo que hace con una lista normal ------------------------------------
""
"--- 1. filas repetidas y filas malas ---"
$r = Escribe 'pl_normal.csv' $CUERPO $ANSI
$l = Get-ListaOrdenes -Ruta $r
if ($l.Count -eq 3) { Ok "3 clientes: 1 repetida y 2 filas malas quedaron fuera" }
else { Falla "salieron $($l.Count) clientes, se esperaban 3" }
if (($l | Where-Object { $_.Fila -eq 2 }).Cuenta -eq '595000001') { Ok "la numeracion de fila es la de la hoja (1 = encabezados)" }
else { Falla "la fila 2 no es la primera de datos" }

# ---- 2. las tres codificaciones ---------------------------------------------
# ANSI y UTF-8 con BOM son lo que exporta Excel. UTF-8 sin BOM no lo exporta, pero es lo que
# guarda el Bloc de notas de Windows 11, y abrir el CSV para revisarlo es algo que se hace.
""
"--- 2. codificacion: la razon social tiene que llegar intacta ---"
foreach ($c in @(
    @{ N = 'pl_ansi.csv';   E = $ANSI;   Que = 'ANSI (CSV delimitado por comas)' },
    @{ N = 'pl_bom.csv';    E = $CONBOM; Que = 'UTF-8 con BOM (CSV UTF-8)' },
    @{ N = 'pl_nobom.csv';  E = $SINBOM; Que = 'UTF-8 SIN BOM (Bloc de notas)' })) {
  $ru = Escribe $c.N $CUERPO $c.E
  $li = Get-ListaOrdenes -Ruta $ru -Log $mudo
  $fi = $li | Where-Object { $_.Cuenta -eq '595000001' }
  if ($null -eq $fi)              { Falla "$($c.Que): no se reconocio el encabezado con acento" }
  elseif ($fi.Razon -eq $RAZON)   { Ok    "$($c.Que): razon social intacta" }
  else                            { Falla "$($c.Que): llego '$($fi.Razon)' en vez de '$RAZON'" }
}

# ---- 3. separador punto y coma ----------------------------------------------
""
"--- 3. separador ---"
$r = Escribe 'pl_pyc.csv' ($CUERPO -replace ',', ';') $ANSI
$l = Get-ListaOrdenes -Ruta $r -Log $mudo
if ($l.Count -eq 3) { Ok "con ';' salen los mismos 3 clientes" }
else { Falla "con ';' salieron $($l.Count)" }

# ---- 4. columnas que faltan --------------------------------------------------
""
"--- 4. columnas que faltan ---"
$r = Escribe 'pl_sinrazon.csv' "Folio,Cuenta,Tipo`r`nF1,595000001,REN" $ANSI
$avisos = @()
$l = Get-ListaOrdenes -Ruta $r -Log { param($m) $script:avisos += $m }
if ($avisos -join ' ' -like '*NO se aplicara*') { Ok "sin 'Razon Social' avisa de que la guarda queda apagada" }
else { Falla "sin 'Razon Social' no aviso" }

$r = Escribe 'pl_sincuenta.csv' "Folio,Nombre`r`nF1,ACME" $ANSI
try { $null = Get-ListaOrdenes -Ruta $r -Log $mudo; Falla "sin 'Cuenta' no lanzo" }
catch {
  if ($_.Exception.Message -like "*no tiene una columna 'Cuenta'*") { Ok "sin 'Cuenta' lanza y ensena las columnas que vio" }
  else { Falla "sin 'Cuenta' lanzo otro error: $($_.Exception.Message)" }
}

# ---- 5. Get-OrdenDeLista -----------------------------------------------------
""
"--- 5. una sola fila ---"
$r = Escribe 'pl_normal.csv' $CUERPO $ANSI
$u = Get-OrdenDeLista -Ruta $r -Fila 3
if ($u.Cuenta -eq '508000001' -and $u.Razon -eq 'ACME SA DE CV') { Ok "-Fila 3 devuelve la tercera fila de la hoja" }
else { Falla "-Fila 3 devolvio '$($u.Cuenta)' / '$($u.Razon)'" }

foreach ($f in @(4, 5, 6, 99)) {
  try { $null = Get-OrdenDeLista -Ruta $r -Fila $f; Falla "-Fila $f devolvio algo y no debia" }
  catch { Ok "-Fila $f (repetida, rota, sin cuenta o inexistente) se niega" }
}

Remove-Item (Join-Path $tmp 'pl_*.csv') -Force -ErrorAction SilentlyContinue
""
if ($fallos -eq 0) { "TODO OK" } else { "$fallos comprobacion(es) fallaron"; exit 1 }
