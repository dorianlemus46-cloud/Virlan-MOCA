# Prueba de la MEMORIA de las bases (bases.ps1). NO toca Azul, NI Word, NI la carpeta real.
#
# Trabaja sobre una carpeta temporal propia y con documentos de mentira: archivos de texto con
# nombre de .docx. Puede hacerlo justamente porque bases.ps1 no abre documentos -- solo mira si
# el archivo esta o no esta. Esa separacion es lo que hace que esto se pueda probar.
#
# Comprueba las cinco reglas que sostienen el correlativo, que son las que fallaron la vez
# pasada y por las que el sistema estuvo tres dias sin nombrar bases:
#
#   1. el numero se aparta ANTES de que exista el documento
#   2. el contador nunca retrocede, ni borrando archivos ni moviendolos
#   3. jamas se sobrescribe un documento que ya exista
#   4. el registro se escribe entero o no se escribe, y si esta roto se aborta
#   5. si la base abierta desaparecio del disco, se planta y avisa
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\probar_bases.ps1"
#   ...\probar_bases.ps1 -Conservar     deja la carpeta temporal para mirarla
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

param([switch]$Conservar)
$ErrorActionPreference = 'Stop'

$raiz = Split-Path $PSScriptRoot -Parent
. (Join-Path $raiz 'bases.ps1')

$carpeta = Join-Path $env:TEMP ("probar_bases_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
$fallos  = 0
$dichos  = New-Object System.Collections.ArrayList
$log     = { param($m) [void]$dichos.Add([string]$m) }

function Ok {
  param([string]$Que, [bool]$Cond)
  if ($Cond) { "  OK     $Que" } else { "  FALLA  $Que"; $script:fallos++ }
}
function New-DocFalso {
  param([string]$Ruta)
  Set-Content -LiteralPath $Ruta -Value 'documento de mentira para la prueba' -Encoding ASCII
}

"Carpeta de prueba: $carpeta"
""

try {
  New-Item -ItemType Directory -Path $carpeta -Force | Out-Null

  # ---- 1. una carpeta sin memoria -------------------------------------------
  "--- carpeta nueva ---"
  $reg = Get-RegistroBases -Carpeta $carpeta
  Ok "sin registro, el siguiente numero es el 34" ((([int]$reg.ultimoNumero) + 1) -eq 34)
  Ok "sin registro, no hay ninguna base abierta" ($null -eq $reg.abierta)
  # Leer no debe crear nada: un registro en disco significa "aqui hubo una base".
  Ok "leer la memoria no crea el archivo" (-not (Test-Path -LiteralPath (Get-RutaRegistro -Carpeta $carpeta)))

  # ---- 2. regla 1: el numero se aparta antes que el documento ---------------
  "--- regla 1: el numero se aparta antes de crear el documento ---"
  $b1 = New-BaseAbierta -Carpeta $carpeta -Registro $reg -Log $log
  Ok "el nombre que toca es BASE 034 PS1 VIRLAN.docx" ($b1.Archivo -eq 'BASE 034 PS1 VIRLAN.docx')
  Ok "el registro YA esta escrito en disco" (Test-Path -LiteralPath (Get-RutaRegistro -Carpeta $carpeta))
  Ok "y el documento TODAVIA no existe" (-not (Test-Path -LiteralPath $b1.Ruta))
  $rel = Get-RegistroBases -Carpeta $carpeta
  Ok "releida del disco, la 034 sigue abierta" ((([int]$rel.abierta.numero)) -eq 34)

  # ---- 3. regla 5: la base abierta tiene que estar en el disco --------------
  "--- regla 5: base abierta que desaparecio ---"
  $t = Test-BaseAbierta -Carpeta $carpeta -Registro $rel
  Ok "base abierta sin documento: se detecta el problema" ($t.Hay -and (-not $t.Existe) -and $t.Problema -ne "")
  New-DocFalso $b1.Ruta
  $t2 = Test-BaseAbierta -Carpeta $carpeta -Registro $rel
  Ok "con el documento en su sitio, sin problema" ($t2.Hay -and $t2.Existe -and $t2.Problema -eq "")

  # ---- 4. clientes dentro de la base ----------------------------------------
  "--- los clientes que van dentro ---"
  # Razon social con acento, escrita por su codigo porque este archivo es ASCII puro. Es el
  # viaje que de verdad importa: el nombre acaba en el encabezado del documento del cliente.
  $razon = "DISTRIBUIDORA R" + [char]0x00CD + "O"
  $n1 = Add-ClienteABase -Carpeta $carpeta -Registro $rel -Cuenta '595000001' -Razon $razon
  $n2 = Add-ClienteABase -Carpeta $carpeta -Registro $rel -Cuenta '507000004' -Razon 'CONSULTORES EN SERVICIOS'
  Ok "el primer cliente deja la base en 1" ($n1 -eq 1)
  Ok "el segundo la deja en 2"             ($n2 -eq 2)
  $rel2 = Get-RegistroBases -Carpeta $carpeta
  Ok "releida, la base sigue teniendo 2 clientes" (@($rel2.abierta.clientes).Count -eq 2)
  Ok "la razon social con acento sobrevive al viaje" ((@($rel2.abierta.clientes)[0].razon) -eq $razon)
  Ok "reconoce un cliente ya escrito"  ($null -ne (Find-ClienteEnBases -Registro $rel2 -Cuenta '507000004'))
  Ok "y no inventa uno que no esta"    ($null -eq (Find-ClienteEnBases -Registro $rel2 -Cuenta '999999999'))

  # ---- 5. no quedan temporales sueltos --------------------------------------
  "--- regla 4: escribir entero o no escribir ---"
  Ok "no queda ningun temporal .nuevo suelto" (@(Get-ChildItem -LiteralPath $carpeta -Filter '*.nuevo' -ErrorAction SilentlyContinue).Count -eq 0)

  # ---- 6. regla 3: jamas se sobrescribe -------------------------------------
  # MEDIDO AQUI, 09/09/2026: con la regla 2 puesta, la 3 no llega a dispararse. El barrido de
  # la carpeta ve el 035, el contador salta por encima, y el nombre que sale ya esta libre. La
  # regla 3 es la red de debajo, no el mecanismo -- y esta prueba comprueba el EFECTO, que es
  # lo unico que importa: el documento que ya existia no se toca.
  "--- regla 3: jamas se sobrescribe ---"
  [void](Close-BaseAbierta -Carpeta $carpeta -Registro $rel2 -Motivo 'prueba')
  $ocupado = Join-Path $carpeta 'BASE 035 PS1 VIRLAN.docx'
  New-DocFalso $ocupado                                          # alguien ocupo el 035
  $antes = Get-Content -LiteralPath $ocupado -Raw
  $dichos.Clear()
  $reg3 = Get-RegistroBases -Carpeta $carpeta
  $b2 = New-BaseAbierta -Carpeta $carpeta -Registro $reg3 -Log $log
  Ok "con el 035 ocupado, se salta al 036" ($b2.Numero -eq 36)
  Ok "el documento del 035 queda intacto"  ((Get-Content -LiteralPath $ocupado -Raw) -eq $antes)
  Ok "y avisa de que la carpeta iba por delante" ((($dichos -join ' ') -like '*nunca retrocede*'))

  # ---- 7. regla 2: el contador nunca retrocede ------------------------------
  "--- regla 2: el contador nunca retrocede ---"
  [void](Close-BaseAbierta -Carpeta $carpeta -Registro $reg3 -Motivo 'prueba')
  # Se borran TODOS los documentos: es el caso que rompio el sistema la vez pasada, cuando la
  # carpeta era la unica memoria y renombrar un documento liberaba su numero.
  Get-ChildItem -LiteralPath $carpeta -Filter '*.docx' | Remove-Item -Force
  $reg4 = Get-RegistroBases -Carpeta $carpeta
  $b3 = New-BaseAbierta -Carpeta $carpeta -Registro $reg4 -Log $log
  Ok "con la carpeta vacia de documentos, el numero SIGUE subiendo (037)" ($b3.Numero -eq 37)

  [void](Close-BaseAbierta -Carpeta $carpeta -Registro $reg4 -Motivo 'prueba')
  New-DocFalso (Join-Path $carpeta 'BASE 050 PS1 VIRLAN.docx')   # la carpeta se adelanta
  $dichos.Clear()
  $reg5 = Get-RegistroBases -Carpeta $carpeta
  $b4 = New-BaseAbierta -Carpeta $carpeta -Registro $reg5 -Log $log
  Ok "si la carpeta va por delante, gana la carpeta (051)" ($b4.Numero -eq 51)
  Ok "y tambien lo dice" ((($dichos -join ' ') -like '*nunca retrocede*'))

  # ---- 8. regla 4: un registro roto aborta, no se reinicia ------------------
  "--- regla 4: un registro roto aborta ---"
  [void](Close-BaseAbierta -Carpeta $carpeta -Registro $reg5 -Motivo 'prueba')
  $rutaReg = Get-RutaRegistro -Carpeta $carpeta
  $bueno   = Get-Content -LiteralPath $rutaReg -Raw -Encoding UTF8

  Set-Content -LiteralPath $rutaReg -Value '{esto no es json' -Encoding UTF8
  $abortoRoto = $false
  try { [void](Get-RegistroBases -Carpeta $carpeta) } catch { $abortoRoto = $true }
  Ok "un registro ilegible aborta en vez de reiniciarse" $abortoRoto

  Set-Content -LiteralPath $rutaReg -Value '' -Encoding UTF8
  $abortoVacio = $false
  try { [void](Get-RegistroBases -Carpeta $carpeta) } catch { $abortoVacio = $true }
  Ok "un registro vacio aborta tambien" $abortoVacio

  # Se devuelve el bueno para dejar la carpeta coherente si se conserva.
  Set-Content -LiteralPath $rutaReg -Value $bueno -Encoding UTF8 -NoNewline
  $reg6 = Get-RegistroBases -Carpeta $carpeta
  Ok "restaurado el registro bueno, se vuelve a leer" ((([int]$reg6.ultimoNumero)) -eq 51)
  Ok "y recuerda las 4 bases cerradas" (@($reg6.cerradas).Count -eq 4)
}
finally {
  if ($Conservar) { ""; "Carpeta conservada: $carpeta" }
  elseif (Test-Path -LiteralPath $carpeta) { Remove-Item -LiteralPath $carpeta -Recurse -Force }
}

""
if ($fallos -eq 0) { "TODO OK" } else { "$fallos comprobacion(es) fallaron"; exit 1 }
