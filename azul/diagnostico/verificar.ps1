$ErrorActionPreference = 'Stop'
# Comprobaciones que van ANTES de correr nada contra Azul. Las dos primeras existen porque
# ya costaron horas en su momento:
#
#   1. ASCII puro. PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en
#      silencio: "Plan Movil" deja de coincidir y todo se va a REVISAR sin decir por que.
#   2. El C# va dentro de un here-string, asi que un error de sintaxis no aparece hasta la
#      ejecucion. Aqui se compila solo, sin tocar Azul.
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\verificar.ps1"

$raiz = Split-Path $PSScriptRoot -Parent
$archivos = @('azul_fast.ps1', 'agregar_a_word.ps1', 'captura.ps1', 'buscar_cuenta.ps1',
              'revisar_csv.ps1', 'lote.ps1', 'lista.ps1', 'bases.ps1')
# Los diagnosticos no corren en produccion, pero se rompen por lo mismo: acentos que
# PowerShell 5.1 lee como ANSI, y errores de sintaxis que no aparecen hasta ejecutarlos.
$diagnosticos = @('diagnostico\jab_cliente.ps1', 'diagnostico\zoom.ps1', 'diagnostico\excel_ordenes.ps1',
                  'diagnostico\probar_captura.ps1', 'diagnostico\probar_word.ps1',
                  'diagnostico\probar_lista.ps1', 'diagnostico\probar_bases.ps1')
$fallos = 0

"--- ASCII puro ---"
foreach ($a in ($archivos + $diagnosticos)) {
  $ruta = Join-Path $raiz $a
  if (-not (Test-Path $ruta)) { "  $a : NO EXISTE"; $fallos++; continue }
  $n = ([IO.File]::ReadAllBytes($ruta) | Where-Object { $_ -gt 127 }).Count
  if ($n -eq 0) { "  $a : OK" } else { "  $a : $n bytes fuera de ASCII"; $fallos++ }
}

"--- parseo de PowerShell ---"
foreach ($a in ($archivos + $diagnosticos)) {
  $ruta = Join-Path $raiz $a
  if (-not (Test-Path $ruta)) { continue }
  $t = $null; $e = $null
  [void][System.Management.Automation.Language.Parser]::ParseFile($ruta, [ref]$t, [ref]$e)
  if ($e.Count -eq 0) { "  $a : OK" }
  else { "  $a : $($e.Count) errores"; $e | ForEach-Object { "      " + $_.Message }; $fallos++ }
}

"--- compilacion del C# ---"
# captura.ps1 tambien lleva C# -- y quedo sin comprobar hasta el 04/09/2026, que es
# exactamente lo que este archivo existe para evitar. No se buscaba porque su Add-Type usa la
# forma con -ReferencedAssemblies y aqui se buscaba el literal "Add-Type -TypeDefinition @'".
# Ahora el marcador es solo "-TypeDefinition @'", que sirve para las dos formas, y cada
# archivo declara los ensamblados que su C# necesita para compilar.
$conCsharp = @(
  @{ Archivo = 'azul_fast.ps1';     Refs = @() },
  @{ Archivo = 'buscar_cuenta.ps1'; Refs = @() },
  @{ Archivo = 'captura.ps1';       Refs = @('System.Drawing') }
)
$marca = "-TypeDefinition @'"
foreach ($c in $conCsharp) {
  $a = $c.Archivo
  $ruta = Join-Path $raiz $a
  if (-not (Test-Path $ruta)) { "  $a : NO EXISTE"; $fallos++; continue }
  $src = Get-Content $ruta -Raw
  $ini = $src.IndexOf($marca)
  $fin = $src.IndexOf("`n'@", $ini)
  if ($ini -lt 0 -or $fin -lt 0) {
    "  $a : no encontre el bloque de C#"; $fallos++; continue
  }
  $desde = $ini + $marca.Length
  $cs = $src.Substring($desde, $fin - $desde)
  try {
    if ($c.Refs.Count -gt 0) { Add-Type -ReferencedAssemblies $c.Refs -TypeDefinition $cs }
    else                     { Add-Type -TypeDefinition $cs }
    "  $a : OK"
  }
  catch { "  $a : FALLA"; $_.Exception.Message -split "`n" | ForEach-Object { "      $_" }; $fallos++ }
}

""
if ($fallos -eq 0) { "TODO OK" } else { "$fallos comprobacion(es) fallaron"; exit 1 }
