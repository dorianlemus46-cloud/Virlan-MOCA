# Comprueba si el CSV de una corrida esta COMPLETO, y de paso cuenta sus bloques.
#
# Vivia dentro de azul_fast.ps1, al final. Se saco aqui cuando lote.ps1 paso a decidir el
# volcado a Word por su cuenta (azul_fast corre con -SinWord): con el bloque ahi dentro, la
# guarda dejaba de ejecutarse justo en la corrida desatendida, que es donde mas hace falta.
# Duplicarla habria sido peor que moverla -- dos copias de una guarda acaban divergiendo.
#
# La distincion que importa: una linea AMBIGUA (multiples MPE, fecha ilegible) es un
# RESULTADO legitimo y su cliente si va al documento, en su bloque A REVISAR. Lo que no va es
# una corrida que se ROMPIO a medias.
#
#   $r = & .\revisar_csv.ps1 -Csv ruta.csv
#   $r.Problemas   # lista vacia = corrida completa
#   $r.Renovables / $r.NoRenovables / $r.Revisar / $r.Canceladas
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

param([Parameter(Mandatory=$true)][string]$Csv)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Csv)) { throw "No encuentro el CSV: $Csv" }
$lineas = Get-Content -LiteralPath $Csv -Encoding UTF8

$res = @{
  Problemas    = @()
  Renovables   = -1
  NoRenovables = -1
  Revisar      = 0
  Canceladas   = 0
}

# El cuadre: activas leidas + canceladas excluidas tiene que dar el total de filas de la
# tabla. Si no cuadra, la corrida se dejo lineas por el camino.
$cu = $lineas | Where-Object { $_ -like '#*CUADRE:*' } | Select-Object -First 1
if ($cu -and ($cu -match 'CUADRE:\s*(\d+)\s+activas leidas \+\s*(\d+)\s+canceladas excluidas =\s*(\d+)')) {
  if (([int]$Matches[1] + [int]$Matches[2]) -ne [int]$Matches[3]) {
    $res.Problemas += "el cuadre no cuadra: $($cu.TrimStart('#',' '))"
  }
} else {
  $res.Problemas += "no encuentro la linea de CUADRE"
}

# Marcas de corrida rota, todas nacidas de fallos reales del lector.
foreach ($m in @('corrida abortada', 'la cuenta cambio', 'se agoto el tiempo',
                 'el detalle no aparecio', 'los atributos no cargaron',
                 'no se marco la fila', 'no abrio Ver Productos', 'boton deshabilitado')) {
  if ($lineas | Where-Object { $_ -like "*$m*" }) { $res.Problemas += $m }
}

# Conteo por bloque, leido de los separadores '# NOMBRE (n)'. Es lo que decide si el cliente
# entra al documento: sin nada que renovar y sin nada que revisar, no ensucia la base.
foreach ($ln in $lineas) {
  if ($ln -match '^#\s+RENOVABLES\s*\((\d+)\)')           { $res.Renovables   = [int]$Matches[1] }
  elseif ($ln -match '^#\s+NO RENOVABLES\s*\((\d+)\)')    { $res.NoRenovables = [int]$Matches[1] }
  elseif ($ln -match '^#\s+A REVISAR\s*\((\d+)\)')        { $res.Revisar      = [int]$Matches[1] }
  elseif ($ln -match '^#\s+CANCELADAS EXCLUIDAS\s*\((\d+)\)') { $res.Canceladas = [int]$Matches[1] }
}
# Ojo con el orden: 'NO RENOVABLES' contiene 'RENOVABLES', asi que el primer patron tiene que
# exigir el inicio de linea con '# RENOVABLES'. Por eso van anclados con ^#\s+ y no con -like.

if ($res.Renovables -lt 0) { $res.Problemas += "el CSV no trae el bloque '# RENOVABLES (n)'" }

return $res
