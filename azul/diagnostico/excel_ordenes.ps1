# Sonda de SOLO LECTURA sobre el Excel ABIERTO: la lista de ordenes.
#
# Para que sirve: antes de leer numeros de cuenta hay que saber QUE libro, QUE hoja y QUE
# columnas. Este script no decide nada: muestra lo que hay para que Dorian confirme. La
# leccion del 02/09/2026 con Word es justamente esa -- escribir a ciegas en "el activo"
# termino en una copia de autorrecuperacion, y con datos validos habria sido el cliente
# equivocado sin decir nada. Aqui se listan TODOS los libros abiertos, no solo el activo.
#
# No escribe, no guarda, no cierra y NUNCA llama a Quit(): el Excel es de Dorian y se queda
# exactamente como estaba.
#
#   .\excel_ordenes.ps1
#   .\excel_ordenes.ps1 -Libro "ordenes.xlsx" -Hoja "Hoja1" -Filas 15
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

param(
  # Vacio = todos los libros abiertos. Acepta parte del nombre.
  [string]$Libro = "",
  # Vacio = todas las hojas del libro.
  [string]$Hoja = "",
  [int]$Filas = 8,
  [int]$Columnas = 15
)
$ErrorActionPreference = 'Stop'

try {
  $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
  throw "No hay un Excel abierto al que conectarme. Abre la lista de ordenes y vuelve a correr."
}
if ($xl.Workbooks.Count -eq 0) { throw "Excel esta abierto pero sin libros." }

# A, B, ... Z, AA, AB...  Hace falta para poder decir "la cuenta esta en la columna D" en vez
# de "en la cuarta columna del rango usado", que no es lo mismo si el rango no empieza en A.
function Letra([int]$n) {
  $s = ""
  while ($n -gt 0) {
    $m = ($n - 1) % 26
    $s = [char](65 + $m) + $s
    $n = [int](($n - $m - 1) / 26)
  }
  return $s
}

# Un valor "parece cuenta" si, quitados espacios y guiones, es solo digitos y tiene largo de
# numero de cuenta. Es una PISTA para el informe, no una decision: la columna la confirma
# Dorian.
function Parece-Cuenta([string]$v) {
  $t = ($v -replace '[\s\-]', '')
  if ($t.Length -lt 6 -or $t.Length -gt 12) { return $false }
  return ($t -match '^\d+$')
}

$activo = $null
try { $activo = $xl.ActiveWorkbook.Name } catch { }

Write-Host "Excel: $($xl.Workbooks.Count) libro(s) abierto(s).  Activo: $activo"
Write-Host ""

foreach ($wb in $xl.Workbooks) {
  if ($Libro -ne "" -and $wb.Name -notlike "*$Libro*") { continue }

  $marca = if ($wb.Name -eq $activo) { "  <-- ACTIVO" } else { "" }
  Write-Host "=============================================================="
  Write-Host "LIBRO: $($wb.Name)$marca"
  try { Write-Host "  Ruta: $($wb.Path)" } catch { }
  Write-Host "  Hojas: $($wb.Worksheets.Count)"

  foreach ($ws in $wb.Worksheets) {
    if ($Hoja -ne "" -and $ws.Name -notlike "*$Hoja*") { continue }

    $ur = $null
    try { $ur = $ws.UsedRange } catch { }
    if ($null -eq $ur) { Write-Host "  --- HOJA: $($ws.Name)  (sin rango usado)"; continue }

    $r0 = $ur.Row; $c0 = $ur.Column
    $nr = $ur.Rows.Count; $nc = $ur.Columns.Count
    Write-Host ""
    Write-Host "  --- HOJA: $($ws.Name)   rango usado: $($ur.Address($false,$false))   ${nr} filas x ${nc} columnas"

    $maxF = [Math]::Min($Filas, $nr)
    $maxC = [Math]::Min($Columnas, $nc)

    # Cabecera de columnas con su LETRA real de hoja
    $cab = "      fila |"
    for ($c = 0; $c -lt $maxC; $c++) { $cab += (" {0,-24} |" -f (Letra ($c0 + $c))) }
    Write-Host $cab

    # Se usa .Text (lo que se VE en la celda) y no .Value2: un numero de cuenta puede venir
    # formateado, y lo que hay que teclear en Azul es lo que Dorian ve, no el double crudo.
    $pistas = @{}
    for ($f = 0; $f -lt $maxF; $f++) {
      $ln = ("      {0,4} |" -f ($r0 + $f))
      for ($c = 0; $c -lt $maxC; $c++) {
        $v = ""
        try { $v = [string]$ws.Cells($r0 + $f, $c0 + $c).Text } catch { }
        if ($null -eq $v) { $v = "" }
        $v = $v.Trim()
        if ($f -gt 0 -and (Parece-Cuenta $v)) {
          $k = Letra ($c0 + $c)
          if ($pistas.ContainsKey($k)) { $pistas[$k]++ } else { $pistas[$k] = 1 }
        }
        if ($v.Length -gt 24) { $v = $v.Substring(0, 21) + "..." }
        $ln += (" {0,-24} |" -f $v)
      }
      Write-Host $ln
    }

    if ($pistas.Count -gt 0) {
      $orden = $pistas.GetEnumerator() | Sort-Object Value -Descending
      $txt = ($orden | ForEach-Object { "$($_.Key) ($($_.Value))" }) -join ", "
      Write-Host "      pista: columnas con valores que parecen numero de cuenta -> $txt"
      Write-Host "      (la razon social suele ser la de al lado; confirmalo mirando la tabla de arriba)"
    }
  }
  Write-Host ""
}

Write-Host "=============================================================="
Write-Host "No se escribio, no se guardo y no se cerro nada."
