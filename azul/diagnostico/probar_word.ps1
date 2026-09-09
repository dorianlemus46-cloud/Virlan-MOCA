# Prueba del bloque que agregar_a_word.ps1 escribe en el documento. NO toca Azul.
#
# Es la "prueba 5" de docs/plan-simplificacion.md: comprueba que el bloque de un cliente sale
# con la forma que pidio Dorian --  cuenta, razon social, el hueco de la imagen, y despues las
# tablas -- y que las tablas NO se funden (trampa 12).
#
# Trabaja sobre un CSV sintetico que se genera aqui mismo, con numeros inventados: no lee
# ninguna corrida real y no deja datos de clientes en ninguna parte.
#
#   & powershell.exe -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\probar_word.ps1"
#   ...\probar_word.ps1 -Conservar        deja el .docx para mirarlo con los ojos
#
# YA NO SE NIEGA A CORRER CON WORD ABIERTO (09/09/2026). Se negaba porque agregar_a_word.ps1
# escribia en ActiveDocument: con una base real abierta, esta prueba le meteria un cliente
# inventado dentro -- el incidente del 02/09/2026 (docs/bitacora.md, "Se escribio en una copia
# de autorrecuperacion"). Desde que el escritor acepta -DocObj y escribe en EL documento que se
# le da, no hay ActiveDocument que confundir y la guarda dejo de tener sentido.
#
# Lo que no cambia: la prueba levanta su PROPIO Word y su propio documento, y los cierra al
# terminar. Nunca toca lo que haya abierto Dorian.
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.

param(
  # Donde dejar el .docx y el CSV de prueba. Nunca la carpeta real de bases.
  [string]$Carpeta = $env:TEMP,
  # No borrar el .docx al terminar, para abrirlo y mirarlo.
  [switch]$Conservar
)
$ErrorActionPreference = 'Stop'

$raiz    = Split-Path $PSScriptRoot -Parent
$script  = Join-Path $raiz 'agregar_a_word.ps1'
$csv     = Join-Path $Carpeta 'probar_word_sintetico.csv'
$docx    = Join-Path $Carpeta 'probar_word_base.docx'
$fallos  = 0

if (-not (Test-Path $script)) { throw "No encuentro agregar_a_word.ps1 en $raiz" }
if (-not (Test-Path $Carpeta)) { throw "No existe la carpeta de trabajo: $Carpeta" }

# ---- Word abierto: se avisa, ya no se aborta ---------------------------------
$abierto = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($abierto) {
  "AVISO: hay Word abierto ($($abierto.Count) proceso(s)). No estorba: esta prueba usa su"
  "       propio Word y su propio documento, y le dice al escritor exactamente en cual"
  "       escribir. Nada de lo que tengas abierto se toca."
}

# ---- CSV sintetico -----------------------------------------------------------
# Mismo formato que deja azul_fast.ps1: un encabezado, bloques separados por '# NOMBRE (n)',
# y el cuadre al final. SIN linea '# IMAGEN:', que es el caso normal desde el 05/09/2026.
$lineas = @(
  '# CLIENTE: CLIENTE DE PRUEBA SA DE CV'
  'Numero,Plan,Dispositivo,MPE,PlanForzoso,FechaExpiracion'
  '# RENOVABLES (2)'
  '"5510000001","ATT Armalo Negocios 399","Samsung A15","350","24 Meses","10/10/2026"'
  '"5510000002 (SIM)","","N/A","","",""'
  '# NO RENOVABLES (1)'
  '"5510000003","ATT Armalo Negocios 599","Motorola G54","500","36 Meses","01/06/2027"'
  '# A REVISAR (1) - el motivo va junto al numero'
  '"5510000004 (multiples MPE)","","","","",""'
  '# CANCELADAS EXCLUIDAS (1)'
  '"5510000005","","","","",""'
  '#'
  '# CUADRE: 4 activas leidas + 1 canceladas excluidas = 5 filas de la tabla'
  '# 2 renovables, 1 no renovables, 1 a revisar'
)
$lineas | Out-File -FilePath $csv -Encoding utf8
"CSV sintetico: $csv"

if (Test-Path $docx) { Remove-Item $docx -Force }

$w = $null
try {
  $w = New-Object -ComObject Word.Application
  # Visible para que un Word colgado tras un fallo se vea en pantalla en vez de quedarse como
  # un proceso invisible reteniendo el documento. Ya no hace falta para que GetActiveObject lo
  # encuentre: al escritor se le pasa el documento por objeto.
  $w.Visible = $true

  $d = $w.Documents.Add()
  # SaveAs con [ref] exige variables [string]/[int] peladas: [ref] sobre lo que devuelve
  # Join-Path llega envuelto en psobject y COM lo rechaza con "no se puede convertir".
  [string]$rutaDocx = $docx
  [int]$fmt = 16                       # wdFormatDocumentDefault (.docx)
  $d.SaveAs([ref]$rutaDocx, [ref]$fmt)
  $d.Activate()

  # Si Word no esta licenciado, la automatizacion devuelve basura en vez de fallar: el nombre
  # del documento sale vacio y agregar_a_word.ps1 aborta por su guarda de "sin ruta en disco".
  # Mas vale decirlo aqui que dejar que el fallo parezca del codigo.
  if ([string]::IsNullOrWhiteSpace([string]$d.Name)) {
    "ERROR: Word devolvio un documento sin nombre. Suele significar que la instalacion no"
    "       tiene licencia activa: la automatizacion COM no funciona y esta prueba no vale."
    exit 1
  }

  ""
  "=== salida de agregar_a_word.ps1 ==="
  & $script -Csv $csv -Cuenta '595000001' -Razon 'DISTRIBUIDORA RIO' -DocObj $d

  # ---- lo que quedo escrito --------------------------------------------------
  # Los parrafos DENTRO de tablas se saltan: aqui interesa el esqueleto del bloque.
  $sueltos = @()
  foreach ($p in $d.Paragraphs) {
    if ($p.Range.Information(12)) { continue }      # 12 = wdWithInTable
    $sueltos += ($p.Range.Text -replace "[`r`a]", '')
  }

  ""
  "=== parrafos fuera de tabla, en orden ==="
  for ($k = 0; $k -lt $sueltos.Count; $k++) { "{0,2}. '{1}'" -f $k, $sueltos[$k] }

  ""
  "=== tablas ==="
  "total: $($d.Tables.Count)"
  for ($t = 1; $t -le $d.Tables.Count; $t++) {
    $tb = $d.Tables.Item($t)
    "  tabla $t : $($tb.Rows.Count) filas x $($tb.Columns.Count) columnas"
  }

  # ---- comprobaciones --------------------------------------------------------
  ""
  "=== comprobaciones ==="

  $nImg = ($sueltos | Where-Object { $_ -eq '[imagen]' }).Count
  if ($nImg -eq 1) { "  OK    hay exactamente un parrafo '[imagen]'" }
  else { "  FALLA hay $nImg parrafos '[imagen]', se esperaba 1"; $fallos++ }

  $iCta = -1; $iRaz = -1; $iImg = -1; $iRen = -1
  for ($k = 0; $k -lt $sueltos.Count; $k++) {
    if ($iCta -lt 0 -and $sueltos[$k] -eq '595000001')          { $iCta = $k }
    if ($iRaz -lt 0 -and $sueltos[$k] -eq 'DISTRIBUIDORA RIO')  { $iRaz = $k }
    if ($iImg -lt 0 -and $sueltos[$k] -eq '[imagen]')           { $iImg = $k }
    if ($iRen -lt 0 -and $sueltos[$k] -like 'RENOVABLES - *')   { $iRen = $k }
  }
  if ($iCta -ge 0 -and $iRaz -eq ($iCta + 1) -and $iImg -eq ($iRaz + 1) -and $iRen -gt $iImg) {
    "  OK    orden: cuenta -> razon social -> [imagen] -> RENOVABLES"
  } else {
    "  FALLA orden roto: cuenta=$iCta razon=$iRaz imagen=$iImg renovables=$iRen"; $fallos++
  }

  # Trampa 12: Range.Text sobre un parrafo borra la marca de fin y funde las tablas. Cuatro
  # bloques = cuatro tablas. Si salen menos, se fundieron.
  if ($d.Tables.Count -eq 4) { "  OK    cuatro tablas separadas, una por bloque" }
  else { "  FALLA hay $($d.Tables.Count) tablas, se esperaban 4 (trampa 12)"; $fallos++ }

  # Las canceladas van en una tabla de una sola columna.
  $ultima = $d.Tables.Item($d.Tables.Count)
  if ($ultima.Columns.Count -eq 1) { "  OK    la tabla de canceladas trae una sola columna" }
  else { "  FALLA la tabla de canceladas trae $($ultima.Columns.Count) columnas"; $fallos++ }

  # Encabezado + una fila por linea. El CSV de arriba trae 2, 1, 1 y 1 lineas, asi que las
  # tablas van de 3, 2, 2 y 2 filas.
  #
  # Esto existe por un fallo MEDIDO el 09/09/2026: con UNA sola linea, ConvertFrom-Csv devuelve
  # un objeto suelto en vez de un array y su .Count llega vacio, asi que la tabla se creaba con
  # una sola fila y el bloque salia en el documento SIN encabezados. Con dos o mas lineas nunca
  # fallo, que es por lo que aguanto tanto sin verse. Tres de los cuatro bloques de esta prueba
  # traen una sola linea a proposito.
  $espera = @(3, 2, 2, 2)
  $malas  = @()
  for ($t = 1; $t -le $d.Tables.Count; $t++) {
    $tb = $d.Tables.Item($t)
    if ($t -le $espera.Count -and $tb.Rows.Count -ne $espera[$t-1]) {
      $malas += "la tabla $t trae $($tb.Rows.Count) filas y se esperaban $($espera[$t-1])"
    }
    $primera = ($tb.Cell(1,1).Range.Text -replace "[`r`a]", '')
    if ($primera -ne 'Numero') { $malas += "la tabla $t no empieza por la fila de encabezados, empieza por '$primera'" }
  }
  if ($malas.Count -eq 0) { "  OK    cada tabla lleva su fila de encabezados y sus filas de datos" }
  else { "  FALLA " + ($malas -join '; '); $fallos++ }

  $d.Save()
  $d.Close()
} finally {
  if ($null -ne $w) {
    # Best-effort: con Word sin licencia, Quit() lanza. No debe tapar el resultado.
    try { $w.Quit() } catch { "  AVISO: no se pudo cerrar Word limpiamente: $($_.Exception.Message)" }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    [GC]::Collect()
  }
  Remove-Item $csv -Force -ErrorAction SilentlyContinue
  if (-not $Conservar) { Remove-Item $docx -Force -ErrorAction SilentlyContinue }
}

""
if ($Conservar) { "Documento conservado: $docx" }
if ($fallos -eq 0) { "TODO OK" } else { "$fallos comprobacion(es) fallaron"; exit 1 }
