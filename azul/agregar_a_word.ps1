# Agrega el resultado de una corrida al FINAL del documento de Word que este abierto.
# Empieza con un SALTO DE PAGINA (lo mismo que Ctrl+Enter) para no encimarse con lo
# que ya haya, y escribe una tabla por bloque, cada una con su titulo.
# No guarda: revisa y guarda tu con Ctrl+S. Un Ctrl+Z deshace lo agregado.
#
#   .\agregar_a_word.ps1                     -> usa el CSV mas reciente de .\salidas
#   .\agregar_a_word.ps1 -Csv "ruta\x.csv"   -> usa ese CSV
#   .\agregar_a_word.ps1 -Documento "ruta\BASE 034 PS1 VIRLAN.docx"  -> escribe en ESE, no en
#                                               el activo. Si no esta abierto, lo abre.
#
# NOTA: mantener este archivo en ASCII puro, igual que azul_fast.ps1. PowerShell 5.1
# lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

param(
  [string]$Csv = "",
  # Ruta del PNG con la informacion general del cliente. Normalmente no se pasa: viene en
  # el propio CSV, en la linea '# IMAGEN:'. Esta aqui para poder forzar una a mano.
  [string]$Imagen = "",
  # Encabezado del bloque: numero de cuenta y razon social, como los pide Dorian y como
  # vienen en el Excel de ordenes. Si no se pasan -- ejecucion a mano, sin lote -- se cae al
  # nombre que azul_fast.ps1 leyo de la pantalla de Azul, que es lo que se hacia antes.
  [string]$Cuenta = "",
  [string]$Razon = "",
  # EN CUAL documento escribir. Vacio = el que este activo en Word, que es como se uso siempre
  # a mano. Desde el 09/09/2026 lote.ps1 ya no depende del activo: crea la base el mismo y dice
  # explicitamente donde escribir, porque el activo deja de serlo en cuanto alguien pincha otra
  # ventana de Word a media corrida.
  #
  #   -DocObj    el objeto del documento, tal cual. Es lo que pasa lote.ps1: corre en el mismo
  #              proceso, y asi no depende de que la instancia de Word este registrada para que
  #              GetActiveObject la encuentre.
  #   -Documento la ruta. Se busca entre los documentos ya abiertos y, si no esta, se abre.
  [object]$DocObj = $null,
  [string]$Documento = "",
  # Saltarse las guardas sobre QUE documento es el activo (sin guardar, autorrecuperado).
  [switch]$Forzar
)
$ErrorActionPreference = 'Stop'

if ($Csv -eq "") {
  # Primero el CSV que acaba de escribir azul_fast.ps1. Antes se tomaba directamente el mas
  # nuevo de salidas\, que son corridas archivadas: el 02/09/2026 eso eligio un CSV del dia
  # anterior. Ese fallo aviso porque el archivo era de un formato viejo, pero con uno
  # valido habria metido EL CLIENTE EQUIVOCADO en el documento sin decir nada.
  $reciente = Join-Path $PSScriptRoot "azul_renovables.csv"
  if (Test-Path $reciente) {
    $Csv = $reciente
  } else {
    $ultimo = Get-ChildItem (Join-Path $PSScriptRoot "salidas") -Filter *.csv -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $ultimo) { throw "No hay ningun CSV: ni azul_renovables.csv ni nada en salidas." }
    $Csv = $ultimo.FullName
  }
}
if (-not (Test-Path $Csv)) { throw "No encuentro el CSV: $Csv" }
Write-Host "CSV: $Csv"

# Segunda red por lo mismo: un CSV viejo casi siempre es de otro cliente.
$fecha = (Get-Item $Csv).LastWriteTime
if (((Get-Date) - $fecha).TotalHours -gt 12) {
  Write-Host "AVISO: este CSV es del $($fecha.ToString('dd/MM/yyyy HH:mm')), no de una corrida reciente."
  Write-Host "       Comprueba que sea el cliente que quieres ANTES de guardar el documento."
}

# ---- parseo por bloques -----------------------------------------------------
# "# NOMBRE (n)" abre un bloque; las lineas que empiezan con comilla son sus filas.
# El resto de las lineas '#' (cuadre, resumen) se ignora aqui y se rescata abajo.
$todo    = Get-Content $Csv -Encoding UTF8
$cols    = 'Numero','Plan','Dispositivo','MPE','PlanForzoso','FechaExpiracion'
$bloques = New-Object System.Collections.ArrayList
$actual  = $null

foreach ($ln in $todo) {
  if ($ln -match '^#\s+([A-Z][A-Z ]*[A-Z])\s*\(([0-9]+)\)') {
    $actual = New-Object psobject -Property @{
      Titulo = $Matches[1].Trim()
      Conteo = [int]$Matches[2]
      Filas  = New-Object System.Collections.ArrayList
    }
    [void]$bloques.Add($actual)
  }
  elseif ($ln.StartsWith('"') -and $actual -ne $null) {
    [void]$actual.Filas.Add($ln)
  }
}
if ($bloques.Count -eq 0) { throw "El CSV no trae bloques '# NOMBRE (n)'. Corre azul_fast.ps1 de nuevo." }

$cuadre = $todo | Where-Object { $_ -match '^#\s*CUADRE:' } | Select-Object -First 1
if ($cuadre) { $cuadre = ($cuadre -replace '^#\s*','') }

# Nombre del cliente e imagen de su informacion general, que azul_fast.ps1 antepone al CSV.
# Ojo: hay que volver a evaluar el -match sobre la linea ya elegida. El $Matches que deja un
# Where-Object no es de fiar, porque el bloque corre en su propio ambito.
$clienteNom = ""
$lnCli = $todo | Where-Object { $_ -match '^#\s*CLIENTE:' } | Select-Object -First 1
if ($lnCli -and ($lnCli -match '^#\s*CLIENTE:\s*(.+)$')) { $clienteNom = $Matches[1].Trim() }

if ($Imagen -eq "") {
  $lnImg = $todo | Where-Object { $_ -match '^#\s*IMAGEN:' } | Select-Object -First 1
  if ($lnImg -and ($lnImg -match '^#\s*IMAGEN:\s*(.+)$')) { $Imagen = $Matches[1].Trim() }
}

foreach ($b in $bloques) { Write-Host ("  {0}: {1} filas" -f $b.Titulo, $b.Filas.Count) }

# ---- Word -------------------------------------------------------------------
# Quien elige el documento: o lo dijo quien llama, o es el activo. La diferencia importa,
# porque las dos guardas de mas abajo existen solo para el caso "el activo", que es el unico
# en el que el destino se adivina.
$loEligioQuienLlama = ($null -ne $DocObj -or $Documento -ne "")
$d = $null

if ($null -ne $DocObj) {
  $d = $DocObj
}
else {
  try {
    $w = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
  } catch {
    throw "No hay un Word abierto al que conectarme. Abre el documento y vuelve a correr."
  }
  if ($w.Documents.Count -eq 0) { throw "Word esta abierto pero sin documentos." }

  if ($Documento -ne "") {
    $hoja = Split-Path $Documento -Leaf
    foreach ($doc in $w.Documents) {
      if ($doc.FullName -eq $Documento -or $doc.Name -eq $hoja) { $d = $doc; break }
    }
    if ($null -eq $d) {
      if (-not (Test-Path -LiteralPath $Documento)) {
        throw "No encuentro el documento '$Documento', y no esta abierto en Word."
      }
      $d = $w.Documents.Open($Documento)
    }
  }
  else { $d = $w.ActiveDocument }
}

Write-Host "Documento: $($d.Name)"
Write-Host "Tenia cambios sin guardar antes de escribir: $(-not $d.Saved)"

# El destino es el documento ACTIVO, y eso no siempre es el que uno cree. El 02/09/2026 Word
# levanto una copia de autorrecuperacion y la corrida escribio ahi: al guardarla habria
# creado un archivo nuevo en vez de actualizar el del cliente. Se comprueba antes de escribir.
# (Sin acentos en el patron a proposito: este archivo es ASCII puro.)
if (-not $Forzar -and -not $loEligioQuienLlama) {
  if ($d.Path -eq "") {
    throw "El documento activo ('$($d.Name)') nunca se ha guardado en disco. Activa el documento del cliente y vuelve a correr, o usa -Forzar."
  }
  if ($d.Name -match 'Recuperado autom|AutoRecuperado|AutoRecovered|Recovered') {
    throw "El documento activo es una copia de AUTORRECUPERACION ('$($d.Name)'). Si escribes ahi y guardas, creas un archivo nuevo en vez de actualizar el del cliente. Abre el documento bueno, activalo, y vuelve a correr; o usa -Forzar si de verdad quieres escribir en la copia."
  }
}

# Escribe un parrafo al final del documento.
#
# NO se hace "$p.Range.Text = ...", que es lo obvio y esta mal: el rango de un parrafo
# INCLUYE su marca de fin, y asignarle texto la borra, con lo que el parrafo se funde con lo
# que venga despues. Costo real de esa linea (02/09/2026): los titulos de bloque
# desaparecian, el titulo de la imagen salia pegado DETRAS de la foto en el mismo parrafo, y
# las dos tablas quedaban fundidas en una sola de 20 filas. Con MoveEnd se deja la marca
# fuera del rango y cada parrafo queda entero.
function Add-Parrafo {
  param($Doc, [string]$Texto, [switch]$Negrita)
  $p = $Doc.Paragraphs.Add()
  $r = $p.Range
  [void]$r.MoveEnd(1, -1)      # 1 = wdCharacter, -1 = deja fuera la marca de parrafo
  $r.Text = $Texto
  $r.Bold = [bool]$Negrita
  return $p
}

# Salto de pagina al final, igual que Ctrl+Enter: separa este cliente del anterior.
#
# EN UN DOCUMENTO RECIEN CREADO NO SE PONE. Desde el 09/09/2026 el lote crea la base vacia, y
# un salto sobre un documento en blanco deja una primera pagina vacia delante del primer
# cliente. Antes no pasaba porque el documento lo abria siempre Dorian, con algo dentro.
$vacio = (($d.Content.Text.Trim() -eq "") -and ($d.Tables.Count -eq 0) -and ($d.InlineShapes.Count -eq 0))
if ($vacio) {
  Write-Host "El documento estaba vacio: no se mete salto de pagina delante del primer cliente."
}
else {
  $rng = $d.Content
  $rng.Collapse(0)          # wdCollapseEnd
  $rng.InsertBreak(7)       # wdPageBreak
}

$hoy = Get-Date -Format 'dd/MM/yyyy'

# Cabecera pedida por Dorian el 04/09/2026, confirmada contra una captura suya: numero de
# cuenta en una linea, razon social en la siguiente, y nada mas antes de la imagen. Los dos
# datos salen del Excel de ordenes, no de la pantalla de Azul.
#
# Se quitaron a proposito, porque Dorian los dio por sobrantes:
#   - la fecha y el "- lectura de suscripciones"
#   - la linea del CUADRE. SIGUE calculandose y sigue en el CSV, y revisar_csv.ps1 sigue
#     bloqueando el paso a Word cuando la corrida esta incompleta: lo unico que cambia es
#     que ya no se imprime en el documento.
#   - el titulo "Informacion general del cliente" que iba encima de la imagen.
if ($Cuenta -ne "" -or $Razon -ne "") {
  if ($Cuenta -ne "") { [void](Add-Parrafo $d $Cuenta -Negrita) }
  if ($Razon  -ne "") { [void](Add-Parrafo $d $Razon  -Negrita) }
}
elseif ($clienteNom -ne "") { [void](Add-Parrafo $d "$clienteNom - lectura de suscripciones $hoy" -Negrita) }
else                        { [void](Add-Parrafo $d "Lectura de suscripciones - $hoy" -Negrita) }

# La informacion general del cliente va PRIMERO, antes de las tablas de lineas: es el
# encabezado del bloque de este cliente. La imagen se INCRUSTA (LinkToFile=$false,
# SaveWithDocument=$true), asi que el documento no depende de que el PNG siga en disco.
# Que falte la imagen nunca aborta el paso a Word: las tablas valen por si solas.
if ($Imagen -ne "" -and (Test-Path -LiteralPath $Imagen)) {
  $pImg = $d.Paragraphs.Add()
  $sh = $d.InlineShapes.AddPicture($Imagen, $false, $true, $pImg.Range)
  # Con la proporcion bloqueada basta con tocar el ancho: el alto la sigue y no se deforma.
  $sh.LockAspectRatio = -1
  $util = $d.PageSetup.PageWidth - $d.PageSetup.LeftMargin - $d.PageSetup.RightMargin
  if ($sh.Width -gt $util) { $sh.Width = $util }
  Write-Host "Imagen: $Imagen"
}
elseif ($Imagen -ne "") {
  [void](Add-Parrafo $d "(No se encontro la imagen de la informacion general.)")
  Write-Host "AVISO: el CSV apunta a una imagen que no existe: $Imagen"
}
else {
  # HUECO PARA LA IMAGEN. Desde el 05/09/2026 la foto de la vista general del cliente ya no
  # la saca el sistema: la pega Dorian a mano. Lo que hace falta aqui es dejar el sitio, en el
  # mismo punto del bloque donde iba la foto -- entre la razon social y la primera tabla --
  # para no tener que abrir hueco a mano dentro de un parrafo en negrita.
  #
  # Marcador visible y no un parrafo vacio, a peticion de Dorian: al pegar, Word sustituye la
  # seleccion, asi que '[imagen]' se selecciona de un doble clic y desaparece solo. Un parrafo
  # vacio obligaria a colocar el cursor a ciegas entre dos parrafos.
  #
  # El camino de la imagen NO se borro. Si vuelve a llegar una ruta -- por el parametro
  # -Imagen o por una linea '# IMAGEN:' del CSV -- la primera rama la incrusta igual que
  # siempre. Esto es solo el caso "no hay ruta", que hasta hoy no escribia nada.
  [void](Add-Parrafo $d "[imagen]")
  Write-Host "Hueco '[imagen]' dejado en el bloque: pega ahi la captura del cliente."
}

$titulos = @{
  'RENOVABLES'            = 'RENOVABLES'
  'NO RENOVABLES'         = 'NO RENOVABLES'
  'A REVISAR'             = 'A REVISAR (el motivo va junto al numero)'
  'CANCELADAS EXCLUIDAS'  = 'CANCELADAS (excluidas, no se consideran)'
}

foreach ($b in $bloques) {
  $texto = $titulos[$b.Titulo]
  if (-not $texto) { $texto = $b.Titulo }

  [void](Add-Parrafo $d "$texto - $($b.Filas.Count)" -Negrita)

  if ($b.Filas.Count -eq 0) {
    [void](Add-Parrafo $d "Ninguna.")
    continue
  }

  # El @() no es adorno. Con UNA sola linea, ConvertFrom-Csv devuelve un objeto suelto y no un
  # array, y su .Count llega VACIO: la tabla se creaba con $null + 1 = 1 fila, es decir solo la
  # de encabezados, y la fila de datos acababa encima de ella. MEDIDO el 09/09/2026: un bloque
  # de una sola linea salia en el documento SIN encabezados. Con dos o mas nunca fallo, que es
  # por lo que aguanto tanto sin verse.
  $filas = @($b.Filas | ConvertFrom-Csv -Header $cols)

  # Las canceladas solo traen numero: una columna basta y no deja huecos vacios.
  if ($b.Titulo -eq 'CANCELADAS EXCLUIDAS') {
    $campos = @('Numero')
    $rotulo = @('Numero')
  } else {
    $campos = $cols
    $rotulo = @('Numero','Plan','Dispositivo','MPE','Plan forzoso','Fecha expiracion')
  }
  $nc = $campos.Count

  $tbl = $d.Tables.Add($d.Paragraphs.Add().Range, $filas.Count + 1, $nc)
  $tbl.Borders.InsideLineStyle  = 1
  $tbl.Borders.OutsideLineStyle = 1

  for ($c = 1; $c -le $nc; $c++) {
    $tbl.Cell(1,$c).Range.Text = $rotulo[$c-1]
    $tbl.Cell(1,$c).Range.Bold = $true
  }

  $r = 2
  foreach ($f in $filas) {
    for ($c = 1; $c -le $nc; $c++) {
      $v = $f.($campos[$c-1])
      if ($null -eq $v) { $v = "" }
      $tbl.Cell($r,$c).Range.Text = $v
      $tbl.Cell($r,$c).Range.Bold = $false
    }
    $r++
  }
  $tbl.Columns.AutoFit()
}

Write-Host ""
if ($vacio) { Write-Host "Bloques escritos: $($bloques.Count), al principio del documento." }
else        { Write-Host "Bloques escritos: $($bloques.Count), despues de un salto de pagina." }
Write-Host "NO se guardo el documento. Revisa y guarda tu."
