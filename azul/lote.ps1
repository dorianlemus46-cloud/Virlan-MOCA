# Recorre la lista de ordenes entera y arma las bases de datos en Word.
#
# No reimplementa nada: encadena los dos scripts que ya funcionan y se ocupa solo de Word y
# del recorrido.
#
#   por cada cuenta distinta de la lista de ordenes (CSV):
#     buscar_cuenta.ps1   -> busca en Azul, entra al cliente, abre Suscripciones
#     azul_fast.ps1       -> lee las lineas, clasifica y deja el CSV (la foto la pega Dorian)
#     revisar_csv.ps1     -> corrida completa? cuantas renovables?
#     agregar_a_word.ps1  -> escribe el bloque, SOLO si el cliente tiene algo que aportar
#
# EL SISTEMA VUELVE A CREAR Y A NOMBRAR LAS BASES (09/09/2026, por Dorian). No hace falta
# tener Word abierto ni ningun documento activo: el lote crea "BASE 034 PS1 VIRLAN.docx" en
# azul\bases, escribe dentro y lo guarda cliente a cliente.
#
# Del 06/09 al 09/09/2026 NO nombraba, y por un motivo: nombraba deduciendo el numero de que
# archivos habia en la carpeta, asi que en cuanto Dorian renombraba una base terminada ese
# numero quedaba libre y la corrida siguiente lo reutilizaba. Eso NO ha vuelto. Lo que ha
# vuelto es el nombrado, apoyado ahora en una memoria de verdad -- bases.ps1 -- que solo
# avanza, no se deduce de la carpeta, y jamas sobrescribe un documento que ya exista.
#
# Una base son 10 clientes y una tirada son tres a cinco, asi que UNA BASE CRUZA VARIAS
# CORRIDAS: el lote continua la que quedo a medias y la cierra sola al llegar a 10.
#
# Los dos scripts hijos corren en PowerShell de 32 BITS, que es lo que exige el puente. Este
# de aqui corre en el de siempre, porque solo habla con Word por COM.
#
#   .\lote.ps1 -Simular                   ver el plan sin tocar Azul ni Word
#   .\lote.ps1 -Limite 2                  hacer solo los dos primeros clientes
#   .\lote.ps1 -VerBases                  que base esta abierta, quien va dentro, cuales cerradas
#   .\lote.ps1 -CerrarBase                cerrarla ya, aunque no haya llegado a 10
#   .\lote.ps1 -Documento "...\BASE.docx" escribir en ESE documento y no en la base automatica
#   .\lote.ps1 -PausaLineas 6 -PausaLectura 10 -PausaEscribir 8 -PausaResultado 15
#                                          cambiar el ritmo contra Azul sin tocar ningun archivo
#   .\lote.ps1
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

param(
  # La lista de ordenes, exportada de Excel a CSV. Vacio = lista_ordenes.csv junto al script.
  [string]$Lista = "",
  # ESCAPE A MANO. Vacio = base automatica, que es lo normal. Si se da una ruta, se escribe en
  # ESE documento y la memoria de bases NO se toca: ni cuenta clientes, ni corta a los 10, ni
  # gasta numero. Es para retomar una base vieja o para una prueba suelta.
  [string]$Documento = "",
  # La carpeta donde viven las bases. Vacio = azul\bases. Las pruebas la mandan a otro sitio.
  [string]$Carpeta = "",
  # Cuantos clientes hacen una base. 0 = los 10 de siempre, que define bases.ps1.
  [int]$PorBase = 0,
  # Ensenar el estado de las bases y salir. No toca Azul ni Word.
  [switch]$VerBases,
  # Cerrar ya la base abierta, aunque no haya llegado a 10, y salir. No toca Azul.
  [switch]$CerrarBase,
  # ADOPTAR una base empezada a mano: la copia a la carpeta con el nombre que toca y anota en
  # la memoria los clientes que ya trae dentro. Sin -Confirmar solo mira y cuenta. No toca Azul.
  [string]$Adoptar = "",
  [switch]$Confirmar,
  # 0 = toda la lista.
  [int]$Limite = 0,
  # Parar en cuanto se cierre una base, en vez de seguir y abrir la siguiente. Para pedir "una
  # base entera y nada mas" sin tener que adivinar cuantos clientes de la lista hacen falta:
  # se le da una rebanada larga y se detiene sola al llegar a los 10.
  [switch]$UnaBase,
  # Recorrer y explicar el plan sin tocar nada.
  [switch]$Simular,
  # Tope de tiempo, en segundos, para CADA proceso hijo. Al vencer se mata y se sigue.
  [int]$TopeSeg = 300,
  # Pausas ADICIONALES para no pedirle datos a Azul mas rapido de lo que aguanta una persona
  # (Dorian, 07/09/2026: a mano aguanta horas, con el script se caia en minutos). Se SUMAN a
  # las esperas que ya habia en cada script hijo; ninguna las sustituye. Se pasan tal cual a
  # buscar_cuenta.ps1 y azul_fast.ps1: cambiarlas aqui no toca ningun archivo. En segundos.
  [int]$PausaLineas = 4,      # azul_fast.ps1: entre cerrar el detalle de una linea y la siguiente
  [int]$PausaLectura = 7,     # azul_fast.ps1: entre abrir el detalle de una linea y leerlo
  [int]$PausaEscribir = 5,    # buscar_cuenta.ps1: entre escribir la cuenta y pulsar "Buscar Ahora"
  [int]$PausaResultado = 10   # buscar_cuenta.ps1: entre pulsar "Buscar Ahora" y leer el resultado
)
$ErrorActionPreference = 'Stop'

$raiz      = $PSScriptRoot
$ps32      = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
$csvLector = Join-Path $raiz 'azul_renovables.csv'
$progreso  = Join-Path $raiz 'salidas\lote_progreso.csv'
if ($Lista -eq "") { $Lista = Join-Path $raiz 'lista_ordenes.csv' }

# El lector de la lista de ordenes. Vive aparte porque buscar_cuenta.ps1 tambien lo usa
# para su modificador -Fila, y antes eran dos copias de la misma logica.
. (Join-Path $raiz 'lista.ps1')
# La memoria de las bases: que numero toca, cual esta abierta y quien esta dentro. Vive aparte
# porque no habla con Word ni con Azul, y asi se puede probar sola.
. (Join-Path $raiz 'bases.ps1')

$Carpeta = Get-CarpetaBases -Carpeta $Carpeta
if ($PorBase -le 0) { $PorBase = $BASES_POR_BASE }

# Dos modos, y conviene tenerlos separados en la cabeza: el normal, en el que la base la crea y
# la numera el sistema; y el de -Documento, que escribe donde se le diga y NO toca la memoria
# -- ni cuenta clientes, ni corta a los 10, ni gasta numero.
$modoManual = ($Documento -ne "")
if ($modoManual) {
  if (-not (Test-Path -LiteralPath $Documento)) {
    throw ("No encuentro el documento '$Documento'.`n" +
           "       Con -Documento no se crea nada: o existe, o quitalo y deja que el sistema`n" +
           "       abra la base que toca.")
  }
  $Documento = (Resolve-Path -LiteralPath $Documento).Path
}

function Ahora { (Get-Date).ToString('HH:mm:ss') }
function Log([string]$s) { Write-Host ("[{0}] {1}" -f (Ahora), $s) }

# ---- atajos que no tocan ni Azul ni Word ------------------------------------
# Van antes que todo lo demas a proposito: preguntar por el estado de las bases o cerrar una a
# medias no tiene por que exigir una lista de ordenes valida ni Azul abierto.
if ($VerBases) {
  Show-EstadoBases -Carpeta $Carpeta
  exit 0
}
if ($CerrarBase) {
  $reg = Get-RegistroBases -Carpeta $Carpeta
  if ($null -eq $reg.abierta) { Log "No hay ninguna base abierta que cerrar."; exit 0 }
  $c = Close-BaseAbierta -Carpeta $Carpeta -Registro $reg -Motivo 'cerrada a mano'
  Log "Base cerrada a mano: $($c.archivo) con $($c.clientes) cliente(s)."
  Log "La siguiente corrida abrira BASE $('{0:D3}' -f ([int]$reg.ultimoNumero + 1))."
  exit 0
}

# ---- 0. tope de tiempo sobre los procesos hijos ------------------------------
# Antes se lanzaban con "& $ps32 ... | Out-Host", que espera indefinidamente. Un hijo colgado
# esperando a Azul dejaba al padre parado para siempre: paso el 04/09/2026, 20 minutos sin
# tope, con "Estableciendo conexion a omserver" en pantalla. Con Dorian delante se corta con
# Ctrl+C; una corrida desatendida de 30 no puede depender de eso.
#
# Se lanza con Start-Process -PassThru y se le pone reloj. Al vencer: matar, devolver
# $TOPE_VENCIDO, y que el lote anote FALLO TIEMPO y siga con el cliente siguiente.
#
# -NoNewWindow para que el hijo siga escribiendo en esta misma consola, como hacia Out-Host.
$TOPE_VENCIDO = -999

# Start-Process une los elementos de -ArgumentList con espacios y NO los entrecomilla: una
# razon social como "CONSULTORES EN SERVICIOS" llegaria partida en tres argumentos. Se cita aqui.
function Cita([string]$s) { '"' + ($s -replace '"', '\"') + '"' }

function Invoke-Hijo([string]$script, [string]$argumentos, [int]$Tope = 0) {
  if ($Tope -le 0) { $Tope = $TopeSeg }
  $ruta = Join-Path $raiz $script
  $cmd  = '-NoProfile -ExecutionPolicy Bypass -File ' + (Cita $ruta)
  if ($argumentos) { $cmd += ' ' + $argumentos }

  $p = Start-Process -FilePath $ps32 -ArgumentList $cmd -NoNewWindow -PassThru

  # Sin esta linea, ExitCode devuelve $null SIEMPRE, aunque el hijo termine bien. Start-Process
  # -PassThru entrega el objeto sin conservar el handle del proceso, y .NET ya no puede leer el
  # codigo cuando el hijo sale. Tocar .Handle obliga a cachearlo mientras el proceso vive.
  # Medido el 04/09/2026: sin la linea, "exit 0", "exit 1" y "exit 7" devolvian los tres
  # $null; con ella devuelven 0, 1 y 7. Costo una corrida entera: los dos clientes se buscaron
  # bien y el lote los dio por fallidos, porque $null -ne 0.
  $null = $p.Handle

  if ($p.WaitForExit($Tope * 1000)) {
    $p.WaitForExit()          # cierra el volcado de la salida antes de leer ExitCode
    $codigo = $p.ExitCode
    if ($null -eq $codigo) {
      # No deberia pasar ya. Si pasa, se dice en voz alta en vez de confundirlo con un fallo
      # del hijo, y NO se devuelve 0: dar por buena una busqueda que no se pudo comprobar es
      # justo el error que lleva a leer un cliente bajo el nombre de otro.
      Log "  AVISO: $script termino pero Windows no dio codigo de salida. Se cuenta como fallo."
      return -998
    }
    return $codigo
  }

  Log "  TOPE DE TIEMPO: $script paso de $Tope s sin terminar. Se mata el PID $($p.Id)."
  try { $p.Kill() } catch { Log "  no se pudo matar el PID $($p.Id): $($_.Exception.Message)" }
  $null = $p.WaitForExit(5000)
  return $TOPE_VENCIDO
}

# ---- 1. la lista de ordenes -------------------------------------------------
# Vive en lista.ps1, dot-sourceado arriba. Antes se leia del libro de Excel ABIERTO,
# localizado por la hoja 'resultData'; desde el 05/09/2026 es un CSV que Dorian exporta con
# Guardar como. Lo que NO cambio: se busca por encabezado y no por posicion, y una cuenta
# repetida se salta -- es la misma base de datos, y contarla dos veces romperia el corte.
#
# ---- 2. progreso ------------------------------------------------------------
function Add-Progreso($Cuenta, $Razon, $Estado, $Documento, $Detalle) {
  $dir = Split-Path $progreso -Parent
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  if (-not (Test-Path $progreso)) {
    'cuenta,razon,estado,documento,hora,detalle' | Out-File -FilePath $progreso -Encoding utf8
  }
  $f = { param($s) '"' + ([string]$s).Replace('"','""') + '"' }
  $linea = (& $f $Cuenta) + ',' + (& $f $Razon) + ',' + (& $f $Estado) + ',' +
           (& $f $Documento) + ',' + (& $f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) + ',' + (& $f $Detalle)
  $linea | Out-File -FilePath $progreso -Encoding utf8 -Append
}

# ---- 3. Word ----------------------------------------------------------------
# Word se abre SOLO si hace falta, y no antes del primer cliente que de verdad vaya a entrar al
# documento: si la corrida entera falla en Azul, no se abrio Word para nada.
#
# Si ya hay un Word abierto se usa ese, y al terminar NO se cierra: es el de Dorian. Si no lo
# hay, se abre uno y ese si se cierra al final. La distincion la lleva $wordNuestro.
#
# Visible = $true a proposito. Un Word invisible que quede colgado tras una caida es un proceso
# reteniendo el documento sin nada en pantalla que lo diga; visible, se ve y se cierra.
$script:word       = $null
$script:wordNuestro = $false

function Get-WordApp {
  if ($null -ne $script:word) { return $script:word }
  try {
    $script:word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    $script:wordNuestro = $false
    Log "Word: ya estaba abierto, se usa ese y no se cierra al terminar."
  }
  catch {
    $script:word = New-Object -ComObject Word.Application
    $script:word.Visible = $true
    $script:wordNuestro = $true
    Log "Word: no habia ninguno abierto. Se abre uno y se cierra al terminar."
  }
  return $script:word
}

# Se guarda AL CREAR y no al llegar a 10: agregar_a_word.ps1 se niega a escribir en un
# documento sin ruta en disco, y asi una caida a mitad no se lleva lo ya escrito.
function New-DocumentoBase {
  param([Parameter(Mandatory=$true)][string]$Ruta)
  $w = Get-WordApp
  $doc = $w.Documents.Add()
  $doc.SaveAs([ref]$Ruta, [ref]16)   # 16 = wdFormatDocumentDefault (.docx)
  return $doc
}

function Open-DocumentoBase {
  param([Parameter(Mandatory=$true)][string]$Ruta)
  $w = Get-WordApp
  # Puede que ya este abierto -- por ejemplo si Dorian lo dejo mirando de la corrida anterior.
  # Abrirlo dos veces daria un documento de solo lectura y las tablas se irian al vacio.
  $hoja = Split-Path $Ruta -Leaf
  foreach ($d in $w.Documents) {
    if ($d.FullName -eq $Ruta -or $d.Name -eq $hoja) { return $d }
  }
  return $w.Documents.Open($Ruta)
}

function Close-DocumentoBase {
  param($Doc)
  if ($null -eq $Doc) { return }
  $nombre = $Doc.Name
  try { $Doc.Save(); $Doc.Close() } catch { Log "AVISO: no se pudo cerrar '$nombre': $($_.Exception.Message)" }
}

# Word solo se cierra si lo abrimos nosotros, y solo si no queda ningun documento dentro: si
# Dorian abrio algo suyo mientras corria, cerrarselo seria peor que dejar Word abierto.
# Se llama desde el final de la corrida Y desde las salidas cortas, para no dejar un Word
# suelto por haber salido antes de tiempo.
function Close-WordSiNuestro {
  if (-not $script:wordNuestro -or $null -eq $script:word) { return }
  try {
    if ($script:word.Documents.Count -eq 0) { $script:word.Quit(); Log "Word cerrado (lo habia abierto la corrida)." }
    else { Log "Word queda abierto: tiene $($script:word.Documents.Count) documento(s) dentro." }
  }
  catch { Log "AVISO: no se pudo cerrar Word: $($_.Exception.Message)" }
  $script:word = $null
}

# ---- 3b. la base de esta corrida --------------------------------------------
# EL DOCUMENTO NO SE ABRE HASTA QUE HAY ALGO QUE ESCRIBIR EN EL. Antes se elegia el destino
# antes del primer cliente, porque lo elegia una persona y equivocarse costaba caro; ahora lo
# elige la memoria y no hay nada que adivinar, asi que se puede esperar.
$script:docBase = $null
$script:registro = $null

function Get-DocBase {
  if ($null -ne $script:docBase) { return $script:docBase }

  if ($modoManual) {
    $script:docBase = Open-DocumentoBase -Ruta $Documento
    Log "A mano: se escribe en $($script:docBase.Name). La memoria de bases no se toca."
    return $script:docBase
  }

  if ($null -eq $script:registro.abierta) {
    $nueva = New-BaseAbierta -Carpeta $Carpeta -Registro $script:registro -Log { param($m) Log $m }
    Log "Base nueva: $($nueva.Archivo)"
    try {
      $script:docBase = New-DocumentoBase -Ruta $nueva.Ruta
    }
    catch {
      # El numero ya estaba apartado (regla 1 de bases.ps1). Se QUEMA: se cierra esa base sin
      # documento y la siguiente corrida usara el numero de despues. Quemar un numero no cuesta
      # nada; reutilizarlo es lo que rompio el sistema la vez pasada.
      [void](Close-BaseAbierta -Carpeta $Carpeta -Registro $script:registro -Motivo 'no se pudo crear el documento')
      throw ("No se pudo crear '$($nueva.Archivo)': $($_.Exception.Message)`n" +
             "       Ese numero queda quemado a proposito. La siguiente corrida usa el siguiente.")
    }
  }
  else {
    $ruta = Join-Path $Carpeta ([string]$script:registro.abierta.archivo)
    $script:docBase = Open-DocumentoBase -Ruta $ruta
    Log "Base a medias: $($script:registro.abierta.archivo), $(@($script:registro.abierta.clientes).Count) de $PorBase cliente(s) dentro."
  }
  return $script:docBase
}

# Cierra la base al llegar a los 10, y deja el documento cerrado en el disco. La corrida puede
# seguir: el cliente siguiente abrira la base de despues.
function Complete-BaseSiToca {
  param([int]$Dentro)
  if ($Dentro -lt $PorBase) { return }
  Close-DocumentoBase $script:docBase
  $c = Close-BaseAbierta -Carpeta $Carpeta -Registro $script:registro -Motivo 'completa'
  $script:docBase = $null
  Log "BASE COMPLETA: $($c.archivo) con $($c.clientes) clientes. Guardada y cerrada."
}

# ---- 3c. adoptar una base empezada a mano -----------------------------------
# Sirve para el paso de una vez: un documento que Dorian venia llenando a mano pasa a ser la
# base que el sistema lleva. Se hace con un comando y no editando la memoria con el Bloc de
# notas, porque una memoria escrita a mano es justo lo que puede repartir un numero ya usado.
#
# Se COPIA, no se mueve: el original se queda donde esta. El precio es que quedan dos copias
# de la misma base, asi que hay que decirlo muy claro y que Dorian archive la vieja.
#
# Como se reconocen los clientes que ya trae dentro: cada bloque empieza por un parrafo, fuera
# de tabla, que es SOLO el numero de cuenta -- es como los escribe agregar_a_word.ps1 desde el
# 04/09/2026. La razon social es el parrafo siguiente. Es una heuristica sobre un documento que
# pudo tocar una persona, asi que SIN -Confirmar esto solo mira y ensena lo que encontro.
function Get-ClientesDelDocumento {
  param([Parameter(Mandatory=$true)][string]$Ruta)
  $w = Get-WordApp
  $hoja = Split-Path $Ruta -Leaf
  $doc = $null; $eraNuestro = $false
  foreach ($x in $w.Documents) { if ($x.FullName -eq $Ruta -or $x.Name -eq $hoja) { $doc = $x; break } }
  if ($null -eq $doc) { $doc = $w.Documents.Open($Ruta, $false, $true); $eraNuestro = $true }  # solo lectura

  $sueltos = @()
  foreach ($p in $doc.Paragraphs) {
    if ($p.Range.Information(12)) { continue }        # 12 = wdWithInTable
    $sueltos += ($p.Range.Text -replace "[`r`a]", '').Trim()
  }
  if ($eraNuestro) { $doc.Close(0) }                  # 0 = wdDoNotSaveChanges

  $encontrados = New-Object System.Collections.ArrayList
  for ($k = 0; $k -lt $sueltos.Count; $k++) {
    if ($sueltos[$k] -match '^\d{9}$') {
      $raz = ""
      if ($k + 1 -lt $sueltos.Count) { $raz = $sueltos[$k+1] }
      [void]$encontrados.Add([pscustomobject]@{ Cuenta = $sueltos[$k]; Razon = $raz })
    }
  }
  return $encontrados
}

# La copia se hace DEL DISCO. Si el documento esta abierto en Word con cambios sin guardar, esa
# copia se quedaria sin lo ultimo escrito -- y como el original se conserva, el error no se
# veria hasta echar de menos un cliente dentro de la base. Se aborta y se pide guardar.
function Test-AdoptarSinGuardar {
  param([Parameter(Mandatory=$true)][string]$Ruta)
  $w = Get-WordApp
  $hoja = Split-Path $Ruta -Leaf
  foreach ($x in $w.Documents) {
    if (($x.FullName -eq $Ruta -or $x.Name -eq $hoja) -and (-not $x.Saved)) { return $true }
  }
  return $false
}

if ($Adoptar -ne "") {
  if (-not (Test-Path -LiteralPath $Adoptar)) { throw "No encuentro el documento a adoptar: $Adoptar" }
  $Adoptar = (Resolve-Path -LiteralPath $Adoptar).Path

  $reg = Get-RegistroBases -Carpeta $Carpeta
  if ($null -ne $reg.abierta) {
    throw ("Ya hay una base abierta: $($reg.abierta.archivo), con $(@($reg.abierta.clientes).Count) cliente(s).`n" +
           "       Cierrala antes de adoptar otra:  lote.ps1 -CerrarBase")
  }

  Log "Mirando '$(Split-Path $Adoptar -Leaf)'..."
  $sinGuardar = Test-AdoptarSinGuardar -Ruta $Adoptar
  if ($sinGuardar) {
    Log "AVISO: ese documento esta abierto en Word CON CAMBIOS SIN GUARDAR."
    Log "       Lo de abajo se lee de lo que tienes en pantalla, pero la copia saldria del disco."
    Log "       Guardalo (Ctrl+S) antes de confirmar."
  }
  $dentro = Get-ClientesDelDocumento -Ruta $Adoptar
  ""
  "Clientes encontrados dentro del documento:"
  $i = 0
  foreach ($c in $dentro) { $i++; "  {0,2}. {1}  {2}" -f $i, $c.Cuenta, $c.Razon }
  ""
  Log "Son $($dentro.Count) de $PorBase. Quedarian $($PorBase - $dentro.Count) para cerrar la base."
  if ($dentro.Count -eq 0) { throw "No se reconocio ningun cliente dentro. No se adopta a ciegas." }
  if ($dentro.Count -gt $PorBase) { Log "AVISO: son mas de $PorBase. Comprueba la lista de arriba antes de confirmar." }

  if (-not $Confirmar) {
    ""
    Log "Esto ha sido solo mirar. Si la lista de arriba es correcta, repite con -Confirmar."
    Close-WordSiNuestro
    exit 0
  }

  if ($sinGuardar) {
    throw ("'$(Split-Path $Adoptar -Leaf)' esta abierto en Word CON CAMBIOS SIN GUARDAR.`n" +
           "       La copia se hace del disco, asi que se quedaria sin lo ultimo que escribiste,`n" +
           "       y como el original se conserva no lo notarias hasta echar de menos un cliente.`n" +
           "       Guardalo con Ctrl+S y vuelve a correr.")
  }

  $nueva = New-BaseAbierta -Carpeta $Carpeta -Registro $reg -Log { param($m) Log $m }
  try { Copy-Item -LiteralPath $Adoptar -Destination $nueva.Ruta -ErrorAction Stop }
  catch {
    [void](Close-BaseAbierta -Carpeta $Carpeta -Registro $reg -Motivo 'no se pudo copiar el documento adoptado')
    throw "No se pudo copiar el documento a '$($nueva.Archivo)': $($_.Exception.Message)"
  }
  foreach ($c in $dentro) {
    [void](Add-ClienteABase -Carpeta $Carpeta -Registro $reg -Cuenta $c.Cuenta -Razon $c.Razon -Detalle 'venia en el documento adoptado')
  }
  ""
  Log "ADOPTADA como $($nueva.Archivo), con $($dentro.Count) cliente(s) dentro."
  Log "Copia en: $($nueva.Ruta)"
  Log "EL ORIGINAL NO SE TOCO y ahora es una copia vieja: $Adoptar"
  Log "Archivalo o renombralo, para no seguir escribiendo en el por error."
  Close-WordSiNuestro
  exit 0
}

# ---- 4. recorrido -----------------------------------------------------------
# El bloque de -Log hace que las filas repetidas y las descartadas salgan con la misma marca
# de hora que el resto de la corrida, como cuando el lector vivia aqui dentro.
$ordenes = Get-ListaOrdenes -Ruta $Lista -Log { param($m) Log $m }
if ($Limite -gt 0 -and $Limite -lt $ordenes.Count) { $ordenes = $ordenes[0..($Limite-1)] }

# La simulacion sigue siendo la unica prueba del lote que no toca Azul ni Word, y desde que la
# lista sale de un CSV es tambien la prueba del lector. Ahora ademas ensena en que base caeria
# cada cliente, que es la unica forma de ver el corte de los 10 sin gastar una corrida.
if ($Simular) {
  Log "SIMULACION: no se toca Azul ni Word."
  $i = 0
  if ($modoManual) {
    foreach ($o in $ordenes) { $i++; "  {0,2}. {1}  {2}" -f $i, $o.Cuenta, $o.Razon }
    ""
    Log "Serian $i cliente(s), todos a '$(Split-Path $Documento -Leaf)'. La memoria de bases no se toca."
  }
  else {
    $regS = Get-RegistroBases -Carpeta $Carpeta
    $numS = [int]$regS.ultimoNumero + 1
    $dentroS = 0
    if ($null -ne $regS.abierta) {
      $numS    = [int]$regS.abierta.numero
      $dentroS = @($regS.abierta.clientes).Count
      Log "Base a medias: $($regS.abierta.archivo), con $dentroS de $PorBase cliente(s)."
    } else {
      Log "No hay base a medias. Se abriria BASE $('{0:D3}' -f $numS)."
    }
    foreach ($o in $ordenes) {
      $i++
      if ($dentroS -ge $PorBase) { $numS++; $dentroS = 0 }
      $dentroS++
      "  {0,2}. {1}  {2}   ->  BASE {3:D3}  ({4}/{5})" -f $i, $o.Cuenta, $o.Razon, $numS, $dentroS, $PorBase
    }
    ""
    Log "Serian $i cliente(s), hasta BASE $('{0:D3}' -f $numS)."
  }
  Log "El conteo real puede ser menor: los clientes sin nada que renovar no entran al documento"
  Log "y no cuentan para los $PorBase."
  exit 0
}

# Regla 5 de bases.ps1: si la memoria dice que hay una base abierta, su documento tiene que
# seguir en el disco. Se comprueba ANTES del primer cliente, que es lo unico que cuesta poco.
if (-not $modoManual) {
  $script:registro = Get-RegistroBases -Carpeta $Carpeta
  $chk = Test-BaseAbierta -Carpeta $Carpeta -Registro $script:registro
  if ($chk.Hay -and -not $chk.Existe) { throw $chk.Problema }
}

$hechos  = 0; $saltados = 0; $fallos = 0
$inicio  = Get-Date
if ($modoManual) { Log "Destino a mano: $Documento" }
else {
  Log "Carpeta de bases: $Carpeta"
  Log "Una base son $PorBase clientes. El corte lo lleva el sistema y cruza corridas."
}
Log "Pausas: $PausaLineas s entre lineas, $PausaLectura s antes de leer una linea, $PausaEscribir s antes de buscar, $PausaResultado s de mas en el plazo para reconocer el resultado."

try {
  foreach ($o in $ordenes) {
    # --- sigue Azul ahi? ---
    # Si Azul se ha caido, los hijos fallan en dos segundos y la corrida se come el resto de la
    # lista marcando a todo el mundo como fallo. Paso de verdad el 17/09/2026: Azul murio a los
    # 32 min con 1.7 GB y el lote quemo 14 clientes seguidos en 22 segundos. Se para en seco y
    # no se anota nada de ellos, para que la lista se pueda relanzar tal cual.
    if ($null -eq (Get-Process -Name jp2launcher -ErrorAction SilentlyContinue)) {
      Log "AZUL NO ESTA. Se para aqui: los clientes que faltan no se tocan ni se anotan."
      Log "Vuelve a abrir Azul y relanza la misma lista; los ya escritos se saltan solos."
      break
    }

    Log "=== $($o.Cuenta)  $($o.Razon) ==="

    # --- ya esta escrito? ---
    # La memoria sabe quien esta dentro de la base abierta, asi que relanzar una lista que ya
    # se corrio a medias no duplica clientes. Es lo que hacia el -Reanudar que se quito el
    # 06/09/2026, y ahora sale gratis. Se comprueba ANTES de buscar en Azul: saltarse un
    # cliente ya escrito ahorra los dos o tres minutos que cuesta leerlo.
    if (-not $modoManual) {
      $ya = Find-ClienteEnBases -Registro $script:registro -Cuenta $o.Cuenta
      if ($null -ne $ya) {
        Log "  ya esta en $($ya.Archivo). No se vuelve a escribir."
        Add-Progreso $o.Cuenta $o.Razon 'YA ESTABA' $ya.Archivo "ya escrito en esta base"
        $saltados++
        continue
      }
    }

    # --- buscar ---
    $rc = Invoke-Hijo 'buscar_cuenta.ps1' ("-Cuenta {0} -Razon {1} -Restaurar -PausaEscribir {2} -PausaResultado {3}" -f (Cita $o.Cuenta), (Cita $o.Razon), $PausaEscribir, $PausaResultado)
    if ($rc -eq $TOPE_VENCIDO) {
      Log "  la busqueda se colgo y se mato. Se pasa al siguiente."
      Add-Progreso $o.Cuenta $o.Razon 'FALLO TIEMPO' '' "buscar_cuenta.ps1 paso de $TopeSeg s y se mato"
      $fallos++
      continue
    }
    if ($rc -ne 0) {
      Log "  no se pudo abrir el cliente. Se pasa al siguiente."
      Add-Progreso $o.Cuenta $o.Razon 'FALLO BUSQUEDA' '' "buscar_cuenta.ps1 salio con $rc"
      $fallos++
      continue
    }

    # --- leer ---
    # -SinWord: quien decide si este cliente entra al documento es el lote, no el lector.
    # -Cuenta/-Razon no eligen que se lee: van a la cabecera del CSV para que el archivo diga
    # de quien es. Antes ese nombre salia del panel del cliente, que se abria para la foto.
    $rc = Invoke-Hijo 'azul_fast.ps1' ("-SinWord -Cuenta {0} -Razon {1} -PausaLineas {2} -PausaLectura {3}" -f (Cita $o.Cuenta), (Cita $o.Razon), $PausaLineas, $PausaLectura)
    if ($rc -eq $TOPE_VENCIDO) {
      Log "  la lectura se colgo y se mato. Se pasa al siguiente."
      Add-Progreso $o.Cuenta $o.Razon 'FALLO TIEMPO' '' "azul_fast.ps1 paso de $TopeSeg s y se mato"
      $fallos++
      continue
    }
    if ($rc -ne 0 -or -not (Test-Path $csvLector)) {
      Log "  la lectura fallo. Se pasa al siguiente."
      Add-Progreso $o.Cuenta $o.Razon 'FALLO LECTURA' '' "azul_fast.ps1 salio con $rc"
      $fallos++
      continue
    }

    # --- corrida completa? ---
    $r = & (Join-Path $raiz 'revisar_csv.ps1') -Csv $csvLector
    if ($r.Problemas.Count -gt 0) {
      Log "  corrida incompleta, NO se escribe: $($r.Problemas -join '; ')"
      Add-Progreso $o.Cuenta $o.Razon 'FALLO LECTURA' '' ($r.Problemas -join '; ')
      $fallos++
      continue
    }

    # --- vale la pena meterlo en la base? ---
    # Un cliente sin nada que renovar no entra al documento y no cuenta para los 10, pero queda
    # anotado con el motivo: no se pierde, simplemente no ensucia la base. Esta regla nunca ha
    # cambiado, ni cuando el corte lo llevaba Dorian ni ahora que lo vuelve a llevar el codigo.
    if ($r.Renovables -le 0 -and $r.Revisar -le 0) {
      Log "  sin nada que renovar ($($r.NoRenovables) no renovables). No entra a la base."
      Add-Progreso $o.Cuenta $o.Razon 'SIN RENOVABLES' '' "0 renovables, $($r.NoRenovables) no renovables"
      $saltados++
      continue
    }

    # --- escribir ---
    # Aqui, y no antes, se crea o se abre el documento: si la corrida entera hubiera fallado en
    # Azul, no se habria abierto Word para nada ni gastado un numero de base.
    $doc = Get-DocBase
    $nombreDoc = $doc.Name
    # Se le pasa el documento POR OBJETO y no se depende del activo. Antes se hacia Activate()
    # en cada cliente porque el escritor iba a ActiveDocument y bastaba con que alguien pinchara
    # otra ventana de Word para que dejara de ser el nuestro. Ahora no hay a que apuntar mal.
    & (Join-Path $raiz 'agregar_a_word.ps1') -Csv $csvLector -Cuenta $o.Cuenta -Razon $o.Razon -DocObj $doc | Out-Host
    # Se guarda cliente a cliente. agregar_a_word.ps1 no guarda a proposito -- a mano, Dorian
    # revisa y guarda -- pero aqui van varios seguidos y una caida a mitad se llevaria lo ya
    # escrito. Es el mismo motivo por el que la base se guarda al crearla.
    $doc.Save()

    $hechos++
    # La memoria se anota DESPUES de guardar el documento. Si la maquina se cayera entre las dos
    # cosas, el cliente estaria en el documento y no en la memoria: se ve al abrir el documento
    # y se corrige. Al reves -- anotado y sin escribir -- lo perderia en silencio.
    $dentro = 0
    if (-not $modoManual) {
      $dentro = Add-ClienteABase -Carpeta $Carpeta -Registro $script:registro -Cuenta $o.Cuenta -Razon $o.Razon `
                                 -Detalle "$($r.Renovables) renovables, $($r.Revisar) a revisar"
      Log "  escrito en $nombreDoc  ($dentro de $PorBase en la base, $hechos en esta corrida)  -- $($r.Renovables) renovables, $($r.Revisar) a revisar"
    }
    else {
      Log "  escrito en $nombreDoc  ($hechos en esta corrida)  -- $($r.Renovables) renovables, $($r.Revisar) a revisar"
    }
    Add-Progreso $o.Cuenta $o.Razon 'ESCRITO' $nombreDoc "$($r.Renovables) renovables, $($r.NoRenovables) no renovables, $($r.Revisar) a revisar"

    if (-not $modoManual) {
      Complete-BaseSiToca -Dentro $dentro
      if ($UnaBase -and $dentro -ge $PorBase) {
        Log "Se pidio UNA base: la lista sigue teniendo clientes, pero la corrida para aqui."
        break
      }
    }
  }
} finally {
  # Se guarda pase lo que pase: es trabajo real.
  #
  # Con -Documento NO se cierra: ese documento lo eligio Dorian y puede estar mirandolo.
  # Con la base automatica si se cierra, porque el documento lo abrio el sistema y la base sigue
  # viva en la memoria -- la corrida siguiente la vuelve a abrir donde se quedo.
  if ($null -ne $script:docBase) {
    $n = $script:docBase.Name
    if ($modoManual) {
      try { $script:docBase.Save(); Log "Guardado (no se cierra, es tuyo): $n" }
      catch { Log "AVISO: no se pudo guardar '$n': $($_.Exception.Message)" }
    }
    else {
      Close-DocumentoBase $script:docBase
      Log "Guardada y cerrada: $n"
      $script:docBase = $null
    }
  }
  Close-WordSiNuestro
}

$mins = [math]::Round(((Get-Date) - $inicio).TotalMinutes, 1)
""
Log "FIN. $hechos escrito(s), $saltados saltado(s), $fallos fallo(s), en $mins min."
if ($modoManual) { Log "Documento: $(Split-Path $Documento -Leaf)" }
else {
  # Se relee la memoria del disco a proposito: es lo que de verdad quedo escrito, no lo que
  # esta corrida cree recordar.
  $fin = Get-RegistroBases -Carpeta $Carpeta
  if ($null -eq $fin.abierta) {
    Log "No queda ninguna base a medias. La siguiente sera BASE $('{0:D3}' -f ([int]$fin.ultimoNumero + 1))."
  } else {
    Log "Base a medias: $($fin.abierta.archivo), con $(@($fin.abierta.clientes).Count) de $PorBase cliente(s)."
    Log "La corrida siguiente sigue ahi. Para cerrarla antes de tiempo: lote.ps1 -CerrarBase"
  }
  Log "Carpeta: $Carpeta"
}
Log "Progreso: $progreso"
if ($fallos -gt 0) {
  # Volver a correr la misma lista ya no duplica: la memoria sabe quien esta dentro de la base
  # abierta y se los salta antes de buscarlos en Azul. Eso solo vale mientras la base siga
  # abierta; una vez cerrada, sus clientes ya no se reconocen.
  Log "Para reintentar los fallos puedes volver a correr la misma lista: los ya escritos se saltan."
  Log "Cuales fallaron esta en el progreso, columna estado."
}
