# La MEMORIA de las bases: que numero toca, cual esta abierta y quien esta dentro de ella.
#
# Lo carga con dot-source quien la necesite:  . (Join-Path $PSScriptRoot 'bases.ps1')
#
# ---------------------------------------------------------------------------------------
# POR QUE EXISTE ESTE ARCHIVO
#
# Hasta el 06/09/2026 el sistema nombraba las bases MIRANDO QUE ARCHIVOS HABIA en la carpeta
# -- esa era toda su memoria. Deducir no es recordar: en cuanto Dorian renombraba una base
# terminada, ese numero quedaba libre y la corrida siguiente lo reutilizaba, asi que el
# correlativo dejaba de ser unico. No tenia arreglo mientras el sistema dedujera, y por eso el
# 06/09/2026 se le quito la potestad de nombrar (docs/decisiones.md).
#
# Desde el 09/09/2026 vuelve a nombrar, pero ya no deduce: RECUERDA. El numero sale de este
# registro, que SOLO AVANZA, y la carpeta se consulta unicamente para negarse a sobrescribir.
#
# Las cinco reglas, que son el archivo entero:
#
#   1. EL NUMERO SE APARTA ANTES DE QUE EXISTA EL DOCUMENTO. Primero se escribe en el registro
#      que el 034 esta tomado, y despues se crea el .docx. Si la creacion falla, ese numero se
#      QUEMA y se pasa al siguiente. Quemar un numero no cuesta nada; reutilizar uno es lo que
#      rompio el sistema la vez pasada.
#   2. EL CONTADOR NUNCA RETROCEDE. Aunque se borren o se muevan archivos de la carpeta, el
#      numero solo sube. Y si la carpeta trajera un numero MAS ALTO que el registro, gana el
#      mas alto y se dice en voz alta.
#   3. JAMAS SE SOBRESCRIBE. Si el nombre que toca ya existe, se avanza y se avisa. Es la red
#      por si las dos de arriba fallaran.
#   4. EL REGISTRO SE ESCRIBE ENTERO O NO SE ESCRIBE. Va a un temporal, se vuelve a leer para
#      comprobar que quedo bien, y solo entonces se renombra encima del bueno. Una memoria a
#      medias seria peor que no tener memoria.
#   5. SI LA BASE ABIERTA DESAPARECIO DEL DISCO, SE PLANTA Y PREGUNTA. No empieza otra con el
#      mismo numero ni sigue como si nada. Es REGLAS.md seccion 6 aplicada aqui.
#
# Este archivo NO habla con Word ni con Azul a proposito: es solo memoria y nombres. Asi se
# puede probar entero sin Word abierto y sin Azul delante (diagnostico\probar_bases.ps1).
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

# El primer numero que se va a repartir. Lo eligio Dorian el 09/09/2026: la serie del
# escritorio tiene un hueco en el 034 y ahi quiere que empiece la carpeta nueva. Solo se usa
# cuando TODAVIA NO HAY REGISTRO; en cuanto lo hay, manda el registro.
$BASES_PRIMER_NUMERO = 34
# El nombre queda "BASE 034 PS1 VIRLAN.docx". El sufijo se guarda dentro del registro, asi que
# cambiarlo aqui solo afecta a una carpeta que todavia no tenga registro.
$BASES_SUFIJO        = 'PS1 VIRLAN'
# Una base son 10 clientes (ARRANQUE.md). Con bloques cortos, una base cruza varias corridas.
$BASES_POR_BASE      = 10
$BASES_ARCHIVO       = 'registro_bases.json'
$BASES_VERSION       = 1

# MEDIDO el 09/09/2026: dentro de una funcion de un archivo dot-sourceado, $PSScriptRoot es el
# del ARCHIVO DONDE SE DEFINIO la funcion -- este --, no el de quien lo carga. Asi que la
# carpeta por defecto sale siempre de azul\, tanto si llama lote.ps1 como si llama una prueba
# desde diagnostico\. Se comprobo porque leyendo el codigo parece lo contrario.
function Get-CarpetaBases {
  param([string]$Carpeta = "")
  if ($Carpeta -ne "") { return $Carpeta }
  return (Join-Path $PSScriptRoot 'bases')
}

function Get-RutaRegistro {
  param([Parameter(Mandatory=$true)][string]$Carpeta)
  return (Join-Path $Carpeta $BASES_ARCHIVO)
}

function Get-NombreBase {
  param([Parameter(Mandatory=$true)][int]$Numero, [string]$Sufijo = "")
  if ($Sufijo -eq "") { $Sufijo = $BASES_SUFIJO }
  return ("BASE {0:D3} {1}.docx" -f $Numero, $Sufijo)
}

# El numero mas alto que se ve en la carpeta. NO es la memoria -- es la red de la regla 2:
# si alguien dejo ahi un BASE 040, el contador salta por encima en vez de pisarlo.
function Get-MaxNumeroEnCarpeta {
  param([Parameter(Mandatory=$true)][string]$Carpeta)
  $max = 0
  if (-not (Test-Path -LiteralPath $Carpeta)) { return $max }
  foreach ($f in @(Get-ChildItem -LiteralPath $Carpeta -Filter '*.docx' -File -ErrorAction SilentlyContinue)) {
    if ($f.Name -match '^BASE\s+(\d+)') {
      $n = [int]$Matches[1]
      if ($n -gt $max) { $max = $n }
    }
  }
  return $max
}

# Un registro recien nacido. Nunca se escribe solo: se escribe cuando se aparta el primer
# numero, para que un registro en disco signifique siempre "aqui hubo una base".
function New-RegistroVacio {
  return [pscustomobject]@{
    version      = $BASES_VERSION
    sufijo       = $BASES_SUFIJO
    ultimoNumero = ($BASES_PRIMER_NUMERO - 1)
    abierta      = $null
    cerradas     = @()
  }
}

# Se asegura de que el objeto leido del JSON tenga todos los campos, y de que las listas sean
# listas de verdad. ConvertFrom-Json de PowerShell 5.1 devuelve un solo objeto -- no un array
# de uno -- cuando la lista trae un elemento, y eso reventaria el .Count de mas abajo.
function Repair-FormaRegistro {
  param($Registro)
  foreach ($campo in @('version','sufijo','ultimoNumero','abierta','cerradas')) {
    if (-not $Registro.PSObject.Properties[$campo]) {
      Add-Member -InputObject $Registro -NotePropertyName $campo -NotePropertyValue $null
    }
  }
  if ($null -eq $Registro.version)   { $Registro.version = $BASES_VERSION }
  if ([string]::IsNullOrWhiteSpace([string]$Registro.sufijo)) { $Registro.sufijo = $BASES_SUFIJO }
  if ($null -eq $Registro.ultimoNumero) { $Registro.ultimoNumero = ($BASES_PRIMER_NUMERO - 1) }
  $Registro.ultimoNumero = [int]$Registro.ultimoNumero
  $Registro.cerradas = @($Registro.cerradas | Where-Object { $null -ne $_ })
  if ($null -ne $Registro.abierta) {
    if (-not $Registro.abierta.PSObject.Properties['clientes']) {
      Add-Member -InputObject $Registro.abierta -NotePropertyName 'clientes' -NotePropertyValue @()
    }
    $Registro.abierta.clientes = @($Registro.abierta.clientes | Where-Object { $null -ne $_ })
  }
  return $Registro
}

# Lee la memoria. Si el archivo no esta, devuelve uno vacio EN MEMORIA y no toca el disco.
#
# Si el archivo esta y no se puede leer, se ABORTA con el mensaje. No se reinicia el registro
# en silencio: reiniciarlo repartiria numeros ya usados, que es exactamente el fallo que este
# archivo viene a cerrar. Un registro roto se arregla a mano -- es texto y se lee con los ojos.
function Get-RegistroBases {
  param([Parameter(Mandatory=$true)][string]$Carpeta)
  $ruta = Get-RutaRegistro -Carpeta $Carpeta
  if (-not (Test-Path -LiteralPath $ruta)) { return (New-RegistroVacio) }

  $texto = Get-Content -LiteralPath $ruta -Raw -Encoding UTF8
  if ($null -eq $texto -or $texto.Trim() -eq "") {
    throw ("El registro de bases esta VACIO: $ruta`n" +
           "       No se reinicia solo, porque repartiria numeros ya usados.`n" +
           "       Mira que documentos hay en la carpeta y escribe a mano el ultimo numero.")
  }
  try { $r = $texto | ConvertFrom-Json }
  catch {
    throw ("El registro de bases no se puede leer: $ruta`n" +
           "       $($_.Exception.Message)`n" +
           "       No se reinicia solo, porque repartiria numeros ya usados. Arreglalo a mano.")
  }
  return (Repair-FormaRegistro $r)
}

# Regla 4: entero o nada. Se escribe a un temporal, se RELEE para comprobar que quedo bien, y
# solo entonces se renombra encima del bueno. Si la maquina se cae a medio guardar, lo que
# queda a medias es el temporal y el registro bueno sigue intacto.
function Save-RegistroBases {
  param(
    [Parameter(Mandatory=$true)][string]$Carpeta,
    [Parameter(Mandatory=$true)]$Registro
  )
  if (-not (Test-Path -LiteralPath $Carpeta)) {
    New-Item -ItemType Directory -Path $Carpeta -Force | Out-Null
  }
  $ruta = Get-RutaRegistro -Carpeta $Carpeta
  $tmp  = "$ruta.nuevo"
  ($Registro | ConvertTo-Json -Depth 6) | Out-File -FilePath $tmp -Encoding utf8 -Force
  $comprobar = Get-Content -LiteralPath $tmp -Raw -Encoding UTF8
  $null = $comprobar | ConvertFrom-Json     # si esto revienta, el bueno no se toca
  Move-Item -LiteralPath $tmp -Destination $ruta -Force
}

# Aparta el numero siguiente y lo deja escrito ANTES de que exista ningun documento (regla 1).
# Devuelve el nombre y la ruta que le tocan; crear el .docx es cosa de quien llama.
#
# Si la creacion falla, quien llama tiene que cerrar la base con Close-BaseAbierta y motivo:
# el numero se queda quemado, que es justo lo que se quiere.
function New-BaseAbierta {
  param(
    [Parameter(Mandatory=$true)][string]$Carpeta,
    [Parameter(Mandatory=$true)]$Registro,
    [scriptblock]$Log
  )
  if (-not $Log) { $Log = { param($m) Write-Host $m } }
  if ($null -ne $Registro.abierta) {
    throw "Ya hay una base abierta (BASE $('{0:D3}' -f [int]$Registro.abierta.numero)). Cierrala antes de abrir otra."
  }

  $n = [int]$Registro.ultimoNumero

  # Regla 2: si la carpeta va por delante del registro, gana la carpeta y se dice.
  $enCarpeta = Get-MaxNumeroEnCarpeta -Carpeta $Carpeta
  if ($enCarpeta -gt $n) {
    & $Log "AVISO: la carpeta trae un BASE $('{0:D3}' -f $enCarpeta) y el registro iba por el $('{0:D3}' -f $n)."
    & $Log "       Manda el mas alto: el contador nunca retrocede."
    $n = $enCarpeta
  }

  # Regla 3: jamas se sobrescribe. Se avanza hasta un nombre libre.
  do {
    $n++
    $nombre = Get-NombreBase -Numero $n -Sufijo $Registro.sufijo
    $ruta   = Join-Path $Carpeta $nombre
    $ocupado = Test-Path -LiteralPath $ruta
    if ($ocupado) { & $Log "AVISO: '$nombre' ya existe. No se sobrescribe: se pasa al numero siguiente." }
  } while ($ocupado)

  $Registro.ultimoNumero = $n
  $Registro.abierta = [pscustomobject]@{
    numero   = $n
    archivo  = $nombre
    creada   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    clientes = @()
  }
  Save-RegistroBases -Carpeta $Carpeta -Registro $Registro

  return [pscustomobject]@{ Numero = $n; Archivo = $nombre; Ruta = $ruta }
}

# Anota un cliente dentro de la base abierta. Se llama DESPUES de guardar el documento: si la
# maquina se cayera entre las dos cosas, el cliente estaria en el documento y no en la memoria,
# que es el lado seguro -- se ve al abrir el documento y se corrige a mano.
function Add-ClienteABase {
  param(
    [Parameter(Mandatory=$true)][string]$Carpeta,
    [Parameter(Mandatory=$true)]$Registro,
    [Parameter(Mandatory=$true)][string]$Cuenta,
    [string]$Razon = "",
    [string]$Detalle = ""
  )
  if ($null -eq $Registro.abierta) { throw "No hay ninguna base abierta a la que anotar el cliente $Cuenta." }
  $Registro.abierta.clientes = @($Registro.abierta.clientes) + [pscustomobject]@{
    cuenta  = $Cuenta
    razon   = $Razon
    hora    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    detalle = $Detalle
  }
  Save-RegistroBases -Carpeta $Carpeta -Registro $Registro
  return @($Registro.abierta.clientes).Count
}

function Close-BaseAbierta {
  param(
    [Parameter(Mandatory=$true)][string]$Carpeta,
    [Parameter(Mandatory=$true)]$Registro,
    [string]$Motivo = 'completa'
  )
  if ($null -eq $Registro.abierta) { return $null }
  $a = $Registro.abierta
  $cerrada = [pscustomobject]@{
    numero   = [int]$a.numero
    archivo  = [string]$a.archivo
    clientes = @($a.clientes).Count
    creada   = [string]$a.creada
    cerrada  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    motivo   = $Motivo
  }
  $Registro.cerradas = @($Registro.cerradas) + $cerrada
  $Registro.abierta  = $null
  Save-RegistroBases -Carpeta $Carpeta -Registro $Registro
  return $cerrada
}

# Regla 5. La base abierta tiene que seguir estando en el disco. Si no esta, NO se decide nada
# aqui: se devuelve el problema y quien llama aborta y pregunta.
function Test-BaseAbierta {
  param(
    [Parameter(Mandatory=$true)][string]$Carpeta,
    [Parameter(Mandatory=$true)]$Registro
  )
  if ($null -eq $Registro.abierta) { return [pscustomobject]@{ Hay = $false; Existe = $false; Ruta = ""; Problema = "" } }
  $ruta = Join-Path $Carpeta ([string]$Registro.abierta.archivo)
  if (Test-Path -LiteralPath $ruta) {
    return [pscustomobject]@{ Hay = $true; Existe = $true; Ruta = $ruta; Problema = "" }
  }
  $p = ("La memoria dice que hay una base abierta, BASE $('{0:D3}' -f [int]$Registro.abierta.numero) con " +
        "$(@($Registro.abierta.clientes).Count) cliente(s), pero el documento NO esta en el disco:`n" +
        "       $ruta`n" +
        "       No se empieza otra con ese numero ni se sigue como si nada. O lo devuelves a su`n" +
        "       sitio, o cierras esa base a mano con -CerrarBase y se abre la siguiente.")
  return [pscustomobject]@{ Hay = $true; Existe = $false; Ruta = $ruta; Problema = $p }
}

# Donde esta ya escrita esta cuenta, si es que esta. Es lo que impide duplicar un cliente al
# relanzar una lista, y sustituye al -Reanudar que se quito el 06/09/2026.
function Find-ClienteEnBases {
  param(
    [Parameter(Mandatory=$true)]$Registro,
    [Parameter(Mandatory=$true)][string]$Cuenta
  )
  if ($null -ne $Registro.abierta) {
    foreach ($c in @($Registro.abierta.clientes)) {
      if ([string]$c.cuenta -eq $Cuenta) {
        return [pscustomobject]@{ Donde = 'abierta'; Numero = [int]$Registro.abierta.numero; Archivo = [string]$Registro.abierta.archivo }
      }
    }
  }
  foreach ($b in @($Registro.cerradas)) {
    # Las cerradas solo guardan el conteo, no la lista: la busqueda de arriba es la que cuenta.
    # Este bucle existe para el dia en que se quiera guardar tambien sus cuentas.
    if ($b.PSObject.Properties['cuentas']) {
      foreach ($cu in @($b.cuentas)) {
        if ([string]$cu -eq $Cuenta) {
          return [pscustomobject]@{ Donde = 'cerrada'; Numero = [int]$b.numero; Archivo = [string]$b.archivo }
        }
      }
    }
  }
  return $null
}

function Show-EstadoBases {
  param([Parameter(Mandatory=$true)][string]$Carpeta)
  $reg = Get-RegistroBases -Carpeta $Carpeta
  "Carpeta:  $Carpeta"
  "Registro: $(Get-RutaRegistro -Carpeta $Carpeta)"
  ""
  if ($null -eq $reg.abierta) {
    "Base abierta: ninguna. La siguiente seria BASE $('{0:D3}' -f ([int]$reg.ultimoNumero + 1))."
  } else {
    $a = $reg.abierta
    $n = @($a.clientes).Count
    "Base abierta: $($a.archivo)  --  $n de $BASES_POR_BASE cliente(s), desde $($a.creada)"
    $i = 0
    foreach ($c in @($a.clientes)) {
      $i++
      "   {0,2}. {1}  {2}   {3}" -f $i, $c.cuenta, $c.razon, $c.hora
    }
    $ruta = Join-Path $Carpeta ([string]$a.archivo)
    if (-not (Test-Path -LiteralPath $ruta)) { "   AVISO: ese documento NO esta en el disco." }
  }
  ""
  $cer = @($reg.cerradas)
  if ($cer.Count -eq 0) { "Bases cerradas: ninguna todavia." }
  else {
    "Bases cerradas: $($cer.Count)"
    foreach ($b in $cer) { "   $($b.archivo)  --  $($b.clientes) cliente(s), cerrada $($b.cerrada) ($($b.motivo))" }
  }
}
