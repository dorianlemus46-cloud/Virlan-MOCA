# La lista de ordenes: de donde salen las cuentas que hay que buscar en Azul.
#
# Lo carga con dot-source quien la necesite:  . (Join-Path $PSScriptRoot 'lista.ps1')
#
# Desde el 05/09/2026 se lee un CSV en disco y NO el libro de Excel abierto. Dorian lo
# exporta con Guardar como sobre la hoja resultData, sin borrar columnas: aqui se busca por
# ENCABEZADO, igual que hacia el lector de Excel, asi que las otras siete no estorban.
#
# Vivia duplicado, casi linea por linea, en lote.ps1 (Get-Ordenes) y en buscar_cuenta.ps1
# (Get-OrdenDeExcel). Dos copias de la misma logica acaban divergiendo -- es el mismo motivo
# por el que revisar_csv.ps1 tiene su propio archivo.
#
#   Get-ListaOrdenes -Ruta ...\lista_ordenes.csv          la lista entera, sin duplicados
#   Get-OrdenDeLista -Ruta ...\lista_ordenes.csv -Fila 6  una sola, como la numera la hoja
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

# Excel escribe el separador de lista de Windows: coma en configuracion inglesa, punto y coma
# en varias configuraciones en espanol. Si no coincide con el que se supone, el archivo entero
# se lee como UNA sola columna, no aparece el encabezado 'Cuenta', y el sintoma es un error
# raro en vez de "el separador no es el que crees". Se decide mirando la linea de encabezados,
# que no lleva campos entrecomillados con comas dentro.
function Get-DelimitadorLista {
  param([Parameter(Mandatory=$true)][string]$Encabezado)
  $comas       = ([regex]::Matches($Encabezado, ',')).Count
  $puntoYComas = ([regex]::Matches($Encabezado, ';')).Count
  if ($puntoYComas -gt $comas) { return ';' }
  return ','
}

# Estructura de UTF-8: un byte lider dice cuantas continuaciones 10xxxxxx le siguen. El texto
# ANSI con acentos casi nunca la cumple -- 'MIA' con la I acentuada son los bytes CD 41, y CD
# exige una continuacion entre 80 y BF, que 'A' (41) no es. Por eso esto distingue de verdad
# entre un CSV en ANSI y uno en UTF-8, en vez de adivinar.
function Test-BytesUtf8 {
  param([byte[]]$Bytes)
  $i = 0
  while ($i -lt $Bytes.Length) {
    $b = $Bytes[$i]
    if ($b -lt 0x80) { $i++; continue }
    if     ($b -ge 0xC2 -and $b -le 0xDF) { $n = 1 }
    elseif ($b -ge 0xE0 -and $b -le 0xEF) { $n = 2 }
    elseif ($b -ge 0xF0 -and $b -le 0xF4) { $n = 3 }
    else { return $false }
    if ($i + $n -ge $Bytes.Length) { return $false }
    for ($k = 1; $k -le $n; $k++) {
      $c = $Bytes[$i + $k]
      if ($c -lt 0x80 -or $c -gt 0xBF) { return $false }
    }
    $i += $n + 1
  }
  return $true
}

# El texto de la lista, decodificado a conciencia en vez de a la suerte.
#
# MEDIDO el 05/09/2026 sobre las dos opciones que ofrece Excel y sobre la tercera que no
# ofrece: "CSV (delimitado por comas)" sale en ANSI y "CSV UTF-8" sale con BOM -- las dos se
# leen bien. La que rompe es UTF-8 SIN BOM: 'DISTRIBUIDORA RIO' con la I acentuada llegaba
# como 'DISTRIBUIDORA RAO' con la basura de en medio, EN SILENCIO. Excel no exporta asi, pero
# el Bloc de notas de Windows 11 guarda asi por defecto, y abrir el CSV para comprobarlo con
# los ojos es justo lo que se recomienda hacer antes de una corrida.
#
# Importa porque esa cadena tiene dos destinos: la guarda que compara contra la fila de Azul
# --que fallaria hacia el lado seguro, abortando-- y el encabezado del documento del cliente,
# donde el nombre saldria mal escrito.
function Read-TextoLista {
  param([Parameter(Mandatory=$true)][string]$Ruta, [scriptblock]$Log)
  if (-not $Log) { $Log = { param($m) Write-Host $m } }
  $bytes = [IO.File]::ReadAllBytes($Ruta)
  if ($bytes.Length -eq 0) { throw "La lista de ordenes '$Ruta' esta vacia." }

  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    return [Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
  }
  $altos = $false
  foreach ($b in $bytes) { if ($b -gt 127) { $altos = $true; break } }
  if (-not $altos) { return [Text.Encoding]::ASCII.GetString($bytes) }

  if (Test-BytesUtf8 -Bytes $bytes) {
    & $Log "AVISO: la lista viene en UTF-8 SIN BOM, que no es lo que exporta Excel."
    & $Log "       Se lee como UTF-8 para no romper los acentos. Si el archivo se guardo con"
    & $Log "       el Bloc de notas, es normal; vuelve a exportarlo de Excel si dudas."
    return (New-Object Text.UTF8Encoding($false)).GetString($bytes)
  }
  return [Text.Encoding]::Default.GetString($bytes)
}

# El nombre real de una columna, buscado por patron sobre los encabezados que trajo el CSV.
# Hace falta porque 'Razon Social' viene con acento y este archivo es ASCII puro: no se puede
# escribir $fila.'Razon Social' con la tilde, asi que se recorren las propiedades y se compara.
function Get-NombreColumna {
  param($Fila, [Parameter(Mandatory=$true)][string]$Patron)
  foreach ($p in $Fila.PSObject.Properties) {
    if ($p.Name.Trim() -like $Patron) { return $p.Name }
  }
  return ""
}

function Import-ListaCruda {
  param([Parameter(Mandatory=$true)][string]$Ruta, [scriptblock]$Log)
  if (-not (Test-Path -LiteralPath $Ruta)) {
    throw ("No encuentro la lista de ordenes: $Ruta`n" +
           "       Exportala desde Excel: hoja resultData > Guardar como > CSV.")
  }
  $texto   = Read-TextoLista -Ruta $Ruta -Log $Log
  $renglon = @($texto -split "`r?`n")
  $del     = Get-DelimitadorLista -Encabezado $renglon[0]
  $filas   = @($renglon | ConvertFrom-Csv -Delimiter $del)
  return @{ Filas = $filas; Delimitador = $del }
}

# La lista entera, ya sin cuentas repetidas y con las filas malas denunciadas en voz alta.
#
# -Log recibe un bloque para informar. lote.ps1 le pasa el suyo, que antepone la hora; quien
# no pase nada escribe por consola tal cual.
function Get-ListaOrdenes {
  param(
    [Parameter(Mandatory=$true)][string]$Ruta,
    [scriptblock]$Log
  )
  if (-not $Log) { $Log = { param($m) Write-Host $m } }

  $crudo = Import-ListaCruda -Ruta $Ruta -Log $Log
  $filas = $crudo.Filas
  if ($filas.Count -eq 0) { throw "La lista de ordenes '$Ruta' no trae ninguna fila de datos." }

  $colCuenta = Get-NombreColumna -Fila $filas[0] -Patron 'Cuenta'
  $colRazon  = Get-NombreColumna -Fila $filas[0] -Patron 'Raz*n Social'
  if ($colCuenta -eq "") {
    throw ("La lista '$Ruta' no tiene una columna 'Cuenta'.`n" +
           "       Se leyo con el separador '$($crudo.Delimitador)' y salieron estas columnas:`n" +
           "       " + (($filas[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ' | ') + "`n" +
           "       Si eso parece una sola columna, el separador del archivo no es el que se dedujo.")
  }
  # La comprobacion de la razon social es la guarda que impide entrar al cliente equivocado, y
  # es condicional: sin el dato, no se aplica. Callarlo seria dejarla apagada en silencio.
  if ($colRazon -eq "") {
    & $Log "AVISO: la lista no trae columna 'Razon Social'. La comprobacion contra el cliente equivocado NO se aplicara."
  }

  $lista  = New-Object System.Collections.ArrayList
  $vistas = @{}
  $malas  = 0
  # La fila 1 de la hoja son los encabezados, asi que la primera de datos es la 2. Se cuenta
  # asi para que -Fila 6 siga queriendo decir "la sexta fila que se ve en la hoja".
  $fila = 1
  foreach ($r in $filas) {
    $fila++
    $num = ([string]$r.$colCuenta).Trim()
    $raz = ""
    if ($colRazon -ne "") { $raz = ([string]$r.$colRazon).Trim() }

    # Fila del todo vacia: es la cola que Excel arrastra al exportar. No se avisa.
    if ($num -eq "" -and $raz -eq "") { continue }

    # A partir de aqui, TODO descarte se dice en voz alta. Con el libro de Excel vivo esto
    # solo se comia el encabezado; con un CSV exportado a mano, una cuenta que salio en
    # notacion cientifica (5,95E+08) o sin el cero de la izquierda desaparecia del lote sin
    # que nada lo dijera, y no se notaba hasta que faltaba un cliente en la base.
    if ($num -eq "") {
      & $Log "  fila $fila  DESCARTADA: '$raz' viene sin numero de cuenta."
      $malas++
      continue
    }
    if ($num -notmatch '^\d{4,15}$') {
      & $Log "  fila $fila  DESCARTADA: '$num' no es un numero de cuenta ($raz)."
      & $Log "                Si en Excel se ve bien, revisa como quedo en el CSV: notacion cientifica o ceros perdidos."
      $malas++
      continue
    }
    # Una cuenta repetida es la MISMA base de datos. Meterla dos veces seria un error
    # silencioso, y ademas el corte de 10 dejaria de significar 10 clientes.
    if ($vistas.ContainsKey($num)) {
      & $Log "  fila $fila  $num $raz  -> repetida (ya salio en la fila $($vistas[$num])), se salta"
      continue
    }
    $vistas[$num] = $fila
    [void]$lista.Add([pscustomobject]@{ Fila = $fila; Cuenta = $num; Razon = $raz })
  }

  $resumen = "Lista: $(Split-Path $Ruta -Leaf), $($lista.Count) cliente(s) distinto(s)"
  if ($malas -gt 0) { $resumen += ", $malas fila(s) DESCARTADA(S)" }
  & $Log "$resumen."
  return $lista
}

# Una sola orden, por el numero de fila tal como lo ensena la hoja: la 1 son los encabezados,
# asi que -Fila 6 es la sexta fila que se ve.
#
# Se apoya en Get-ListaOrdenes a proposito, en vez de leer el archivo por su cuenta: asi la
# fila suelta pasa por las MISMAS comprobaciones que las del lote. Que este camino y el del
# lote validen distinto es exactamente el problema que este archivo viene a cerrar.
function Get-OrdenDeLista {
  param(
    [Parameter(Mandatory=$true)][string]$Ruta,
    [Parameter(Mandatory=$true)][int]$Fila
  )
  $todo = Get-ListaOrdenes -Ruta $Ruta
  $o = $todo | Where-Object { $_.Fila -eq $Fila } | Select-Object -First 1
  if ($null -eq $o) {
    throw ("La fila $Fila no trae una cuenta utilizable en '$Ruta'.`n" +
           "       O esta fuera de la lista, o se descarto, o es una cuenta repetida: el motivo`n" +
           "       sale en los avisos de arriba. Los datos empiezan en la fila 2.")
  }
  return $o
}
