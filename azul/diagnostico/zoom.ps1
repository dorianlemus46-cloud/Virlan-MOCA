# Recorta y amplia un pedazo de la ventana de Azul, con una rejilla de coordenadas encima.
#
# Para que sirve: los clics reales van en coordenadas RELATIVAS a la esquina de la ventana
# (asi es como esta clavado el enlace "Cliente:" en 529,120). Medir esas coordenadas a ojo
# sobre una captura de 1366x768 es adivinar; aqui se amplia la zona y se dibuja una rejilla
# rotulada con la coordenada de origen, para leer el centro de un icono al pixel.
#
# SOLO LECTURA. Usa PrintWindow igual que captura.ps1: no sube Azul al frente, no mueve el
# mouse, no pulsa nada. Funciona aunque Azul este detras de otras ventanas (pero NO
# minimizado: ahi Windows la aparca en -32000 y la captura sale basura).
#
#   .\zoom.ps1 -X 700 -Y 100 -Ancho 240 -Alto 45
#   .\zoom.ps1 -Fuente "..\salidas\capturas\cliente_20260902_130319.png" -X 700 -Y 100 -Ancho 240 -Alto 45
#   .\zoom.ps1                     -> la ventana entera, sin ampliar (Escala 1)
#
# NOTA: mantener este archivo en ASCII puro, igual que el resto de los .ps1 del proyecto.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

param(
  # Region de origen, en coordenadas relativas a la esquina superior izquierda de la ventana.
  [int]$X = 0,
  [int]$Y = 0,
  # 0 = hasta el borde de la ventana
  [int]$Ancho = 0,
  [int]$Alto = 0,
  [int]$Escala = 6,
  # PNG ya capturado, en vez de capturar de nuevo. Sirve para medir sobre una captura vieja
  # sin depender de que Azul este abierto y en el mismo estado.
  [string]$Fuente = "",
  [string]$Ruta = "",
  # Rejilla cada N pixeles de ORIGEN, con rotulo cada 5 lineas. 0 la apaga.
  [int]$Rejilla = 10
)
$ErrorActionPreference = 'Stop'

$raiz = Split-Path $PSScriptRoot -Parent
. (Join-Path $raiz 'captura.ps1')

# ---- imagen de origen -------------------------------------------------------
if ($Fuente -ne "") {
  if (-not (Test-Path -LiteralPath $Fuente)) { throw "No encuentro la imagen: $Fuente" }
  # Image::FromFile deja el archivo BLOQUEADO mientras el objeto viva. Se copian los pixeles
  # a un Bitmap propio y se suelta el original enseguida, igual que en Test-CapturasDistintas.
  $img = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Fuente).Path)
  $src = New-Object System.Drawing.Bitmap $img
  $img.Dispose()
  Write-Host "Fuente: $Fuente"
} else {
  $hwnd = Get-AzulHwnd
  $src = [AzulShot]::Capture($hwnd)
  if ($null -eq $src) { throw "La ventana de Azul no tiene tamano utilizable." }
  $r = Get-AzulRect -Hwnd $hwnd
  Write-Host "Ventana de Azul en pantalla: ($($r.Left),$($r.Top)) $($src.Width)x$($src.Height)"
}

try {
  # ---- region, recortada a lo que de verdad existe --------------------------
  if ($X -lt 0) { $X = 0 }
  if ($Y -lt 0) { $Y = 0 }
  if ($X -ge $src.Width -or $Y -ge $src.Height) {
    throw "El origen ($X,$Y) cae fuera de la imagen de $($src.Width)x$($src.Height)."
  }
  if ($Ancho -le 0) { $Ancho = $src.Width  - $X }
  if ($Alto  -le 0) { $Alto  = $src.Height - $Y }
  if ($X + $Ancho -gt $src.Width)  { $Ancho = $src.Width  - $X }
  if ($Y + $Alto  -gt $src.Height) { $Alto  = $src.Height - $Y }
  if ($Escala -lt 1) { $Escala = 1 }

  if ($Ruta -eq "") {
    $Ruta = Join-Path $raiz ("salidas\zoom\zoom_{0}_{1}x{2}_{3}x{4}.png" -f `
            (Get-Date -Format 'HHmmss'), $X, $Y, $Ancho, $Alto)
  }
  $dir = Split-Path $Ruta -Parent
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

  $dst = New-Object System.Drawing.Bitmap (($Ancho * $Escala), ($Alto * $Escala))
  $g = [System.Drawing.Graphics]::FromImage($dst)
  try {
    # NearestNeighbor: al ampliar, un pixel de origen tiene que quedar como un cuadrado
    # nitido. Con interpolacion suave los bordes del icono se difuminan y el centro deja de
    # ser medible, que es justo para lo que existe este script.
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
    $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
    $rd = New-Object System.Drawing.Rectangle 0, 0, ($Ancho * $Escala), ($Alto * $Escala)
    $rs = New-Object System.Drawing.Rectangle $X, $Y, $Ancho, $Alto
    $g.DrawImage($src, $rd, $rs, [System.Drawing.GraphicsUnit]::Pixel)

    # ---- rejilla rotulada con la coordenada de ORIGEN ------------------------
    if ($Rejilla -gt 0 -and $Escala -ge 3) {
      $fino  = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(70, 255, 0, 0)), 1
      $grue  = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(170, 255, 0, 0)), 1
      $tinta = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(230, 255, 0, 0))
      $fondo = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(190, 255, 255, 255))
      $fnt   = New-Object System.Drawing.Font "Consolas", 9
      try {
        $primeroX = [int][Math]::Ceiling($X / [double]$Rejilla) * $Rejilla
        for ($sx = $primeroX; $sx -lt ($X + $Ancho); $sx += $Rejilla) {
          $px = ($sx - $X) * $Escala
          $rot = (($sx % ($Rejilla * 5)) -eq 0)
          $g.DrawLine($(if ($rot) { $grue } else { $fino }), $px, 0, $px, $dst.Height)
          if ($rot) {
            $t = "$sx"
            $sz = $g.MeasureString($t, $fnt)
            $g.FillRectangle($fondo, ($px + 1), 0, $sz.Width, $sz.Height)
            $g.DrawString($t, $fnt, $tinta, ($px + 1), 0)
          }
        }
        $primeroY = [int][Math]::Ceiling($Y / [double]$Rejilla) * $Rejilla
        for ($sy = $primeroY; $sy -lt ($Y + $Alto); $sy += $Rejilla) {
          $py = ($sy - $Y) * $Escala
          $rot = (($sy % ($Rejilla * 5)) -eq 0)
          $g.DrawLine($(if ($rot) { $grue } else { $fino }), 0, $py, $dst.Width, $py)
          if ($rot) {
            $t = "$sy"
            $sz = $g.MeasureString($t, $fnt)
            $g.FillRectangle($fondo, 0, ($py + 1), $sz.Width, $sz.Height)
            $g.DrawString($t, $fnt, $tinta, 0, ($py + 1))
          }
        }
      } finally {
        $fino.Dispose(); $grue.Dispose(); $tinta.Dispose(); $fondo.Dispose(); $fnt.Dispose()
      }
    }
  } finally { $g.Dispose() }

  try { $dst.Save($Ruta, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $dst.Dispose() }

  Write-Host "Region: x=$X y=$Y ${Ancho}x${Alto}  escala=${Escala}x  rejilla=$Rejilla"
  Write-Host "Zoom: $Ruta"
  return $Ruta
} finally { $src.Dispose() }
