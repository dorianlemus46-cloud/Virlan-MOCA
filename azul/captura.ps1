# Auxiliar de captura y clic. Lo carga azul_fast.ps1 con dot-source; tambien se puede correr
# solo para sacar una foto suelta de la ventana de Azul.
#
# Aqui vive TODO lo que toca el sistema operativo (capturar, traer al frente, mover el mouse).
# azul_fast.ps1 se queda solo con el puente. Separarlos deja claro donde esta lo que toca
# la maquina de Dorian y donde lo que solo lee.
#
# Save-AzulShot captura SOLO la ventana de Azul, por su hwnd, con PrintWindow: no la trae al
# frente, no mueve el mouse, y funciona aunque Azul este detras de otras ventanas.
#
# Invoke-AzulClick SI toca la ventana: la sube al frente y hace un clic real. Hace falta
# para la banda de "Cliente:", que es contenido de un navegador incrustado y por eso el
# Java Access Bridge no la ve ni la puede pulsar. Autorizado por Dorian el 02/09/2026.
# Devuelve el mouse a donde estaba, para no dejar el cursor movido.
#
# El codigo nativo sale de respaldo-capturas\azul_tools.ps1, que ya usaba la bandera 2
# (PW_RENDERFULLCONTENT), la unica que funciona con ventanas que se dibujan solas como las
# de Swing. Se copio en vez de cargar aquel archivo porque trae rutas de una sesion vieja y
# funciones de clic que aqui no pintan.
#
#   . .\captura.ps1
#   $h = Get-AzulHwnd
#   $p = Save-AzulShot -Hwnd $h -Ruta "C:\...\cliente.png"
#   Test-CapturaValida -Ruta $p
#
# NOTA: mantener este archivo en ASCII puro, igual que azul_fast.ps1. PowerShell 5.1 lee
# los .ps1 sin BOM como ANSI y rompe los acentos en silencio.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not ("AzulShot" -as [type])) {
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

// Relleno hasta el tamano de MOUSEINPUT, que es el miembro mas grande de la union INPUT.
// Sin los dos int de cola, SendInput recibe una estructura corta y no hace nada, en silencio.
// Los tamanos salen bien tanto en 32 como en 64 bits, que importa porque este archivo lo
// cargan los dos: azul_fast.ps1 y buscar_cuenta.ps1 corren en 32, zoom.ps1 en 64.
[StructLayout(LayoutKind.Sequential)]
public struct KBIN {
  public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo;
  public int relleno1; public int relleno2;
}
[StructLayout(LayoutKind.Sequential)]
public struct INP { public uint type; public KBIN ki; }

public class AzulShot {
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
  [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool at);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
  [DllImport("user32.dll")] static extern void mouse_event(uint f, uint dx, uint dy, uint d, IntPtr e);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }

  // AQUI VIVIA Front(), Y SE BORRO EL 06/09/2026. No lo reescribas.
  //
  // Era identica a Frente() salvo en una linea: llamaba a ShowWindow(h, 9) SIEMPRE, y
  // SW_RESTORE sobre una ventana maximizada la des-maximiza -- cambia su posicion y su tamano.
  // Como el clic se da justo despues de medir la geometria, la ventana se movia entre medir y
  // pulsar y el clic caia fuera del boton. Costo una sesion de diagnostico el 03/09/2026
  // (docs/bitacora.md, "El clic caia fuera del boton que se acababa de medir"; docs/trampas.md
  // trampa 7).
  //
  // Se arreglo creando Frente(), pero Front() se quedo aqui: compilada, publica, sin una sola
  // llamada, y con un nombre a una letra del bueno. Lo unico que impedia volver a llamarla por
  // error era un comentario. Eso no es una guarda, es una trampa esperando.
  //
  // USAR SIEMPRE Frente(), abajo.

  [DllImport("user32.dll", SetLastError=true)] static extern uint SendInput(uint n, INP[] p, int cb);
  const uint KEYUP = 0x0002, UNICODE = 0x0004;

  // Teclado real. Hace falta porque setTextContents del Java Access Bridge devuelve exito y
  // deja el campo "ID de CF" vacio (medido el 03/09/2026): el puente sabe LEER ese campo pero
  // no escribirlo. Se manda el caracter por scancode Unicode y no por tecla virtual, para no
  // depender de la distribucion del teclado.
  //
  // OJO: esto va a donde este el foco del sistema. Quien llame tiene que haber traido Azul al
  // frente Y haber enfocado el campo por el puente, y tiene que releer el campo despues.
  public static void Escribir(string s) {
    foreach (char c in s) {
      INP[] a = new INP[2];
      a[0].type = 1; a[0].ki.wScan = (ushort)c; a[0].ki.dwFlags = UNICODE;
      a[1].type = 1; a[1].ki.wScan = (ushort)c; a[1].ki.dwFlags = UNICODE | KEYUP;
      SendInput(2, a, Marshal.SizeOf(typeof(INP)));
      // 40 ms y no 25: con 25 el campo "ID de CF" se quedaba a medias -- 3 de 9 digitos --
      // porque Azul no da abasto procesando las pulsaciones. Medido el 03/09/2026. Quien
      // llame tiene que comprobar el resultado igualmente; esto solo baja la frecuencia.
      System.Threading.Thread.Sleep(40);
    }
  }

  public static void Tecla(ushort vk) {
    INP[] a = new INP[2];
    a[0].type = 1; a[0].ki.wVk = vk;
    a[1].type = 1; a[1].ki.wVk = vk; a[1].ki.dwFlags = KEYUP;
    SendInput(2, a, Marshal.SizeOf(typeof(INP)));
    System.Threading.Thread.Sleep(25);
  }

  // LA UNICA forma de traer Azul al frente. Sube la ventana SIN cambiar su estado: solo
  // restaura si de verdad esta minimizada, que es el unico caso en que hace falta.
  //
  // La version anterior (Front(), borrada el 06/09/2026) llamaba a ShowWindow(9) siempre, y
  // SW_RESTORE sobre una ventana MAXIMIZADA la des-maximiza: cambia su posicion y su tamano.
  // Como Invoke-AzulClick sube la ventana justo antes de pulsar, se movia entre el momento de
  // medir y el de hacer clic, y el clic caia fuera del boton.
  // Medido el 03/09/2026: Azul aparecia unas veces en (0,0) 1366x768, otras maximizada en
  // (-8,-8) 1382x736, y otras en (0,113); ninguna la movia Dorian.
  public static bool Frente(IntPtr h) {
    for (int i = 0; i < 6; i++) {
      if (GetForegroundWindow() == h) return true;
      IntPtr fg = GetForegroundWindow();
      uint me = GetCurrentThreadId();
      uint tf = (fg == IntPtr.Zero) ? 0 : GetWindowThreadProcessId(fg, IntPtr.Zero);
      uint tt = GetWindowThreadProcessId(h, IntPtr.Zero);
      if (tf != 0 && tf != me) AttachThreadInput(me, tf, true);
      if (tt != me) AttachThreadInput(me, tt, true);
      if (IsIconic(h)) ShowWindow(h, 9);
      BringWindowToTop(h); SetForegroundWindow(h);
      if (tt != me) AttachThreadInput(me, tt, false);
      if (tf != 0 && tf != me) AttachThreadInput(me, tf, false);
      System.Threading.Thread.Sleep(350);
    }
    return GetForegroundWindow() == h;
  }

  public static void Click(int x, int y) {
    SetCursorPos(x, y); System.Threading.Thread.Sleep(120);
    mouse_event(0x0002, 0, 0, 0, IntPtr.Zero);   // izquierdo abajo
    System.Threading.Thread.Sleep(60);
    mouse_event(0x0004, 0, 0, 0, IntPtr.Zero);   // izquierdo arriba
  }

  public static Bitmap Capture(IntPtr h) {
    RECT r; GetWindowRect(h, out r);
    int w = r.Right - r.Left, hh = r.Bottom - r.Top;
    if (w <= 0 || hh <= 0) return null;
    Bitmap bmp = new Bitmap(w, hh);
    using (Graphics g = Graphics.FromImage(bmp)) {
      IntPtr hdc = g.GetHdc();
      PrintWindow(h, hdc, 2);   // 2 = PW_RENDERFULLCONTENT
      g.ReleaseHdc(hdc);
    }
    return bmp;
  }

  // Que fraccion de la rejilla cambio entre dos capturas, comparando EN MEMORIA.
  //
  // La via anterior daba un rodeo por el disco: se guardaban los dos PNG, se volvian a leer,
  // se clonaban a Bitmap y se muestreaban con GetPixel. GetPixel bloquea y desbloquea el
  // bitmap en CADA punto, y eran 1600 puntos por imagen, 3200 por comparacion. Con LockBits
  // el buffer se bloquea una vez y se indexa directo.
  //
  // El muestreo es el mismo que hacia Test-CapturasDistintas: misma rejilla, mismos puntos,
  // mismo criterio de igualdad (ARGB exacto). Lo que cambia es por donde pasan los pixeles,
  // no que se compara. Devuelve 1.0 -- "cambio todo" -- si los tamanos no coinciden, que es
  // lo que devolvia $true en la version vieja.
  // TRAMPA: PowerShell REDONDEA al castear a [int] y C# TRUNCA. [int]9.6 da 10; (int)9.6 da
  // 9. El muestreo anterior estaba escrito en PowerShell, asi que traducirlo a C# con un
  // cast movia la rejilla a otros puntos -- en silencio, y dando resultados parecidos, que
  // es lo peor. Se replica el redondeo de PowerShell (al par en los empates) para que caiga
  // EXACTAMENTE donde caia antes. Lo cazo la prueba que compara los dos muestreos.
  static int Punto(int idx, int tamano, int rejilla) {
    int v = (int)Math.Round((idx + 0.5) * tamano / rejilla, MidpointRounding.ToEven);
    if (v >= tamano) v = tamano - 1;
    if (v < 0) v = 0;
    return v;
  }

  public static double Diferencia(Bitmap a, Bitmap b, int rejilla) {
    if (a == null || b == null) return 1.0;
    if (a.Width != b.Width || a.Height != b.Height) return 1.0;
    if (rejilla <= 0) return 0.0;
    Rectangle rc = new Rectangle(0, 0, a.Width, a.Height);
    BitmapData da = a.LockBits(rc, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
    BitmapData db = null;
    try {
      db = b.LockBits(rc, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
      int dif = 0, total = rejilla * rejilla;
      for (int i = 0; i < rejilla; i++) {
        int x = Punto(i, a.Width, rejilla);
        for (int j = 0; j < rejilla; j++) {
          int y = Punto(j, a.Height, rejilla);
          // Stride por separado en cada imagen: son del mismo tamano, pero el relleno de
          // fila lo decide GDI+ y no hay por que dar por hecho que coinciden.
          int oa = y * da.Stride + x * 4;
          int ob = y * db.Stride + x * 4;
          if (Marshal.ReadInt32(da.Scan0, oa) != Marshal.ReadInt32(db.Scan0, ob)) dif++;
        }
      }
      return (double)dif / total;
    } finally {
      if (db != null) b.UnlockBits(db);
      a.UnlockBits(da);
    }
  }

  // ---- encontrar un icono MIRANDOLO, en vez de recordar donde estaba ----------
  // El icono del telefono con lupa es contenido de un navegador incrustado: el puente de
  // accesibilidad no lo ve, asi que no hay a quien preguntarle donde esta. Hasta el
  // 18/09/2026 se picaba en una posicion medida a mano el 03/09/2026 (791,121). Ese dia
  // fallo en una de las laptops: la barra de Azul aparecio corrida 55 pixeles a la derecha,
  // no se sabe por que, y el clic cayo antes del icono, en un campo vacio. Recuerda al
  // movimiento que ya estaba anotado para la pestana de Suscripciones, que se iba de
  // x=826 a x=895.
  //
  // Asi que no se recuerda una posicion: se reconoce el dibujo. La huella es el icono
  // reducido a tinta/no-tinta por luminancia, que aguanta cambios de fondo y de suavizado
  // mucho mejor que comparar colores exactos.
  //
  // Devuelve la ESQUINA del mejor encaje, no el centro, y con ella tres numeros en
  // diezmilesimas: cuanto coincide en total, cuanta de la tinta de la huella aparece, y
  // cuanto coincide el segundo mejor encaje lejos del primero. El que llama decide si eso
  // basta. Con un segundo encaje casi tan bueno hay dos cosas parecidas en pantalla, y lo
  // prudente entonces es no picar.
  public static string BuscaHuella(Bitmap b, string mascara, int mw, int mh,
                                   int x0, int y0, int x1, int y1, int umbral) {
    if (b == null || mascara == null) return "NADA|huella o imagen vacia";
    if (mw <= 0 || mh <= 0 || mascara.Length != mw * mh) return "NADA|la huella no mide mw*mh";
    if (x0 < 0) x0 = 0;
    if (y0 < 0) y0 = 0;
    if (x1 > b.Width  - mw) x1 = b.Width  - mw;
    if (y1 > b.Height - mh) y1 = b.Height - mh;
    if (x1 < x0 || y1 < y0) return "NADA|la zona de busqueda no cabe en la ventana";
    int ntinta = 0;
    for (int i = 0; i < mascara.Length; i++) if (mascara[i] == '1') ntinta++;
    if (ntinta == 0) return "NADA|la huella no tiene tinta";

    int ancho = b.Width, alto = b.Height;
    byte[] tinta = new byte[ancho * alto];
    Rectangle rc = new Rectangle(0, 0, ancho, alto);
    BitmapData d = b.LockBits(rc, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
    try {
      byte[] fila = new byte[d.Stride];
      for (int y = 0; y < alto; y++) {
        Marshal.Copy((IntPtr)(d.Scan0.ToInt64() + (long)y * d.Stride), fila, 0, d.Stride);
        int bas = y * ancho;
        for (int x = 0; x < ancho; x++) {
          int o = x * 4;   // BGRA: la media de los tres no depende del orden
          int lum = (fila[o] + fila[o + 1] + fila[o + 2]) / 3;
          tinta[bas + x] = (byte)(lum < umbral ? 1 : 0);
        }
      }
    } finally {
      b.UnlockBits(d);
    }

    int celdas = mw * mh;
    int mejor = -1, segundo = -1, mx = -1, my = -1, mejorTinta = 0;
    for (int y = y0; y <= y1; y++) {
      for (int x = x0; x <= x1; x++) {
        int acuerdo = 0, aciertos = 0;
        for (int j = 0; j < mh; j++) {
          int fo = (y + j) * ancho + x;
          int mo = j * mw;
          for (int i = 0; i < mw; i++) {
            bool t = mascara[mo + i] == '1';
            bool c = tinta[fo + i] != 0;
            if (t == c) acuerdo++;
            if (t && c) aciertos++;
          }
        }
        int s = (int)((long)acuerdo * 10000 / celdas);
        bool lejos = (mx < 0) || (Math.Abs(x - mx) > mw) || (Math.Abs(y - my) > mh);
        if (s > mejor) {
          if (lejos) segundo = mejor;
          mejor = s; mx = x; my = y;
          mejorTinta = (int)((long)aciertos * 10000 / ntinta);
        } else if (s > segundo && lejos) {
          segundo = s;
        }
      }
    }
    if (mx < 0) return "NADA|no se recorrio ninguna posicion";
    return "OK|" + mx + "," + my + "," + mejor + "," + mejorTinta + "," + segundo;
  }
}
'@
}

function Get-AzulHwnd {
  $p = Get-Process -Name jp2launcher -ErrorAction SilentlyContinue |
       Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if ($null -eq $p) { throw "Azul no esta corriendo (no hay ventana de jp2launcher)." }
  # Minimizada, Windows aparca la ventana en -32000,-32000 y la captura sale en basura.
  if ([AzulShot]::IsIconic($p.MainWindowHandle)) {
    throw "Azul esta minimizado. Restauralo (puede quedar detras de otras ventanas) y reintenta."
  }
  return $p.MainWindowHandle
}

function Save-AzulShot {
  param(
    [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
    [Parameter(Mandatory=$true)][string]$Ruta
  )
  $dir = Split-Path $Ruta -Parent
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $bmp = [AzulShot]::Capture($Hwnd)
  if ($null -eq $bmp) { throw "La ventana de Azul no tiene tamano utilizable." }
  try { $bmp.Save($Ruta, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $bmp.Dispose() }
  return $Ruta
}

function Get-AzulRect {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $r = New-Object AzulShot+RECT
  [void][AzulShot]::GetWindowRect($Hwnd, [ref]$r)
  return $r
}

# Clic real en coordenadas RELATIVAS a la esquina superior izquierda de la ventana de Azul.
# Relativas y no absolutas a proposito: si Dorian mueve la ventana, las absolutas apuntarian
# a otra cosa. Sube Azul al frente primero, porque un clic sobre una ventana tapada le llega
# a la que esta encima.
function Invoke-AzulClick {
  param(
    [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
    [Parameter(Mandatory=$true)][int]$X,
    [Parameter(Mandatory=$true)][int]$Y
  )
  $r = Get-AzulRect -Hwnd $Hwnd
  $px = $r.Left + $X
  $py = $r.Top  + $Y
  if ($px -lt $r.Left -or $px -gt $r.Right -or $py -lt $r.Top -or $py -gt $r.Bottom) {
    throw "El punto ($X,$Y) cae fuera de la ventana de Azul. No se hace clic a ciegas."
  }
  # Frente() no cambia el estado de la ventana. Si aqui se des-maximizara, se moveria el punto
  # que se acaba de calcular y el clic caeria fuera (trampa 7). Ver el comentario de Frente().
  if (-not [AzulShot]::Frente($Hwnd)) {
    throw "No se pudo traer Azul al frente. Sin eso el clic le llegaria a otra ventana."
  }
  # Se guarda donde estaba el cursor para devolverlo al final: el mouse es de Dorian.
  $antes = New-Object AzulShot+POINT
  $hay = [AzulShot]::GetCursorPos([ref]$antes)
  try { [AzulShot]::Click($px, $py) }
  finally { if ($hay) { [void][AzulShot]::SetCursorPos($antes.X, $antes.Y) } }
  return @{ X = $px; Y = $py }
}

# Guarda del clic: si la pantalla no cambio, el clic no dio en el enlace o el panel no
# abrio. Se compara una rejilla de pixeles entre el antes y el despues. Mas vale quedarse
# sin imagen que meter en el documento del cliente una foto de la pantalla equivocada.
#
# Compara dos PNG YA GUARDADOS, para comparaciones sueltas contra una captura vieja.
# Wait-PantallaEstable ya no pasa por aqui: dentro de su bucle los bitmaps estan en memoria
# y bajarlos a disco para volver a subirlos era el rodeo mas caro de la captura. Ver
# [AzulShot]::Diferencia, que hace el mismo muestreo sobre bitmaps vivos.

# ---- como se pica en ESTA maquina: en las posiciones de siempre, o mirando -----
# La lupita del telefono y el enlace "Cliente:" son contenido del navegador incrustado y el
# puente no los ve. Hay dos formas de saber donde picarlos:
#
#   - FIJA, lo de siempre, y lo que se hace si no se dice nada: las posiciones medidas a
#     mano, 791,121 la lupita y 500,120 el enlace (REGLAS.md seccion 5).
#   - MIRANDO: se busca la lupita en la ventana (Find-IconoLupa, aqui abajo) y el enlace se
#     saca a 291 pixeles a su izquierda.
#
# Por que no se mira en todas las maquinas: el 18/09/2026, en una de las dos laptops, la
# barra de Azul se corrio 55 pixeles y las posiciones fijas empezaron a fallar; ahi se pica
# mirando. En la otra las fijas aciertan, y Dorian pidio que ahi no se toque (18/09/2026).
# La huella de la lupita se saco en una sola maquina, y con otra escala de pantalla podria
# no reconocerse.
#
# Lo decide un archivo que NO va a git, azul\esta_maquina.json, con {"ClicMirando": true}.
# Si no existe, o no se entiende, se pica en las posiciones fijas: la maquina que no dice
# nada se queda exactamente como estaba.
$script:LUPA_FIJA   = @{ X = 791; Y = 121 }
$script:ENLACE_FIJO = @{ X = 500; Y = 120 }

function Test-ClicMirando {
  $f = Join-Path $PSScriptRoot 'esta_maquina.json'
  if (-not (Test-Path -LiteralPath $f)) { return $false }
  try {
    $c = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json
    return ($c.ClicMirando -eq $true)
  } catch {
    return $false
  }
}

# Huella del icono del telefono con lupa, sacada de Azul de verdad el 18/09/2026 con
# PrintWindow sobre la ventana maximizada en 1366x768. Es el dibujo reducido a tinta/no
# tinta: un 1 por cada pixel mas oscuro que Umbral. Se guarda como texto y no como PNG
# a proposito -- asi el proyecto sigue siendo archivos .ps1 legibles y no hay un binario
# suelto que nadie sepa de donde salio ni como volver a generarlo.
#
# DX y DY llevan de la esquina de la huella al centro del BOTON, que es donde hay que
# picar: el recorte empieza en 835,112 y el boton esta centrado en 846,121.
$script:HuellaLupa = @{
  Ancho  = 25
  Alto   = 19
  DX     = 11
  DY     = 9
  Umbral = 160
  Mascara = '0000000000000000000000000000001100000000000000000000011111000000000000000000001111100000000000000000000111110000111100000000000011110000110001000000000001110000010000010000000000111000001000001000000000011100000100001100000000000111000011001110000000000011100000011001000000000001111000000000010000000000011100010000000000000000000111111110000000000000000001111111000000000000000000011111000000000000000000000011000000000000000000000000000000000000000000000000000000000000000'
}

# Donde esta HOY el icono del telefono con lupa, mirando la ventana en vez de recordarlo.
#
# No sube Azul al frente ni lo toca: PrintWindow pinta la ventana aunque este tapada.
#
# Devuelve una tabla con Ok, y si Ok es falso, Por con el motivo. NUNCA devuelve una
# posicion "a lo mejor": si el dibujo no aparece bastante claro, o aparece dos veces, dice
# que no y el que llama debe negarse a picar. Un clic a ciegas en esa barra cae en los
# iconos vecinos, que abren otras cosas.
function Find-IconoLupa {
  param(
    [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
    # En diezmilesimas. Coincidencia total minima, tinta reconocida minima, y cuanto tiene
    # que quedarse atras el segundo mejor encaje para fiarse del primero.
    [int]$MinAcuerdo    = 9300,
    [int]$MinTinta      = 8000,
    [int]$MargenSegundo = 150,
    # Franja vertical donde vive la barra de iconos. Se busca en todo el ancho porque lo
    # que se mueve es justo eso, pero no en todo el alto: mas abajo empieza el contenido
    # del cliente y no hay por que darle ocasion de parecerse.
    [int]$Arriba = 60,
    [int]$Abajo  = 260
  )
  $hu = $script:HuellaLupa
  $bmp = [AzulShot]::Capture($Hwnd)
  if ($null -eq $bmp) { return @{ Ok = $false; Por = "la ventana de Azul no tiene tamano utilizable" } }
  try {
    $r = [AzulShot]::BuscaHuella($bmp, $hu.Mascara, $hu.Ancho, $hu.Alto,
                                 0, $Arriba, $bmp.Width, $Abajo, $hu.Umbral)
  } finally { $bmp.Dispose() }

  if ($r -notlike 'OK|*') { return @{ Ok = $false; Por = ($r -split '\|', 2)[1] } }
  $n = ($r -split '\|', 2)[1] -split ','
  $x = [int]$n[0]; $y = [int]$n[1]
  $acuerdo = [int]$n[2]; $tinta = [int]$n[3]; $segundo = [int]$n[4]
  $cx = $x + $hu.DX
  $cy = $y + $hu.DY
  $detalle = "encaje $acuerdo, tinta $tinta, segundo $segundo (de 10000)"

  if ($acuerdo -lt $MinAcuerdo) {
    return @{ Ok = $false; Por = "el icono no aparece claro en la ventana: $detalle"; X = $cx; Y = $cy }
  }
  if ($tinta -lt $MinTinta) {
    return @{ Ok = $false; Por = "el dibujo del icono no cuadra: $detalle"; X = $cx; Y = $cy }
  }
  if ($segundo -ge 0 -and ($acuerdo - $segundo) -lt $MargenSegundo) {
    return @{ Ok = $false; Por = "hay dos sitios que se parecen al icono, no se pica: $detalle"; X = $cx; Y = $cy }
  }
  return @{ Ok = $true; X = $cx; Y = $cy; Acuerdo = $acuerdo; Tinta = $tinta; Segundo = $segundo; Detalle = $detalle }
}
function Test-CapturasDistintas {
  param(
    [Parameter(Mandatory=$true)][string]$Antes,
    [Parameter(Mandatory=$true)][string]$Despues,
    [int]$Rejilla = 40,
    [double]$MinimoCambio = 0.02
  )
  # Ojo: Image::FromFile deja el archivo BLOQUEADO mientras el objeto viva, y el que llama
  # necesita poder sobrescribir estos PNG en la vuelta siguiente. Se copian los pixeles a un
  # Bitmap propio y se suelta el original enseguida.
  $ia = [System.Drawing.Image]::FromFile($Antes)
  $a  = New-Object System.Drawing.Bitmap $ia
  $ia.Dispose()
  try {
    $ib = [System.Drawing.Image]::FromFile($Despues)
    $b  = New-Object System.Drawing.Bitmap $ib
    $ib.Dispose()
    try {
      if ($a.Width -ne $b.Width -or $a.Height -ne $b.Height) { return $true }
      $dif = 0; $total = $Rejilla * $Rejilla
      for ($i = 0; $i -lt $Rejilla; $i++) {
        for ($j = 0; $j -lt $Rejilla; $j++) {
          $x = [int](($i + 0.5) * $a.Width  / $Rejilla)
          $y = [int](($j + 0.5) * $a.Height / $Rejilla)
          if ($x -ge $a.Width)  { $x = $a.Width  - 1 }
          if ($y -ge $a.Height) { $y = $a.Height - 1 }
          if ($a.GetPixel($x,$y).ToArgb() -ne $b.GetPixel($x,$y).ToArgb()) { $dif++ }
        }
      }
      return (($dif / $total) -ge $MinimoCambio)
    } finally { $b.Dispose() }
  } finally { $a.Dispose() }
}

# Espera a que la ventana DEJE DE CAMBIAR: captura cada tanto hasta que dos seguidas salen
# iguales. Es el equivalente visual de la guarda que ya espera a que el arbol de atributos
# deje de crecer. Hace falta porque el titulo del marco cambia a "Cuenta: <nombre>" antes de
# que terminen de pintarse los campos, y una captura en ese hueco sale a medias.
# Devuelve $true si se estabilizo, $false si se acabo el tiempo (el que llama decide).
#
# Las capturas NO tocan el disco. La version anterior escribia dos PNG por vuelta, los volvia
# a leer, los clonaba y los muestreaba con GetPixel: por cada comparacion salian dos
# codificaciones PNG, dos lecturas de archivo, un Copy-Item y 3200 GetPixel. Nada de eso
# aportaba: los bitmaps ya estaban en memoria recien salidos de PrintWindow. Ahora se quedan
# ahi y se comparan con LockBits.
#
# El numero de PrintWindow es el mismo que antes -- uno por vuelta, reutilizando el anterior
# como termino de comparacion -- asi que lo que se quita es rodeo, no muestreo.
function Wait-PantallaEstable {
  param(
    [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
    [int]$SegundosTope = 30,
    # 250 ms y 6 confirmaciones, no 700 y 2. La garantia es la MISMA -- lo que hay que exigir
    # es un rato seguido sin cambios, y 250x6 = 1500 ms es incluso mas que los 1400 de 700x2 --
    # pero se muestrea seis veces mas fino, asi que el momento en que Azul deja de pintar se
    # detecta antes en vez de esperar al siguiente multiplo de 700.
    #
    # Esto SOLO es asumible desde que comparar cuesta 0.24 ms en vez de 57 (ver Diferencia).
    # Con el rodeo por disco, sondear cada 250 ms habria puesto el coste de comparar por
    # encima del intervalo. Bajar el intervalo es el arreglo; quitar el disco es lo que lo
    # hace posible.
    [int]$Intervalo = 250,
    # Cuantas comparaciones iguales seguidas hacen falta. Con una sola bastaria que Azul
    # pausara el pintado un instante para dar por buena una pantalla a medias.
    [int]$Confirmaciones = 6,
    # Mismo umbral minimo de siempre: que un solo punto de la rejilla cambie ya cuenta como
    # "sigue pintando".
    [double]$MinimoCambio = 0.0006,
    [int]$Rejilla = 40
  )
  $a = [AzulShot]::Capture($Hwnd)
  if ($null -eq $a) { throw "La ventana de Azul no tiene tamano utilizable." }
  try {
    $tope = (Get-Date).AddSeconds($SegundosTope)
    $iguales = 0
    while ((Get-Date) -lt $tope) {
      Start-Sleep -Milliseconds $Intervalo
      $b = [AzulShot]::Capture($Hwnd)
      if ($null -eq $b) { continue }
      $cambio = [AzulShot]::Diferencia($a, $b, $Rejilla)
      # La nueva pasa a ser la referencia y la vieja se suelta enseguida: son bitmaps de
      # pantalla completa y dejarlos vivos por 30 s de espera se nota en memoria.
      $a.Dispose()
      $a = $b
      if ($cambio -ge $MinimoCambio) { $iguales = 0 }
      else {
        $iguales++
        if ($iguales -ge $Confirmaciones) { return $true }
      }
    }
    return $false
  } finally {
    if ($null -ne $a) { $a.Dispose() }
  }
}

# PrintWindow sobre Swing a veces devuelve una imagen en blanco o en negro SIN fallar.
# Se muestrea una rejilla y si casi todo es del mismo color, la captura se declara invalida.
# Una imagen en negro en el documento de un cliente es peor que no tener imagen.
function Test-CapturaValida {
  param(
    [Parameter(Mandatory=$true)][string]$Ruta,
    [int]$Rejilla = 20,
    [double]$Umbral = 0.99
  )
  if (-not (Test-Path $Ruta)) { return $false }
  $img = [System.Drawing.Image]::FromFile($Ruta)
  try {
    $bmp = New-Object System.Drawing.Bitmap $img
    try {
      $conteo = @{}
      for ($i = 0; $i -lt $Rejilla; $i++) {
        for ($j = 0; $j -lt $Rejilla; $j++) {
          $x = [int](($i + 0.5) * $bmp.Width  / $Rejilla)
          $y = [int](($j + 0.5) * $bmp.Height / $Rejilla)
          if ($x -ge $bmp.Width)  { $x = $bmp.Width  - 1 }
          if ($y -ge $bmp.Height) { $y = $bmp.Height - 1 }
          $k = $bmp.GetPixel($x, $y).ToArgb()
          if ($conteo.ContainsKey($k)) { $conteo[$k]++ } else { $conteo[$k] = 1 }
        }
      }
      $total = $Rejilla * $Rejilla
      $mayor = ($conteo.Values | Measure-Object -Maximum).Maximum
      return (($mayor / $total) -lt $Umbral)
    } finally { $bmp.Dispose() }
  } finally { $img.Dispose() }
}
