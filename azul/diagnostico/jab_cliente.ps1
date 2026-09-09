param(
  # -Todo vuelca TODO lo que se ve en la ventana, no solo la franja de arriba.
  # Sirve cuando lo que se busca no aparece donde se esperaba.
  [switch]$Todo,
  # Texto a buscar en name/description de CUALQUIER nodo, sin filtrar por geometria.
  # Es la prueba decisiva de si algo existe para el puente: si el nombre del cliente no
  # sale aqui, ese widget no esta expuesto y no se puede pulsar sin mouse.
  [string]$Buscar = "Cliente"
)
$ErrorActionPreference='Stop'
# Diagnostico de SOLO LECTURA. No pulsa nada, no mueve el mouse, no toca Azul.
#
# Busca el nodo "Cliente:" de la parte de arriba de Azul y todo lo que se le parezca
# (pestanas, etiquetas, enlaces y botones de la franja superior), para responder tres cosas
# antes de escribir codigo que pulse:
#
#   1. que widget es "Cliente:" (etiqueta, enlace, boton)
#   2. si se puede pulsar POR EL PUENTE (columna act=SI) o solo con un clic real del mouse
#   3. que nodo hay que pulsar para volver a la pestana Suscripciones
#
# Se corre DOS veces y se comparan las salidas:
#   1) con Azul en Suscripciones y el panel del cliente cerrado
#   2) con el panel del cliente abierto A MANO
#
# Uso (PowerShell de 32 bits, igual que todo lo demas):
#   & "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "...\diagnostico\jab_cliente.ps1"
#
# NOTA: mantener este archivo en ASCII puro. PowerShell 5.1 lee los .ps1 sin BOM como ANSI
# y rompe los acentos en silencio.

if (-not ("JC" -as [type])) {
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct ACI {
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=1024)] public string name;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=1024)] public string description;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string role;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string role_en_US;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string states;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string states_en_US;
  public int indexInParent; public int childrenCount;
  public int x; public int y; public int width; public int height;
  public int accessibleComponent; public int accessibleAction;
  public int accessibleSelection; public int accessibleText; public int accessibleInterfaces; }
[StructLayout(LayoutKind.Sequential)]
public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int px; public int py; }
public class JC {
  const string D="WindowsAccessBridge-32.dll";
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern void Windows_run();
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextFromHWND(IntPtr h,out int vm,out long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextInfo(int vm,long ac,out ACI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern long getAccessibleChildFromContext(int vm,long ac,int idx);
  [DllImport("user32.dll")] static extern bool PeekMessage(out MSG m,IntPtr h,uint a,uint b,uint c);
  [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
  [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG m);
  public static void Pump(int ms){ var e=DateTime.Now.AddMilliseconds(ms); MSG m;
    while(DateTime.Now<e){ while(PeekMessage(out m,IntPtr.Zero,0,0,1)){TranslateMessage(ref m);DispatchMessage(ref m);} System.Threading.Thread.Sleep(5);} }
}
'@
}

[JC]::Windows_run(); [JC]::Pump(1500)

$p = Get-Process -Name jp2launcher -ErrorAction SilentlyContinue |
     Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
if (-not $p) { "ERROR: Azul no esta corriendo."; exit 1 }

# Minimizada, Windows aparca la ventana en -32000,-32000: las coordenadas que da el puente
# dejan de servir y una captura saldria en basura. Se avisa aqui para no leer mal la salida.
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class Ico2{[DllImport("user32.dll")]public static extern bool IsIconic(IntPtr h);}'
if ([Ico2]::IsIconic($p.MainWindowHandle)) {
  "ERROR: Azul esta minimizado. Restauralo (puede quedar detras de otras ventanas, pero no"
  "       iconizado) y vuelve a correr. Minimizada, las coordenadas no sirven."
  exit 1
}

$vm = 0; $root = [long]0
if (-not [JC]::getAccessibleContextFromHWND($p.MainWindowHandle, [ref]$vm, [ref]$root)) {
  "ERROR: sin contexto raiz. El puente no ve a Azul."; exit 1
}

function I([long]$ac) {
  $i = New-Object ACI
  if ([JC]::getAccessibleContextInfo($vm, $ac, [ref]$i)) { return $i }
  return $null
}

# En el arbol el texto a veces no viene en name sino en description, envuelto en HTML.
function Strip([string]$s) {
  if ($null -eq $s) { return "" }
  $sb = New-Object System.Text.StringBuilder
  $tag = $false
  foreach ($c in $s.ToCharArray()) {
    if ($c -eq '<') { $tag = $true }
    elseif ($c -eq '>') { $tag = $false }
    elseif (-not $tag) { [void]$sb.Append($c) }
  }
  return $sb.ToString().Replace("&nbsp;", " ").Trim()
}

$ri = I $root
if ($null -eq $ri) { "ERROR: no se pudo leer la raiz."; exit 1 }
$topY = $ri.y
# Hasta abajo de las pestanas: "Cliente:" vive en algun lado de esta franja. NO se filtra
# por rol: la primera version filtraba a etiquetas y botones y no encontro nada, asi que
# el nodo es de otro tipo. Aqui sale todo lo que se vea, y ya lo miramos.
$franja = if ($Todo) { $topY + 100000 } else { $topY + 345 }

"Azul: ventana en x=$($ri.x) y=$($ri.y) $($ri.width)x$($ri.height)"
"Franja superior considerada: y < $franja"
""

$conCliente = @()   # cualquier nodo que mencione al cliente
$franjaSup  = @()   # widgets pulsables de la franja de arriba
$pestanas   = @()   # todas las pestanas, esten donde esten (para el camino de regreso)
$marcos     = @()   # marcos y dialogos: delatan si algo se abrio aparte

$st = New-Object System.Collections.Stack
$st.Push($root)
$n = 0

while ($st.Count -gt 0 -and $n -lt 40000) {
  $ac = $st.Pop(); $n++
  $i = I $ac
  if ($null -eq $i) { continue }

  $r    = $i.role_en_US
  $nm   = if ($null -eq $i.name) { "" } else { $i.name.Trim() }
  $desc = Strip $i.description
  $act  = if ($i.accessibleAction -ne 0) { "SI " } else { "no " }
  $linea = "  [{0,-14}] act={1} y={2,-5} x={3,-5} {4}x{5}  name='{6}'  desc='{7}'  [{8}]" -f `
           $r, $act, $i.y, $i.x, $i.width, $i.height, $nm, $desc, $i.states_en_US

  if ($nm -like "*$Buscar*" -or $desc -like "*$Buscar*") { $conCliente += $linea }

  if ($r -eq "page tab") { $pestanas += $linea }

  if ($r -eq "internal frame" -or $r -eq "dialog" -or $r -eq "option pane") {
    $marcos += "  [$r] '$nm' [$($i.states_en_US)]"
  }

  # Solo lo que de verdad se ve: los nodos sin colocar reportan y=-1 y son cientos.
  if ($i.y -ge 0 -and $i.y -lt $franja -and $i.states_en_US -match 'showing') {
    $franjaSup += $linea
  }

  # no descender en tablas: son miles de celdas y aqui no aportan nada
  if ($r -eq "table") { continue }
  for ($k = $i.childrenCount - 1; $k -ge 0; $k--) {
    $c = [JC]::getAccessibleChildFromContext($vm, $ac, $k)
    if ($c -ne 0) { $st.Push($c) }
  }
}

"nodos recorridos: $n"
""
"=== NODOS QUE MENCIONAN '$Buscar' (todo el arbol, sin filtrar por geometria) ==="
if ($conCliente.Count) { $conCliente } else { "  (ninguno)" }
""
"=== PESTANAS (camino de regreso a Suscripciones) ==="
if ($pestanas.Count) { $pestanas } else { "  (ninguna)" }
""
"=== FRANJA SUPERIOR (y < $franja) ==="
if ($franjaSup.Count) { $franjaSup } else { "  (ninguno)" }
""
"=== MARCOS / DIALOGOS ABIERTOS ==="
if ($marcos.Count) { $marcos } else { "  (ninguno)" }
""
"act=SI  -> se puede pulsar por el puente, sin mouse y sin traer Azul al frente."
"act=no  -> solo con un clic real, que obliga a traer la ventana al frente."
