$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
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
public struct ATI { public long caption; public long summary; public int rowCount; public int columnCount; public long ctx; public long tbl; }
[StructLayout(LayoutKind.Sequential)]
public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int px; public int py; }
public static class JA {
  const string D="WindowsAccessBridge-32.dll"; const int NB=512,MAXA=256,MAXTD=32;
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern void Windows_run();
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextFromHWND(IntPtr h,out int vm,out long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextInfo(int vm,long ac,out ACI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern long getAccessibleChildFromContext(int vm,long ac,int idx);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleTableInfo(int vm,long ac,out ATI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleActions(int vm,long ac,IntPtr a);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool doAccessibleActions(int vm,long ac,IntPtr a,out int f);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern void addAccessibleSelectionFromContext(int vm,long ac,int i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern void clearAccessibleSelectionFromContext(int vm,long ac);
  [DllImport("user32.dll")] static extern bool PeekMessage(out MSG m,IntPtr h,uint a,uint b,uint c);
  [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
  [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG m);
  public static void Pump(int ms){ var e=DateTime.Now.AddMilliseconds(ms); MSG m;
    while(DateTime.Now<e){ while(PeekMessage(out m,IntPtr.Zero,0,0,1)){TranslateMessage(ref m);DispatchMessage(ref m);} System.Threading.Thread.Sleep(4);} }
  public static bool Click(int vm,long ac){
    int sz=4+MAXA*NB; IntPtr b=Marshal.AllocHGlobal(sz); string act=null;
    try{ for(int i=0;i<sz;i++) Marshal.WriteByte(b,i,0);
      if(!getAccessibleActions(vm,ac,b)) return false;
      if(Marshal.ReadInt32(b,0)<=0) return false;
      act=Marshal.PtrToStringUni(new IntPtr(b.ToInt64()+4));
    } finally{ Marshal.FreeHGlobal(b);}
    int s2=4+MAXTD*NB; IntPtr b2=Marshal.AllocHGlobal(s2);
    try{ for(int i=0;i<s2;i++) Marshal.WriteByte(b2,i,0);
      Marshal.WriteInt32(b2,0,1);
      byte[] nb=Encoding.Unicode.GetBytes(act);
      Marshal.Copy(nb,0,new IntPtr(b2.ToInt64()+4),Math.Min(nb.Length,NB-2));
      int f=-1; return doAccessibleActions(vm,ac,b2,out f);
    } finally{ Marshal.FreeHGlobal(b2);} }
}
'@
[JA]::Windows_run(); [JA]::Pump(1200)
$p=Get-Process -Name jp2launcher | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1
$vm=0;$root=[long]0
[JA]::getAccessibleContextFromHWND($p.MainWindowHandle,[ref]$vm,[ref]$root)|Out-Null
function I([long]$ac){$i=New-Object ACI; if([JA]::getAccessibleContextInfo($vm,$ac,[ref]$i)){return $i}; return $null}
function Cel([long]$t,[int]$idx){ $c=[JA]::getAccessibleChildFromContext($vm,$t,$idx); if($c -eq 0){return ""}; $i=I $c; if($null -eq $i){return ""}; if($i.name -eq $null){return ""}; return $i.name.Trim() }
function Walk([long]$s,[scriptblock]$pred){
  $res=@(); $st=New-Object System.Collections.Stack; $st.Push($s); $n=0
  while($st.Count -gt 0 -and $n -lt 40000){
    $ac=$st.Pop(); $n++; $i=I $ac; if($null -eq $i){continue}
    if(& $pred $i){ $res+=$ac }
    if($i.role_en_US -eq "table"){continue}
    for($k=$i.childrenCount-1;$k -ge 0;$k--){ $c=[JA]::getAccessibleChildFromContext($vm,$ac,$k); if($c -ne 0){$st.Push($c)} }
  }
  return $res
}

# cerrar detalle si quedo abierto
$d = Walk $root { param($i) $i.role_en_US -eq "internal frame" -and $i.name -ne $null -and $i.name.StartsWith("Detalles del Producto Asignado") }
if($d.Count -gt 0){
  $cb = Walk $d[0] { param($i) $i.role_en_US -eq "push button" -and $i.name -ne $null -and $i.name.Trim() -eq "Cerrar" }
  if($cb.Count -gt 0){ [JA]::Click($vm,$cb[0])|Out-Null; [JA]::Pump(900) }
}

$sub=0;$SR=0;$SC=0
foreach($t in (Walk $root { param($i) $i.role_en_US -eq "table" })){
  $ti=New-Object ATI
  if(-not [JA]::getAccessibleTableInfo($vm,$t,[ref]$ti)){continue}
  if($ti.columnCount -lt 10 -or $ti.rowCount -lt 1){continue}
  $c0=[JA]::getAccessibleChildFromContext($vm,$t,0); if($c0 -eq 0){continue}
  $i0=I $c0; if($null -ne $i0 -and $i0.role_en_US -eq "radio button"){ $sub=$t;$SR=$ti.rowCount;$SC=$ti.columnCount; break }
}
"tabla $SR x $SC"
$btnL = Walk $root { param($i) $i.role_en_US -eq "push button" -and $i.name -ne $null -and $i.name.Trim() -eq "Ver Productos Asignados" }
$btn=$btnL[0]

[JA]::clearAccessibleSelectionFromContext($vm,$sub); [JA]::Pump(250)
[JA]::addAccessibleSelectionFromContext($vm,$sub,0); [JA]::Pump(700)
"linea: $(Cel $sub 1)"
[JA]::Click($vm,$btn)|Out-Null

$det=0
for($w=0;$w -lt 30 -and $det -eq 0;$w++){
  [JA]::Pump(300)
  $dd = Walk $root { param($i) $i.role_en_US -eq "internal frame" -and $i.name -ne $null -and $i.name.StartsWith("Detalles del Producto Asignado") }
  if($dd.Count -gt 0){ $det=$dd[0] }
}
if($det -eq 0){ "no abrio el detalle"; exit 1 }
[JA]::Pump(2000)

$at=0;$AR=0;$AC=0;$bk=0
foreach($t in (Walk $det { param($i) $i.role_en_US -eq "table" })){
  $ti=New-Object ATI
  if([JA]::getAccessibleTableInfo($vm,$t,[ref]$ti)){
    $k=$ti.rowCount*$ti.columnCount
    if($k -gt $bk){ $bk=$k;$at=$t;$AR=$ti.rowCount;$AC=$ti.columnCount }
  }
}
"atributos: $AR x $AC"
""
"fila | col0 (nombre) | col2 (valor) | col3 (descripcion)"
for($r=0;$r -lt $AR;$r++){
  $c0=Cel $at ($r*$AC)
  $c2=if($AC -ge 3){Cel $at ($r*$AC+2)}else{""}
  $c3=if($AC -ge 4){Cel $at ($r*$AC+3)}else{""}
  if($c0 -eq "" -and $c2 -eq "" -and $c3 -eq ""){continue}
  "{0,3} | {1,-34} | {2,-28} | {3}" -f $r,$c0,$c2,$c3
}
