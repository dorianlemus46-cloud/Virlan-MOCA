$ErrorActionPreference='Stop'
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
public struct ATI { public long caption; public long summary; public int rowCount; public int columnCount; public long ctx; public long tbl; }
[StructLayout(LayoutKind.Sequential)]
public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int px; public int py; }
public class J {
  const string D="WindowsAccessBridge-32.dll";
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern void Windows_run();
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextFromHWND(IntPtr h,out int vm,out long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleContextInfo(int vm,long ac,out ACI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern long getAccessibleChildFromContext(int vm,long ac,int idx);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] public static extern bool getAccessibleTableInfo(int vm,long ac,out ATI i);
  [DllImport("user32.dll")] static extern bool PeekMessage(out MSG m,IntPtr h,uint a,uint b,uint c);
  [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
  [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG m);
  public static void Pump(int ms){ var e=DateTime.Now.AddMilliseconds(ms); MSG m;
    while(DateTime.Now<e){ while(PeekMessage(out m,IntPtr.Zero,0,0,1)){TranslateMessage(ref m);DispatchMessage(ref m);} System.Threading.Thread.Sleep(5);} }
}
'@
[J]::Windows_run(); [J]::Pump(1500)
$p=Get-Process -Name jp2launcher | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1
$vm=0;$root=[long]0
[J]::getAccessibleContextFromHWND($p.MainWindowHandle,[ref]$vm,[ref]$root)|Out-Null
function I([long]$ac){$i=New-Object ACI; if([J]::getAccessibleContextInfo($vm,$ac,[ref]$i)){return $i}; return $null}
$st=New-Object System.Collections.Stack; $st.Push($root); $n=0
$frames=@(); $tabs=@(); $btns=@()
while($st.Count -gt 0 -and $n -lt 40000){
  $ac=$st.Pop(); $n++
  $i=I $ac; if($null -eq $i){continue}
  $r=$i.role_en_US
  if($r -eq "internal frame" -or $r -eq "dialog" -or $r -eq "option pane"){ $frames += "[$r] '$($i.name)' [$($i.states_en_US)]" }
  if($r -eq "table"){
    $t=New-Object ATI
    if([J]::getAccessibleTableInfo($vm,$ac,[ref]$t)){
      $c0=[J]::getAccessibleChildFromContext($vm,$ac,0); $r0=""
      if($c0 -ne 0){ $i0=I $c0; if($null -ne $i0){$r0=$i0.role_en_US} }
      $tabs += "  $($t.rowCount)x$($t.columnCount)  celda0=$r0"
    }
    continue
  }
  if($r -eq "push button" -and ($i.name.Trim() -eq "Ver Productos Asignados" -or $i.name.Trim() -eq "Cerrar")){
    $btns += "  '$($i.name.Trim())' [$($i.states_en_US)]"
  }
  for($k=$i.childrenCount-1;$k -ge 0;$k--){ $c=[J]::getAccessibleChildFromContext($vm,$ac,$k); if($c -ne 0){$st.Push($c)} }
}
"nodos: $n"
"=== MARCOS / DIALOGOS ==="; if($frames.Count){$frames}else{"  (ninguno)"}
"=== TABLAS ==="; if($tabs.Count){$tabs}else{"  (ninguna)"}
"=== BOTONES CLAVE ==="; if($btns.Count){$btns}else{"  (ninguno)"}
