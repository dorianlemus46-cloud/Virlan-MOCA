param(
  # Numero de cuenta a poner en "ID de CF". Si se omite, se toma de la lista de ordenes.
  [string]$Cuenta = "",
  # Razon social del Excel. Con -Fila se saca sola; se acepta como parametro para que lote.ps1
  # pueda pasarla y siga funcionando la comprobacion de que la fila elegida es del cliente
  # pedido, que es la guarda que evita entrar al cliente equivocado.
  [string]$Razon = "",
  # Fila de la lista de ordenes, tal como la numera la hoja: la 1 son los encabezados, asi que
  # -Fila 6 es la sexta fila que se ve. Sale del CSV, no del Excel abierto.
  [int]$Fila = 0,
  # La lista de ordenes en CSV. Vacio = lista_ordenes.csv junto al script.
  [string]$Lista = "",
  # Donde esta el icono del telefono con lupa, relativo a la esquina de la ventana de Azul.
  # En 0 -- lo normal -- NO se usa una posicion: se busca el icono en la ventana y se pica
  # donde este de verdad. Ver Find-IconoLupa en captura.ps1.
  #
  # Hasta el 18/09/2026 aqui habia 791,121, medido a mano el 03/09/2026 con zoom.ps1 sobre
  # la ventana en 1366x768, porque ese icono es contenido de un navegador incrustado y el
  # puente de accesibilidad no lo ve. Funciono 220 clientes y despues fallo: la barra de
  # Azul se corre a la derecha cuando el panel de la izquierda cambia de ancho, y el clic
  # empezo a caer 55 pixeles antes del icono, en un campo vacio. Una posicion aprendida no
  # sirve para algo que se mueve.
  #
  # Se dejan como parametro para poder forzar la posicion a mano si algun dia el dibujo del
  # icono cambia y dejara de reconocerse. Con los dos mayores que 0 se pica ahi sin mirar.
  [int]$IconoX = 0,
  [int]$IconoY = 0,
  # No hacer clic: exige que el formulario ya este abierto. Para correr sin tocar el mouse.
  [switch]$SinClic,
  # Levantar Azul si esta minimizado, en vez de negarse. Apagado por defecto: la guarda de
  # Get-AzulHwnd existe porque minimizada Windows aparca la ventana en -32000 y tanto las
  # coordenadas como la captura salen basura. Esto no relaja la guarda, la resuelve.
  [switch]$Restaurar,
  # Llenar el campo, fotografiar y PARAR, sin pulsar "Buscar Ahora". Sirve para ver con los
  # ojos que el numero llego al campo, en vez de deducirlo de que la busqueda no devolvio nada.
  [switch]$SoloLlenar,
  # Pararse en la rejilla de resultados sin elegir cuenta. Por defecto SI se elige la primera
  # fila y se entra al cliente, que es lo que pidio Dorian el 03/09/2026.
  [switch]$NoEntrar,
  # Saltarse la busqueda y seguir desde la rejilla de resultados que ya este abierta.
  [switch]$DesdeResultados,
  # Buscar aunque haya criterios viejos en otros campos del formulario.
  [switch]$Forzar,
  # Pausas ADICIONALES para no pedirle datos a Azul mas rapido de lo que aguanta una persona
  # (Dorian, 07/09/2026). En segundos.
  #
  # -PausaEscribir sigue siendo una espera clavada: se para esos segundos entre escribir el
  # numero y pulsar "Buscar Ahora".
  #
  # -PausaResultado ya NO se gasta esperando. Desde el 18/09/2026 es cuanto se ALARGA el
  # plazo que se le da a Azul para traer el resultado, y se sale en cuanto lo trae. Antes
  # eran diez segundos parado antes de mirar siquiera, y era donde mas tiempo se iba.
  [int]$PausaEscribir = 5,   # entre escribir el numero de cuenta y pulsar "Buscar Ahora"
  [int]$PausaResultado = 10  # segundos de mas en el plazo para reconocer el resultado
)
$ErrorActionPreference = 'Stop'
# NOTA: este archivo debe permanecer en ASCII puro, igual que azul_fast.ps1.
# PowerShell 5.1 lee los .ps1 sin BOM como ANSI y rompe los acentos en silencio.
#
# Busca una cuenta en Azul por "ID de CF" y SE DETIENE en cuanto Azul devuelve el resultado.
# No selecciona el cliente, no lee suscripciones, no toca Word: eso son pasos que se
# agregaran despues.
#
# El unico momento en que se toca la maquina es el clic en el icono del telefono con lupa:
# ese icono es Chromium y el puente no lo ve. Todo lo demas -- llenar el campo, pulsar
# "Buscar Ahora", saber que llego el resultado -- va por el Java Access Bridge.
#
# El mapa completo (marcos, campos, senales de fin) esta en docs/mapa-datos.md, seccion 2,
# "El formulario de busqueda".
#
#   & "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
#     -File ".\azul\buscar_cuenta.ps1" -Fila 6      (desde la raiz del proyecto)

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
public struct ACI2 {
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=1024)] public string name;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=1024)] public string description;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string role;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string role_en_US;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string states;
  [MarshalAs(UnmanagedType.ByValTStr,SizeConst=256)] public string states_en_US;
  public int indexInParent; public int childrenCount;
  public int x; public int y; public int width; public int height;
  public int accessibleComponent; public int accessibleAction;
  public int accessibleSelection; public int accessibleText; public int accessibleInterfaces;
}
[StructLayout(LayoutKind.Sequential)]
public struct ATI2 { public long caption; public long summary; public int rowCount; public int columnCount; public long ctx; public long tbl; }
[StructLayout(LayoutKind.Sequential)]
public struct ATXI { public int charCount; public int caretIndex; public int indexAtPoint; }
[StructLayout(LayoutKind.Sequential)]
public struct MSG2 { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int px; public int py; }

public static class Busca {
  const string D="WindowsAccessBridge-32.dll";
  const int NB=512, MAXA=256, MAXTD=32;

  // Nombres de Azul con acentos, en escapes para sobrevivir la codificacion ASCII
  const string F_CIM = "Encontrar Comunicante B\u00FAsqueda CIM";
  // La rejilla de resultados no siempre es el mismo marco: con una coincidencia unica Azul
  // carga al cliente y no abre ninguna; con varias abre "Identificar Cliente" o
  // "Busqueda: Contacto y Cuenta Financiera", segun por donde entre. En vez de ir apuntando
  // nombres, se reconoce por lo que TIENE: un boton contador "N Registro/s". Eso es lo que
  // hace a un applet de resultados de Siebel, se llame como se llame.
  const string CAMPO = "ID de CF:";
  // Cuando la busqueda ACIERTA, Azul carga al cliente y abre este marco:
  // "Inicio de Interaccion [2] - 682000001". Es la unica senal POSITIVA que hay, y hace falta:
  // que el formulario se cierre pasa igual cuando no encuentra nada, asi que sin esto una
  // busqueda fallida se reporta como exito. Costo dos corridas creerselo.
  const string F_INT = "Inicio de Interacci\u00F3n";   // "Interaccion" con o acentuada
  const string B_BUSCAR = "Buscar Ahora";
  const string B_SELEC  = "Seleccionar";
  const string B_CONT   = "Registro";

  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void Windows_run();
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleContextFromHWND(IntPtr h,out int vm,out long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleContextInfo(int vm,long ac,out ACI2 i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern long getAccessibleChildFromContext(int vm,long ac,int idx);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleTableInfo(int vm,long ac,out ATI2 i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleActions(int vm,long ac,IntPtr a);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool doAccessibleActions(int vm,long ac,IntPtr a,out int f);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool requestFocus(int vm,long ac);
  // Seleccion por el modelo real de la tabla. El indice es de CELDA, no de fila: para la
  // fila r hay que pasar r*columnas. Es la misma via que usa azul_fast.ps1 con la tabla de
  // Suscripciones, donde marcar el radio por accion de accesibilidad NO habilitaba el boton
  // y esto si.
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void addAccessibleSelectionFromContext(int vm,long ac,int i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void clearAccessibleSelectionFromContext(int vm,long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleTextInfo(int vm,long ac,out ATXI i,int x,int y);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleTextRange(int vm,long ac,int start,int end,IntPtr t,short len);
  [DllImport("user32.dll")] static extern bool PeekMessage(out MSG2 m,IntPtr h,uint a,uint b,uint c);
  [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG2 m);
  [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG2 m);

  static int vm; static long root;
  static List<long> panes = new List<long>();

  public static string PROG="";
  static void P(string s){
    try{ System.IO.File.AppendAllText(PROG, DateTime.Now.ToString("HH:mm:ss")+"  "+s+"\r\n"); }catch{}
  }
  public static void Nota(string s){ P(s); }

  // Esperar bombeando mensajes, no con Start-Sleep: el puente necesita que el hilo despache.
  static void Pump(int ms){
    var e=DateTime.Now.AddMilliseconds(ms); MSG2 m;
    while(DateTime.Now<e){
      while(PeekMessage(out m,IntPtr.Zero,0,0,1)){ TranslateMessage(ref m); DispatchMessage(ref m); }
      System.Threading.Thread.Sleep(3);
    }
  }
  public static void Espera(int ms){ Pump(ms); }

  // Contador de llamadas al puente. Inf y Kid son los dos puntos por los que pasa TODO
  // recorrido del arbol, asi que contarlos aqui mide el costo real sin instrumentar nada mas.
  // Es un incremento de un long: no cambia el comportamiento y no pesa de forma medible.
  // Sirve para comparar una corrida contra otra sin cronometro: el reloj depende de lo
  // cargado que este Azul, el numero de llamadas no.
  static long jab=0;
  public static long Llamadas(){ return jab; }
  static bool Inf(long ac,out ACI2 i){ jab++; return getAccessibleContextInfo(vm,ac,out i); }
  static long Kid(long ac,int k){ jab++; return getAccessibleChildFromContext(vm,ac,k); }
  static bool Est(ACI2 i,string s){ return i.states_en_US!=null && i.states_en_US.IndexOf(s)>=0; }

  static List<long> Find(long start,Func<ACI2,long,bool> pred,int cap,bool firstOnly){
    var res=new List<long>(); var st=new Stack<long>(); st.Push(start); int n=0; ACI2 i;
    while(st.Count>0 && n<cap){
      long ac=st.Pop(); n++;
      if(!Inf(ac,out i)) continue;
      if(pred(i,ac)){ res.Add(ac); if(firstOnly) return res; }
      // no descender en tablas: son miles de celdas. La tabla en si ya la vio el predicado.
      if(i.role_en_US=="table") continue;
      for(int k=i.childrenCount-1;k>=0;k--){ long c=Kid(ac,k); if(c!=0) st.Push(c); }
    }
    return res;
  }
  static long First(long s,Func<ACI2,long,bool> p){ var r=Find(s,p,60000,true); return r.Count>0?r[0]:0; }

  static bool Click(long ac){
    int sz=4+MAXA*NB; IntPtr b=Marshal.AllocHGlobal(sz); string act=null;
    try{
      for(int i=0;i<sz;i++) Marshal.WriteByte(b,i,0);
      if(!getAccessibleActions(vm,ac,b)) return false;
      if(Marshal.ReadInt32(b,0)<=0) return false;
      act=Marshal.PtrToStringUni(new IntPtr(b.ToInt64()+4));
    } finally { Marshal.FreeHGlobal(b); }
    int s2=4+MAXTD*NB; IntPtr b2=Marshal.AllocHGlobal(s2);
    try{
      for(int i=0;i<s2;i++) Marshal.WriteByte(b2,i,0);
      Marshal.WriteInt32(b2,0,1);
      byte[] nb=Encoding.Unicode.GetBytes(act);
      Marshal.Copy(nb,0,new IntPtr(b2.ToInt64()+4),Math.Min(nb.Length,NB-2));
      int f=-1; return doAccessibleActions(vm,ac,b2,out f);
    } finally { Marshal.FreeHGlobal(b2); }
  }

  // Contenido real de un campo de texto. Sin esto no hay forma de comprobar que lo que se
  // escribio es lo que quedo: setTextContents puede devolver true y no dejar nada.
  static string Texto(long ac){
    ATXI ti;
    if(!getAccessibleTextInfo(vm,ac,out ti,-1,-1)) return null;
    if(ti.charCount<=0) return "";
    int n=Math.Min(ti.charCount,500);
    IntPtr b=Marshal.AllocHGlobal(4096);
    try{
      for(int i=0;i<4096;i++) Marshal.WriteByte(b,i,0);
      if(!getAccessibleTextRange(vm,ac,0,n-1,b,(short)(n+1))) return null;
      string s=Marshal.PtrToStringUni(b);
      return s==null?"":s;
    } finally { Marshal.FreeHGlobal(b); }
  }

  // Los marcos internos son hijos DIRECTOS de un desktop pane; con los panes cacheados cada
  // sondeo cuesta ~10 llamadas JAB en vez de recorrer el arbol entero.
  // Por PREFIJO, no por igualdad. Azul numera las instancias: la segunda vez que se abre el
  // formulario se llama "Encontrar Comunicante Busqueda CIM [2]", y con comparacion exacta el
  // marco se encuentra al abrirlo y deja de encontrarse en cuanto Azul le pone el numero.
  // Es la misma razon por la que azul_fast.ps1 busca "Detalles del Producto Asignado" y
  // "Cuenta: " con StartsWith.
  // Solo por los panes cacheados. Cuesta ~10 llamadas JAB tanto si encuentra como si no:
  // esta es la version que va DENTRO de un bucle de sondeo.
  static long MarcoPanes(string prefijo){
    for(int p=0;p<panes.Count;p++){
      ACI2 pi; if(!Inf(panes[p],out pi)) continue;
      for(int k=0;k<pi.childrenCount;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI2 ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.Trim().StartsWith(prefijo)) return c;
      }
    }
    return 0;
  }
  // Recorrido completo del arbol: red de seguridad por si el marco no cuelga de ninguno de
  // los desktop panes cacheados. NUNCA dentro de un bucle de espera -- es la regla que ya
  // recoge docs/trampas.md, y el mismo reparto que hacen DetFrame/DetFrameFull en
  // azul_fast.ps1. Ojo con el reparto de costos: el caso caro no es encontrar, es NO
  // encontrar, porque entonces se visita el arbol entero antes de poder decir que no esta.
  static long MarcoFull(string prefijo){
    return First(root,(i,ac)=> i.role_en_US=="internal frame" && i.name!=null && i.name.Trim().StartsWith(prefijo));
  }
  static long Marco(string prefijo){
    long m=MarcoPanes(prefijo);
    return (m!=0) ? m : MarcoFull(prefijo);
  }

  public static bool Init(IntPtr hwnd){
    Windows_run();
    // Sondeo por condicion, no espera por reloj: se pregunta por lo que de verdad hace falta,
    // que es que el puente conteste con el contexto raiz de la ventana.
    //
    // 5 s y no 800 ms (05/09/2026), el mismo cambio y por el mismo motivo que en azul_fast.ps1.
    // El tope es techo, no coste: el bucle sale en cuanto el puente contesta. Se sube por el
    // Azul DEGRADADO -- el 04/09/2026 a las 13:24, con Azul vivo pero quemando un nucleo al
    // 100%, una corrida murio con "sin contexto raiz. El puente no ve a Azul". Con 800 ms, un
    // puente lento no es un puente ausente y se trataba igual.
    DateTime tope=DateTime.Now.AddMilliseconds(5000);
    while(!getAccessibleContextFromHWND(hwnd,out vm,out root) || root==0){
      if(DateTime.Now>=tope) return false;
      Pump(50);
    }
    panes = Find(root,(i,ac)=> i.role_en_US=="desktop pane",60000,false);
    // Justo despues de que el contexto contesta, el arbol puede venir a medias.
    //
    // PLAZO PROPIO, no el que sobre del anterior. Antes los dos bucles compartian 'tope': si
    // la raiz tardaba 700 de los 800 ms, a los panes les quedaban 100. Reintentar el cacheo
    // dentro de las sobras del presupuesto anterior es apostar a que el primer paso fue
    // rapido, justo cuando lo que se esta atendiendo es que no lo fue.
    DateTime topePanes=DateTime.Now.AddMilliseconds(2000);
    while(panes.Count==0 && DateTime.Now<topePanes){
      Pump(50);
      panes = Find(root,(i,ac)=> i.role_en_US=="desktop pane",60000,false);
    }
    P("panes cacheados: "+panes.Count);
    // Un cache VACIO no falla, y ese es el problema: Init devuelve true y la corrida sigue,
    // pero MarcoPanes() se queda sin donde mirar y cada sondeo cae en MarcoFull(), que visita
    // el arbol entero. Aqui pesa mas que en el lector: EsperarCIM sondea hasta 40 s y
    // EsperarResultado hasta 90. No se aborta, pero deja de ser silencioso.
    if(panes.Count==0) P("AVISO: cache de panes VACIO. Cada sondeo de marco va a recorrer el arbol entero.");
    return true;
  }

  public static bool HayCIM(){ return Marco(F_CIM)!=0; }

  public static bool EsperarCIM(int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      // Solo la via de panes. Con Marco() cada vuelta que no encontraba costaba un recorrido
      // completo del arbol, y son hasta ~100 vueltas en los 40 s: el sondeo se pasaba el
      // tiempo recorriendo Azul entero para decir "todavia no".
      if(MarcoPanes(F_CIM)!=0) return true;
      Pump(400);
    }
    // Una sola comprobacion completa antes de rendirse, que es la que ya habia aqui: si el
    // marco existe pero no cuelga de un pane cacheado, se ve en esta y no se pierde.
    return Marco(F_CIM)!=0;
  }

  // El campo, NO la etiqueta. Hay dos nodos con name 'ID de CF:': un label (act=no) y el
  // campo. Se exige rol 'text' y estado 'editable' o se escribe sobre la etiqueta y no pasa
  // nada, en silencio.
  static long CampoCF(long cim){
    return First(cim,(i,ac)=> i.role_en_US=="text" && Est(i,"editable")
                           && i.name!=null && i.name.Trim()==CAMPO);
  }

  // Criterios viejos en otros campos: la busqueda los combina, asi que un resto de una
  // consulta anterior devolveria otro cliente. Se avisa antes de pulsar nada.
  public static string CamposConTexto(){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    var sb=new StringBuilder();
    foreach(long t in Find(cim,(i,ac)=> i.role_en_US=="text" && Est(i,"editable"),60000,false)){
      ACI2 i; if(!Inf(t,out i)) continue;
      string nom=(i.name==null?"":i.name.Trim());
      if(nom==CAMPO) continue;
      string v=Texto(t);
      // OJO: se mira v.Length, NO v.Trim().Length. Un campo con SOLO ESPACIOS sigue siendo
      // un criterio para Azul y la busqueda no devuelve nada -- que se ve exactamente igual
      // que "esa cuenta no existe". La primera version trimeaba aqui y por eso un resto
      // invisible habria pasado desapercibido. Dato de Dorian, 03/09/2026.
      if(v!=null && v.Length>0){
        string muestra = (v.Trim().Length>0) ? ("'"+v+"'") : ("SOLO ESPACIOS ("+v.Length+" caracteres)");
        sb.AppendLine("  " + (nom.Length>0?nom:("campo en y="+i.y)) + "  ->  " + muestra);
      }
    }
    return sb.ToString();
  }

  // ---- llenado del campo, en tres tiempos --------------------------------
  // Primero se intenta por el puente, que no toca la maquina. Si el puente no puede (es lo
  // que pasa con este campo: setTextContents devuelve exito y lo deja vacio), PowerShell
  // teclea de verdad. Se separan los pasos porque el teclado vive en captura.ps1, que es
  // donde vive todo lo que toca la maquina, y el foco vive aqui, que es el puente.

  // setTextContents ya no se usa. Devolvia exito y a veces dejaba el campo vacio; y las dos
  // corridas que buscaron con el valor puesto por el puente devolvieron cero, mientras que la
  // que lo llevaba tecleado si encontro al cliente. Teclear cuesta ~250 ms y quita la duda.

  // Enfoca el campo y COMPRUEBA que quedo enfocado. Sin esta comprobacion, teclear seria
  // disparar a ciegas: las pulsaciones van a donde este el foco del sistema, y si el campo
  // no lo tiene acaban en otro sitio.
  public static string EnfocarCF(){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    long c=CampoCF(cim);
    if(c==0) return "ERROR: no encuentro el campo '"+CAMPO+"' (text, editable) en el formulario";
    for(int t=0;t<5;t++){
      requestFocus(vm,c); Pump(250);
      ACI2 i;
      if(Inf(c,out i) && Est(i,"focused")){
        string v=Texto(c); if(v==null) v="";
        return "OK|"+v.Length;
      }
    }
    return "ERROR: no consigo el foco en '"+CAMPO+"'. No se teclea a ciegas.";
  }

  // Enfoca el PRIMER campo con restos, distinto de ID de CF, y dice cuanto hay que borrar.
  // Devuelve VACIO cuando ya no queda ninguno. Se hace de uno en uno porque borrar es teclear,
  // y teclear vive en PowerShell: asi el bucle queda arriba y aqui solo el puente.
  public static string EnfocarSucio(){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    long campo=CampoCF(cim);
    foreach(long t in Find(cim,(i,ac)=> i.role_en_US=="text" && Est(i,"editable"),60000,false)){
      if(t==campo) continue;
      string v=Texto(t); if(v==null || v.Length==0) continue;
      for(int k=0;k<5;k++){
        requestFocus(vm,t); Pump(250);
        ACI2 fi;
        if(Inf(t,out fi) && Est(fi,"focused")){
          string nom=(fi.name==null?"":fi.name.Trim());
          if(nom.Length==0) nom="campo en y="+fi.y;
          return "OK|"+v.Length+"|"+nom;
        }
      }
      return "ERROR: no consigo el foco en el campo que trae '"+v+"'";
    }
    return "VACIO";
  }

  public static string LeerCF(){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    long c=CampoCF(cim); if(c==0) return "ERROR: no encuentro el campo";
    string v=Texto(c);
    return v==null?"ERROR: campo ilegible":("OK|"+v);
  }

  // COMMIT. En Swing el texto vive en el documento del campo, pero la aplicacion suele
  // tomarlo cuando el campo PIERDE EL FOCO. Sin esto, releer el campo dice que el numero
  // esta ahi y la busqueda sale igual vacia, porque Siebel nunca lo recogio.
  public static string CommitCF(string cuenta){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    long c=CampoCF(cim); if(c==0) return "ERROR: no encuentro el campo";
    long otro=First(cim,(i,ac)=> i.role_en_US=="text" && Est(i,"editable") && ac!=c);
    if(otro!=0){
      requestFocus(vm,otro);
      // CUIDADO CON LO QUE SE SONDEA AQUI (05/09/2026). Lo obvio seria sondear el valor del
      // campo hasta que cuadre, y esta MAL: el texto no cambia cuando Siebel lo recoge, asi
      // que ese sondeo saldria a la primera vuelta, a los 0 ms, sin haber esperado nada. Es
      // justo lo que dice el comentario de arriba -- releer el campo NO prueba el commit.
      //
      // La condicion que si es observable, y que ademas es el MECANISMO, es que el OTRO campo
      // gane el foco: eso significa que 'ID de CF' lo perdio, que es lo que dispara el commit.
      // Se pregunta antes de esperar, asi que si el foco aterriza al instante no cuesta nada.
      // Mismo patron que EnfocarCF y EnfocarSucio, unas lineas mas arriba.
      DateTime topeFoco=DateTime.Now.AddSeconds(3);
      bool movido=false;
      while(true){
        ACI2 fi;
        if(Inf(otro,out fi) && Est(fi,"focused")){ movido=true; break; }
        if(DateTime.Now>=topeFoco) break;
        requestFocus(vm,otro);
        Pump(150);
      }
      if(!movido) P("commit: el foco no llego al otro campo en 3 s; se sigue y se comprueba el valor igual");
      // Y DESPUES los 400 ms de siempre, sin recortar. Que el foco haya aterrizado no prueba
      // que Siebel terminara de procesarlo: eso no se le puede preguntar a nadie. Esta espera
      // es el unico margen que hay para el listener, y bajarla seria ir mas rapido justo en el
      // sitio donde nada avisa si sales antes de tiempo. El cambio de hoy solo AGREGA espera
      // cuando el foco tarda; nunca quita.
      Pump(400);
    }
    // Releer con reintentos en vez de una sola lectura. Si el campo se queda ilegible un
    // instante mientras Azul repinta, una lectura unica aborta el cliente entero -- y el
    // aborto es caro: buscar_cuenta.ps1 sale con exit 1 y el lote da la cuenta por fallida.
    // Si el valor de verdad se perdio, esto tarda 2 s mas en decir lo mismo.
    string tras=null;
    DateTime topeValor=DateTime.Now.AddSeconds(2);
    while(true){
      tras=Texto(c);
      if(tras!=null && tras.Trim()==cuenta) return "OK";
      if(DateTime.Now>=topeValor) break;
      Pump(200);
    }
    return "ERROR: el campo perdio el valor al cambiar el foco: '"+(tras==null?"(ilegible)":tras.Trim())+"'";
  }

  // Dos versiones a proposito. Quien YA tiene resuelto el marco CIM le pasa el suyo en vez
  // de hacer que se busque otra vez: cuando la busqueda acierta el formulario ya no existe,
  // y entonces Marco() no lo encuentra y cae en el recorrido completo del arbol. Firma()
  // pedia el mismo marco dos veces por llamada y pagaba ese recorrido dos veces.
  static long MarcoResultados(){ return MarcoResultados(Marco(F_CIM)); }
  static long MarcoResultados(long cim){
    foreach(long f in Find(root,(i,ac)=> i.role_en_US=="internal frame",60000,false)){
      if(f==cim) continue;
      if(First(f,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.IndexOf(B_CONT)>=0)!=0) return f;
    }
    return 0;
  }

  static long TablaDe(long marco){
    long mejor=0; int mx=0;
    foreach(long t in Find(marco,(i,ac)=> i.role_en_US=="table",60000,false)){
      ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) continue;
      int k=ti.rowCount*ti.columnCount;
      if(k>mx){ mx=k; mejor=t; }
    }
    return mejor;
  }

  // Foto del estado, en una linea. Las tres senales de "ya llego el resultado" van aqui:
  // que el formulario desaparezca, que 'Seleccionar' gane 'enabled', y que cambie el
  // contador de registros. Ninguna es una espera por reloj.
  public static string Firma(){
    long cim=Marco(F_CIM);
    long res=MarcoResultados(cim);   // se reaprovecha el cim de arriba, no se vuelve a buscar
    string sel="-", cont="-", tab="-";
    if(res!=0){
      long b=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==B_SELEC);
      if(b!=0){ ACI2 bi; if(Inf(b,out bi)) sel=Est(bi,"enabled")?"1":"0"; }
      long k=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.IndexOf(B_CONT)>=0);
      if(k!=0){ ACI2 ki; if(Inf(k,out ki)) cont=ki.name.Trim(); }
      long t=TablaDe(res);
      if(t!=0){ ATI2 ti; if(getAccessibleTableInfo(vm,t,out ti)) tab=ti.rowCount+"x"+ti.columnCount; }
    }
    // El numero de marcos entra en la firma para que un aviso emergente ("no se encontraron
    // registros", un error) cuente como cambio y no se quede esperando a los 90 s.
    int nm=Find(root,(i,ac)=> i.role_en_US=="internal frame" || i.role_en_US=="dialog" || i.role_en_US=="option pane",60000,false).Count;
    string it=Interaccion(); if(it.Length==0) it="-";
    return "marcos="+nm+" cim="+(cim!=0?"1":"0")+" res="+(res!=0?"1":"0")+" sel="+sel+" cont="+cont+" tabla="+tab+" inter="+it;
  }

  // El marco de interaccion, que es lo que Azul abre cuando ENCUENTRA al cliente. Su nombre
  // trae el numero de interaccion, que cambia en cada busqueda: por eso sirve para distinguir
  // "cargo este cliente ahora" de "ya habia uno cargado de antes".
  //
  // Un marco recien nacido se llama todavia "Inicio de Interaccion [%]": Siebel pinta el
  // titulo con su plantilla sin rellenar y pone el numero un momento despues. Ese marco NO
  // cuenta como cliente cargado -- es el cartel, no el cliente. Mientras el 10/09/2026 hubo
  // diez segundos de espera clavada antes de mirar, la plantilla nunca se llegaba a ver y
  // esto no se noto; el 18/09/2026, quitada esa espera, la primera busqueda devolvio "[%]",
  // el tramo de 20 s para abrir 'Suscripciones' arranco con Azul todavia montando la vista y
  // se agoto. Se descarta por el hueco sin rellenar y no por la falta del numero, que es lo
  // que de verdad distingue la plantilla de un nombre bueno.
  public static string Interaccion(){
    long f=First(root,(i,ac)=> i.role_en_US=="internal frame" && i.name!=null
                               && i.name.StartsWith(F_INT) && i.name.IndexOf('%')<0);
    if(f==0) return "";
    ACI2 fi; if(!Inf(f,out fi)) return "";
    if(fi.name==null) return "";
    string n=fi.name.Trim();
    return n.IndexOf('%')>=0 ? "" : n;
  }

  // Todo lo que hay abierto. Si Azul contesto con un aviso en vez de con resultados, aqui se
  // ve; si no se mira, un dialogo de error pasa por "no encontro nada".
  public static string Marcos(){
    var sb=new StringBuilder();
    foreach(long f in Find(root,(i,ac)=> i.role_en_US=="internal frame" || i.role_en_US=="dialog" || i.role_en_US=="option pane",60000,false)){
      ACI2 i; if(!Inf(f,out i)) continue;
      sb.AppendLine("  ["+i.role_en_US+"] '"+(i.name==null?"":i.name.Trim())+"'");
    }
    return sb.ToString();
  }

  // 'Buscar Ahora' NO se pulsa por el puente. doAccessibleActions lo dispara -- el formulario
  // se cierra, se ve en la firma -- pero Siebel no ejecuta la consulta, y la corrida parece
  // haber buscado sin haber buscado. Medido el 03/09/2026: dos corridas con cero resultados,
  // y el mismo formulario, con el mismo valor, devolvio el cliente en cuanto Dorian pulso el
  // boton con el mouse. Asi que aqui solo se da la GEOMETRIA y el clic real lo hace
  // captura.ps1, que es donde vive todo lo que toca la maquina.
  //
  // No se clavan coordenadas: al boton si lo ve el puente, asi que se le pregunta cada vez.
  public static string GeoBotonBuscar(){
    long cim=Marco(F_CIM); if(cim==0) return "ERROR: el formulario no esta abierto";
    long b=First(cim,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==B_BUSCAR);
    if(b==0) return "ERROR: no encuentro el boton '"+B_BUSCAR+"'";
    ACI2 bi;
    if(!Inf(b,out bi)) return "ERROR: no puedo leer la geometria de '"+B_BUSCAR+"'";
    if(!Est(bi,"enabled")) return "ERROR: '"+B_BUSCAR+"' esta deshabilitado";
    if(bi.width<=0 || bi.height<=0) return "ERROR: '"+B_BUSCAR+"' no tiene tamano utilizable";
    return "OK|"+bi.x+","+bi.y+","+bi.width+","+bi.height;
  }

  // Espera a que la firma CAMBIE y luego a que deje de cambiar. Lo segundo importa tanto
  // como lo primero: el formulario se cierra antes de que la rejilla termine de llenarse, y
  // mirar en ese hueco daria un conteo a medias. Es el mismo principio que ya gobierna la
  // espera del arbol de atributos y la de la captura.
  public static string EsperarResultado(string antes,int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    string f=antes;
    while(DateTime.Now<tope){
      f=Firma();
      if(f!=antes) break;
      Pump(300);
    }
    if(f==antes) return "TIEMPO|"+f;
    string prev=f; int iguales=0;
    while(DateTime.Now<tope){
      Pump(500);
      f=Firma();
      if(f==prev){ iguales++; if(iguales>=2) break; } else { iguales=0; prev=f; }
    }
    return "OK|"+f;
  }

  // ---- esperar POR RECONOCIMIENTO, no por reloj ----------------------------
  // Despues de 'Buscar Ahora' Azul hace una de dos cosas, y las dos se ven en cuanto pasan:
  // o carga al cliente y abre su marco de interaccion, o abre la rejilla para que se elija
  // cuenta. EsperarResultado no miraba ninguna de las dos: miraba una firma de pantalla, y
  // encima esperaba a que esa firma dejara de moverse. Eso son varios segundos DESPUES de que
  // en pantalla ya esta lo que hace falta -- que es lo que se veia como "piensa demasiado".
  //
  //   CARGADO|<marco>   el marco de interaccion aparecio y no es el que ya habia
  //   REJILLA|<fila 1>  la rejilla ya tiene primera fila con datos y 'Seleccionar' en pantalla
  //   NADA|             todavia no se sabe
  public static string Desenlace(string interAntes){
    string r=EsperarInteraccion(interAntes,0);
    if(r.StartsWith("CARGADO|")) return r;
    long res=MarcoResultados(); if(res==0) return "NADA|";
    long t=TablaDe(res); if(t==0) return "NADA|";
    ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) return "NADA|";
    if(ti.rowCount<1 || ti.columnCount<1) return "NADA|";
    // Sin el boton no hay nada que pulsar todavia: la rejilla esta a medio montar.
    long b=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==B_SELEC);
    if(b==0) return "NADA|";
    var sb=new StringBuilder();
    for(int c=0;c<ti.columnCount;c++){ sb.Append(Cel(t,c)); sb.Append(" "); }
    string fila0=sb.ToString().Trim();
    if(fila0.Length==0) return "NADA|";   // rejilla vacia de Siebel: fila en blanco, no resultado
    return "REJILLA|"+fila0;
  }

  // Solo el marco de interaccion. Es lo que se espera despues de 'Seleccionar', donde la
  // rejilla ya no es un desenlace sino lo que se acaba de dejar atras. Con seg=0 es una
  // mirada suelta, sin espera.
  public static string EsperarInteraccion(string interAntes,int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(true){
      string inter=Interaccion();
      if(inter.Length>0 && inter!=interAntes) return "CARGADO|"+inter;
      if(DateTime.Now>=tope) return "NADA|";
      Pump(200);
    }
  }

  // La primera fila se exige LEIDA DOS VECES IGUAL. Siebel pinta la rejilla mientras la llena,
  // y sobre esa misma fila se comprueba despues la razon social del Excel: leerla a medias
  // haria fallar la comprobacion en un cliente bueno. Son 200 ms, no los segundos de antes.
  public static string EsperarDesenlace(string interAntes,int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    string previo="";
    while(true){
      string d=Desenlace(interAntes);
      if(d.StartsWith("CARGADO|")) return d;
      if(d.StartsWith("REJILLA|")){
        if(d==previo) return d;
        previo=d;
      } else previo="";
      if(DateTime.Now>=tope) return "NADA|";
      Pump(200);
    }
  }

  // Lo que trajo la busqueda, para poder pararse aqui sabiendo que se trajo.
  // Cuantas filas con CONTENIDO trae la rejilla. Una rejilla vacia de Siebel no tiene cero
  // filas: tiene una fila en blanco, que es como se veia el "1x9" de las corridas que no
  // encontraron nada. Contar filas a secas daria un resultado donde no lo hay.
  public static int FilasConDatos(){
    long res=MarcoResultados(); if(res==0) return 0;
    long t=TablaDe(res); if(t==0) return 0;
    ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) return 0;
    int n=0;
    for(int r=0;r<ti.rowCount;r++){
      for(int c=0;c<ti.columnCount;c++){
        long cc=Kid(t,r*ti.columnCount+c); if(cc==0) continue;
        ACI2 ci; if(!Inf(cc,out ci)) continue;
        if(ci.name!=null && ci.name.Trim().Length>0){ n++; break; }
      }
    }
    return n;
  }

  // Marca una fila de la rejilla. La prueba de que funciono no es que la llamada no falle,
  // sino que 'Seleccionar' pase a estar habilitado: ese boton esta apagado mientras no hay
  // fila elegida, asi que es la confirmacion que el propio Azul da.
  public static string SeleccionarFila(int fila){
    long res=MarcoResultados(); if(res==0) return "ERROR: no hay rejilla de resultados";
    long t=TablaDe(res); if(t==0) return "ERROR: la rejilla no tiene tabla";
    ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) return "ERROR: la tabla no da informacion";
    if(fila<0 || fila>=ti.rowCount) return "ERROR: la fila "+fila+" no existe (hay "+ti.rowCount+")";
    long b=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==B_SELEC);
    if(b==0) return "ERROR: no encuentro el boton '"+B_SELEC+"'";
    clearAccessibleSelectionFromContext(vm,t); Pump(120);
    addAccessibleSelectionFromContext(vm,t,fila*ti.columnCount);
    for(int w=0;w<12;w++){
      ACI2 bi;
      if(Inf(b,out bi) && Est(bi,"enabled")) return "OK";
      Pump(150);
    }
    return "NO";   // no se habilito: que lo intente PowerShell con un clic real
  }

  // Todo el texto de una fila, para poder comprobar CONTRA EL EXCEL que la fila que se va a
  // elegir es la del cliente pedido. Elegir la fila equivocada mete al cliente equivocado en
  // el documento, y eso no se ve hasta que ya esta escrito.
  public static string FilaTexto(int fila){
    long res=MarcoResultados(); if(res==0) return "";
    long t=TablaDe(res); if(t==0) return "";
    ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) return "";
    if(fila<0 || fila>=ti.rowCount) return "";
    var sb=new StringBuilder();
    for(int c=0;c<ti.columnCount;c++){ sb.Append(Cel(t,fila*ti.columnCount+c)); sb.Append(" "); }
    return sb.ToString();
  }

  // Geometria de una celda, para poder pulsar la fila de verdad si el puente no la marca.
  public static string GeoCelda(int fila,int col){
    long res=MarcoResultados(); if(res==0) return "ERROR: no hay rejilla de resultados";
    long t=TablaDe(res); if(t==0) return "ERROR: la rejilla no tiene tabla";
    ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) return "ERROR: la tabla no da informacion";
    if(fila<0 || fila>=ti.rowCount || col<0 || col>=ti.columnCount) return "ERROR: celda fuera de la tabla";
    long c=Kid(t,fila*ti.columnCount+col); if(c==0) return "ERROR: celda vacia";
    ACI2 ci; if(!Inf(c,out ci)) return "ERROR: no puedo leer la celda";
    if(ci.width<=0||ci.height<=0) return "ERROR: la celda no tiene tamano utilizable";
    return "OK|"+ci.x+","+ci.y+","+ci.width+","+ci.height;
  }

  // Geometria de un boton de la rejilla de resultados ('Seleccionar'), para el clic real.
  public static string GeoBotonResultados(string nombre){
    long res=MarcoResultados(); if(res==0) return "ERROR: no hay rejilla de resultados";
    long b=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==nombre);
    if(b==0) return "ERROR: no encuentro el boton '"+nombre+"' en la rejilla";
    ACI2 bi; if(!Inf(b,out bi)) return "ERROR: no puedo leer la geometria de '"+nombre+"'";
    if(!Est(bi,"enabled")) return "ERROR: '"+nombre+"' esta deshabilitado";
    if(bi.width<=0||bi.height<=0) return "ERROR: '"+nombre+"' no tiene tamano utilizable";
    return "OK|"+bi.x+","+bi.y+","+bi.width+","+bi.height;
  }

  // ---- pestanas del cliente ----------------------------------------------
  // Las pestanas (Viaje, Informacion de Contacto, ..., Suscripciones) SI estan en el arbol,
  // como 'page tab', con geometria y con el estado 'selected' que dice cual esta abierta.
  // Pero vienen con act=no: no tienen accion de accesibilidad, asi que hay que pulsarlas de
  // verdad. Otra vez lo mismo: el puente las ve y no las conduce.
  //
  // La pestana se busca UNA vez y el nodo se guarda. MEDIDO el 11/09/2026 sobre 27 clientes
  // con el contador de llamadas JAB que ya deja este archivo en el progreso: un recorrido
  // completo del arbol cuesta ~22.000 llamadas, y el paso de la pestana gastaba de 90.000 a
  // 109.000 -- cuatro o cinco recorridos -- contra ~22.000, uno solo, en los clientes donde
  // la pestana ya estaba abierta de antes. Los tres sitios de abajo recorrian el arbol entero
  // por su cuenta, y EsperarPestana lo recorria ADEMAS en cada vuelta de sondeo. De ahi salia
  // el tope de tramo, que entonces era de 10 s: se los comia el puente buscando cuatro veces
  // la misma pestana, no Azul contestando. Con el cache puesto y MEDIDO el mismo dia sobre un
  // cliente sano, el tramo bajo a 40.604 llamadas y 9 s, y por eso el tope subio a 20 s.
  //
  // Es el mismo reparto que ya hacen EsperarCIM y MarcoPanes mas arriba: el recorrido completo
  // FUERA de los bucles de espera, y dentro solo lecturas baratas sobre un nodo ya conocido.
  //
  // OJO, limitacion que esto NO arregla: First() devuelve la PRIMERA pestana con ese nombre en
  // orden de arbol. Con un marco de interaccion viejo abierto -- que pasa, y paso el 11/09 --
  // esa puede ser la del marco anterior, y entonces el clic va a la pestana equivocada. Es
  // candidato a explicar el clic medido en (895,87) cuando el resto fueron en (895,341).
  static long pestCache=0;
  static string pestCacheNombre="";

  // El nodo de la pestana: del cache si sigue siendo el que se pidio, y del arbol si no. Un
  // handle JAB se queda viejo si Azul reconstruye la vista, asi que no se confia en el a ciegas:
  // se le relee el papel y el nombre, y solo se reusa si los dos siguen cuadrando.
  static long PestanaNodo(string nombre){
    if(pestCache!=0 && pestCacheNombre==nombre){
      ACI2 ci;
      if(Inf(pestCache,out ci) && ci.role_en_US=="page tab" && ci.name!=null && ci.name.Trim()==nombre) return pestCache;
      pestCache=0; pestCacheNombre="";
    }
    long t=First(root,(i,ac)=> i.role_en_US=="page tab" && i.name!=null && i.name.Trim()==nombre);
    if(t!=0){ pestCache=t; pestCacheNombre=nombre; }
    return t;
  }
  public static string GeoPestana(string nombre){
    long t=PestanaNodo(nombre);
    if(t==0) return "ERROR: no encuentro la pestana '"+nombre+"'";
    ACI2 ti; if(!Inf(t,out ti)) return "ERROR: no puedo leer la geometria de la pestana '"+nombre+"'";
    if(ti.width<=0||ti.height<=0) return "ERROR: la pestana '"+nombre+"' no tiene tamano utilizable";
    return "OK|"+ti.x+","+ti.y+","+ti.width+","+ti.height;
  }
  public static bool PestanaSeleccionada(string nombre){
    long t=PestanaNodo(nombre);
    if(t==0) return false;
    ACI2 ti; if(!Inf(t,out ti)) return false;
    return Est(ti,"selected");
  }
  // Espera a que la pestana exista Y SE ESTE QUIETA, no a que este seleccionada. Hace falta desde
  // que la espera de la busqueda dejo de ser por reloj: ahora se llega aqui en cuanto aparece el
  // marco de interaccion, y eso puede ser antes de que sus pestanas esten colocadas.
  //
  // Lo de "quieta" no es adorno. El 17/09/2026, con Azul de verdad, CLIENTE EJEMPLO DOS y
  // CLIENTE EJEMPLO TRES murieron los dos igual: la pestana existia ya, pero estaba en
  // x=826 y un instante despues se fue a x=895. El clic salio a la posicion vieja, cayo en otro
  // sitio y 'Suscripciones' no llego a seleccionarse nunca. Leer la geometria dos veces seguidas
  // y exigir que de lo mismo cuesta 300 ms y cierra ese agujero; ademas evita pulsar a ciegas
  // donde ni siquiera se sabe que hay.
  public static bool EsperarPestanaExiste(string nombre,int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    string previo="";
    while(true){
      string g=GeoPestana(nombre);
      if(g.StartsWith("OK|")){
        if(g==previo) return true;
        previo=g;
      } else previo="";
      if(DateTime.Now>=tope) return false;
      Pump(300);
    }
  }
  public static bool EsperarPestana(string nombre,int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      if(PestanaSeleccionada(nombre)) return true;
      Pump(400);
    }
    return PestanaSeleccionada(nombre);
  }

  static string Cel(long tbl,int idx){
    long c=Kid(tbl,idx); if(c==0) return "";
    ACI2 i; if(!Inf(c,out i)) return "";
    return i.name==null?"":i.name.Trim();
  }

  // La tabla de Suscripciones se reconoce igual que en azul_fast.ps1: N x >=10 columnas y un
  // radio button en la celda 0. Columna 1 = numero, columna 7 = estado.
  public static string ResumenSuscripciones(int maxFilas){
    long mejor=0; int R=0,C=0;
    foreach(long t in Find(root,(i,ac)=> i.role_en_US=="table",60000,false)){
      ATI2 ti; if(!getAccessibleTableInfo(vm,t,out ti)) continue;
      if(ti.columnCount<10||ti.rowCount<1) continue;
      long c0=Kid(t,0); if(c0==0) continue;
      ACI2 i0; if(!Inf(c0,out i0)) continue;
      if(i0.role_en_US!="radio button") continue;
      if(ti.rowCount*ti.columnCount > R*C){ mejor=t; R=ti.rowCount; C=ti.columnCount; }
    }
    if(mejor==0) return "";
    var sb=new StringBuilder();
    sb.AppendLine("Tabla de Suscripciones: "+R+" filas x "+C+" columnas");
    int n=Math.Min(R,maxFilas);
    for(int r=0;r<n;r++){
      sb.AppendLine("  "+Cel(mejor,r*C+1)+"   "+((C>7)?Cel(mejor,r*C+7):""));
    }
    if(R>n) sb.AppendLine("  ... y "+(R-n)+" mas");
    return sb.ToString();
  }

  public static string Resultado(int maxFilas){
    var sb=new StringBuilder();
    long res=MarcoResultados();
    if(res==0){ sb.AppendLine("No hay ninguna rejilla de resultados abierta."); return sb.ToString(); }
    ACI2 ri; if(Inf(res,out ri)) sb.AppendLine("Rejilla: "+(ri.name==null?"":ri.name.Trim()));
    long k=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.IndexOf(B_CONT)>=0);
    if(k!=0){ ACI2 ki; if(Inf(k,out ki)) sb.AppendLine("Contador: "+ki.name.Trim()); }
    long b=First(res,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()==B_SELEC);
    if(b!=0){ ACI2 bi; if(Inf(b,out bi)) sb.AppendLine("Boton 'Seleccionar' habilitado: "+(Est(bi,"enabled")?"si":"no")); }
    long t=TablaDe(res);
    if(t==0){ sb.AppendLine("Sin tabla de resultados."); return sb.ToString(); }
    ATI2 ti;
    if(!getAccessibleTableInfo(vm,t,out ti)){ sb.AppendLine("La tabla no da informacion."); return sb.ToString(); }
    sb.AppendLine("Tabla: "+ti.rowCount+" filas x "+ti.columnCount+" columnas");
    int nf=Math.Min(ti.rowCount,maxFilas);
    for(int r=0;r<nf;r++){
      var cel=new List<string>();
      for(int c=0;c<Math.Min(ti.columnCount,12);c++){
        long cc=Kid(t,r*ti.columnCount+c);
        string v="";
        if(cc!=0){ ACI2 ci; if(Inf(cc,out ci)) v=(ci.name==null?"":ci.name.Trim()); }
        cel.Add(v);
      }
      sb.AppendLine("  fila "+(r+1)+": "+string.Join(" | ",cel.ToArray()));
    }
    return sb.ToString();
  }
}
'@

. (Join-Path $PSScriptRoot 'captura.ps1')
# De aqui sale el numero de cuenta cuando se pasa -Fila. Lo comparte con lote.ps1: antes
# habia una copia de este lector en cada archivo, y dos copias acaban divergiendo.
. (Join-Path $PSScriptRoot 'lista.ps1')
if ($Lista -eq "") { $Lista = Join-Path $PSScriptRoot 'lista_ordenes.csv' }

# ---- de donde sale el numero de cuenta --------------------------------------
$razon = $Razon
if ($Cuenta -eq "") {
  if ($Fila -le 0) { "ERROR: pasa -Cuenta <numero> o -Fila <n> de la lista de ordenes."; exit 1 }
  $o = Get-OrdenDeLista -Ruta $Lista -Fila $Fila
  $Cuenta = $o.Cuenta
  $razon  = $o.Razon
  "Lista: $(Split-Path $Lista -Leaf), fila $Fila"
}
if ($Cuenta -notmatch '^\d{4,15}$') {
  "ERROR: '$Cuenta' no parece un numero de cuenta. No se busca a ciegas."; exit 1
}
"Cuenta: $Cuenta$(if ($razon -ne '') { "  ($razon)" })"
""

if ($Restaurar) {
  # Frente() hace el baile de AttachThreadInput que hace falta contra Swing, y restaura la
  # ventana justo cuando esta minimizada, que es el caso que se comprueba aqui abajo. Es el
  # mismo camino que usa el clic autorizado, no uno nuevo.
  $pz = Get-Process -Name jp2launcher -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if ($pz -and [AzulShot]::IsIconic($pz.MainWindowHandle)) {
    "Azul estaba minimizado: se restaura."
    [void][AzulShot]::Frente($pz.MainWindowHandle)
    Start-Sleep -Milliseconds 800
  }
}
try { $hwnd = Get-AzulHwnd } catch { "ERROR: $($_.Exception.Message)"; exit 1 }
[Busca]::PROG = Join-Path $PSScriptRoot 'progreso.txt'
if (-not [Busca]::Init($hwnd)) { "ERROR: sin contexto raiz. El puente no ve a Azul."; exit 1 }

# Auxiliares comunes. Viven FUERA del bloque de busqueda porque -DesdeResultados se lo
# salta entero y aun asi las necesita.
function Save-Prueba {
  param([string]$Etiqueta)
  [void][AzulShot]::Frente($hwnd)
  Start-Sleep -Milliseconds 400
  $f = Join-Path $PSScriptRoot ("salidas\busquedas\" + $Etiqueta + "_" + $Cuenta + "_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".png")
  [void](Save-AzulShot -Hwnd $hwnd -Ruta $f)
  return $f
}

function Invoke-ClicGeo {
  param(
    # Se pasa un bloque, no una cadena ya calculada, para que la geometria se lea DESPUES de
    # subir la ventana al frente: esa es justo la carrera que hacia fallar el clic.
    [Parameter(Mandatory=$true)][scriptblock]$Geometria,
    [string]$Que = "el boton"
  )
  if (-not [AzulShot]::Frente($hwnd)) { return "ERROR: no se pudo traer Azul al frente." }
  Start-Sleep -Milliseconds 250
  $g0 = & $Geometria
  if ($g0 -like 'ERROR*') { return $g0 }
  $g = (($g0 -split '\|')[1]).Split(',')
  $rc = Get-AzulRect -Hwnd $hwnd
  $x = [int]$g[0] - $rc.Left + [int][Math]::Floor([int]$g[2] / 2)
  $y = [int]$g[1] - $rc.Top  + [int][Math]::Floor([int]$g[3] / 2)
  [Busca]::Nota("clic real en $Que en ($x,$y)")
  [void](Invoke-AzulClick -Hwnd $hwnd -X $x -Y $y)
  return "OK|$x,$y"
}

function Invoke-Buscar { Invoke-ClicGeo -Geometria { [Busca]::GeoBotonBuscar() } -Que "'Buscar Ahora'" }

# Con -DesdeResultados se salta toda la busqueda y se sigue desde la rejilla que ya
# esta abierta. Sirve para retomar sin repetir la consulta, y porque con una rejilla
# de resultados abierta Azul IGNORA el clic en el icono del telefono.
if ($DesdeResultados) {
  $interAntes = [Busca]::Interaccion()
} else {
  # ---- 1. abrir el formulario -------------------------------------------------
  if ([Busca]::HayCIM()) {
    "Formulario de busqueda: ya estaba abierto, no se toca el mouse."
  } elseif ($SinClic) {
    "ERROR: el formulario no esta abierto y se pidio -SinClic."; exit 1
  } else {
    # Donde picar no se recuerda: se mira. Ver Find-IconoLupa en captura.ps1 y el porque
    # en el comentario de -IconoX. Con -IconoX/-IconoY a mano se pica ahi sin mirar, que
    # es la salida de emergencia por si algun dia el dibujo del icono cambia.
    if ($IconoX -gt 0 -and $IconoY -gt 0) {
      $ix = $IconoX; $iy = $IconoY
      "Clic real en el icono del telefono con lupa ($ix,$iy), posicion forzada a mano."
      [Busca]::Nota("buscar: clic forzado en el icono en ($ix,$iy)")
    } else {
      $lupa = Find-IconoLupa -Hwnd $hwnd
      if (-not $lupa.Ok) {
        "ERROR: no se encontro el icono del telefono con lupa en la ventana de Azul."
        "       $($lupa.Por)"
        "       No se pica a ciegas: en esa barra los vecinos abren otras cosas."
        [Busca]::Nota("buscar: icono no encontrado -- $($lupa.Por)")
        exit 1
      }
      $ix = $lupa.X; $iy = $lupa.Y
      "Icono del telefono con lupa encontrado en $ix,$iy  ($($lupa.Detalle))."
      [Busca]::Nota("buscar: icono encontrado en ($ix,$iy), $($lupa.Detalle)")
    }
    [void](Invoke-AzulClick -Hwnd $hwnd -X $ix -Y $iy)
    if (-not [Busca]::EsperarCIM(40)) {
      "ERROR: el formulario 'Encontrar Comunicante Busqueda CIM' no abrio en 40 s."
      "       El clic dio en $ix,$iy. O Azul no respondio, o el icono no era ese."
      exit 1
    }
    "Formulario abierto."
  }
  
  # ---- 2. criterios viejos ----------------------------------------------------
  $otros = [Busca]::CamposConTexto()
  if ($otros -like 'ERROR*') { $otros; exit 1 }
  if ($otros.Trim() -ne "") {
    "El formulario traia criterios en otros campos:"
    $otros.TrimEnd()
    if ($Forzar) {
      "Se buscara CON ellos por -Forzar."
    } else {
      # Se limpian con el mismo teclado que llena ID de CF. Son campos de un formulario de
      # consulta, no datos del cliente: borrarlos no modifica nada de la cuenta. Dejarlos si
      # tendria consecuencias, porque se combinan con la cuenta y devuelven otro cliente.
      if (-not [AzulShot]::Frente($hwnd)) {
        "ERROR: no se pudo traer Azul al frente para limpiar el formulario."; exit 1
      }
      for ($n = 0; $n -lt 25; $n++) {
        $s = [Busca]::EnfocarSucio()
        if ($s -eq 'VACIO') { break }
        if ($s -like 'ERROR*') { $s; exit 1 }
        $p = $s -split '\|'
        $len = [int]$p[1]
        [AzulShot]::Tecla(0x23)                                        # FIN
        for ($i = 0; $i -lt ($len + 3); $i++) { [AzulShot]::Tecla(0x08) }   # RETROCESO
        [Busca]::Espera(200)
        "  limpiado: $($p[2])"
      }
      $otros = [Busca]::CamposConTexto()
      if ($otros.Trim() -ne "") {
        "ERROR: no pude dejar limpios estos campos:"
        $otros.TrimEnd()
        "Limpialos en Azul y reintenta, o usa -Forzar para buscar con ellos."
        exit 1
      }
      "Formulario limpio: solo queda el criterio de la cuenta."
    }
  }
  
  # ---- 3. llenar el campo y comprobarlo ---------------------------------------
  # Siempre con teclado real. El puente escribe en el documento del campo pero Azul no se da
  # por enterado, y la busqueda sale vacia con el numero a la vista. Ver el comentario del C#.
  # Hasta tres intentos. Azul se come pulsaciones cuando esta ocupado: la primera vez que
  # paso, de '507000002' solo entro '507' -- y un numero truncado no falla, DEVUELVE OTRO
  # CLIENTE. Por eso no basta con teclear: hay que releer, y si no cuadra, borrar y repetir.
  $val = ""
  for ($intento = 1; $intento -le 3; $intento++) {
    $r = [Busca]::EnfocarCF()
    if ($r -like 'ERROR*') { $r; exit 1 }
    $largo = [int](($r -split '\|')[1])
    if (-not [AzulShot]::Frente($hwnd)) {
      "ERROR: no se pudo traer Azul al frente. Sin eso las pulsaciones irian a otra ventana."
      exit 1
    }
    # Fin de linea y borrar lo que hubiera, tantas veces como caracteres habia (mas margen).
    # 0x23 = FIN, 0x08 = RETROCESO.
    [AzulShot]::Tecla(0x23)
    for ($i = 0; $i -lt ($largo + 3); $i++) { [AzulShot]::Tecla(0x08) }
    [AzulShot]::Escribir($Cuenta)

    # Espera activa: se pregunta por el valor en vez de dormir una cantidad fija, que es lo
    # que se quedaba corto. En cuanto cuadra, se sigue.
    # Se PREGUNTA antes de esperar. Escribir() ya durmio 40 ms por caracter, asi que para una
    # cuenta de 9 digitos el valor lleva ~360 ms puesto cuando se llega aqui: la espera de
    # entrada era 250 ms que pagaba igual el tecleo que habia salido perfecto. El presupuesto
    # del peor caso no cambia, son las mismas 20 vueltas.
    for ($w = 0; $w -lt 20; $w++) {
      $r = [Busca]::LeerCF()
      if ($r -like 'ERROR*') { $r; exit 1 }
      $val = (($r -split '\|', 2)[1]).Trim()
      if ($val -eq $Cuenta) { break }
      [Busca]::Espera(250)
    }
    if ($val -eq $Cuenta) { break }
    if ($intento -lt 3) { "El campo quedo con '$val'; se borra y se reintenta." }
  }
  if ($val -ne $Cuenta) {
    "ERROR: tras 3 intentos el campo quedo con '$val' y se pidio '$Cuenta'. No se busca con eso."
    exit 1
  }
  
  # El valor no cuenta hasta que Azul lo recoge, y lo recoge al perder el foco.
  $r = [Busca]::CommitCF($Cuenta)
  if ($r -like 'ERROR*') { $r; exit 1 }
  "Campo 'ID de CF' = $Cuenta  (releido de Azul, no supuesto)"
  
  # Con Azul detras de otras ventanas, PrintWindow devuelve en negro las zonas de Chromium.
  # Se sube al frente antes de fotografiar; ya esta al frente de todos modos si hubo clic.
  
  if ($SoloLlenar) {
    $f = Save-Prueba -Etiqueta 'llenado'
    "Captura del formulario lleno: $f"
    ""
    "ALTO por -SoloLlenar: el campo esta puesto y NO se pulso 'Buscar Ahora'."
    exit 0
  }

  # Pausa ADICIONAL entre escribir el numero de cuenta y pulsar "Buscar Ahora" (Dorian,
  # 07/09/2026): ritmo mas humano para no golpear a Azul mas rapido de lo que aguanta.
  Start-Sleep -Seconds $PausaEscribir

  # ---- 4. buscar --------------------------------------------------------------
  $antes = [Busca]::Firma()
  $interAntes = [Busca]::Interaccion()
  "Antes:  $antes"
  
  # Clic REAL sobre el centro del boton. La geometria la da el puente en coordenadas de
  # pantalla; Invoke-AzulClick las quiere relativas a la ventana, asi que se resta su esquina,
  # igual que hace Close-MarcoCuenta en azul_fast.ps1.
  #
  # ORDEN IMPORTANTE: primero se sube Azul al frente, DESPUES se mide. Medido el 03/09/2026:
  # la ventana paso de (0,0) 1366x768 a maximizada (-8,-8) 1382x736 entre el calculo y el clic,
  # y el punto quedo viejo por 9 px. Frente() ya no cambia el tamano ni la posicion, pero el
  # orden se conserva: midiendo despues de subirla, las dos lecturas -- geometria del boton y
  # esquina de la ventana -- hablan de la misma ventana en el mismo instante, sea cual sea lo
  # que haya hecho el gestor de ventanas por su cuenta.
  
  
  $r = Invoke-Buscar
  if ($r -like 'ERROR*') { $r; exit 1 }
  "'Buscar Ahora' pulsado con clic real en $(($r -split '\|')[1])."


  # Aqui habia un Start-Sleep -Seconds $PausaResultado: diez segundos clavados antes de
  # mirar siquiera, se hubiera enterado Azul o no. Tenia sentido cuando el resultado se
  # reconocia por una firma de pantalla; desde que existe EsperarDesenlace ya no, porque esa
  # sonda no vuelve por un cambio cualquiera sino solo cuando reconoce el marco de
  # interaccion, o una rejilla con su primera fila leida dos veces igual. Era el tramo donde
  # mas tiempo se iba (Dorian, 18/09/2026: "tarda 20 segundos, es donde mas se va el tiempo").
  #
  # Queda un respiro corto, para no preguntarle a Azul en el mismo instante del clic. Y los
  # segundos de -PausaResultado no se pierden: se SUMAN al presupuesto de la sonda. En el
  # peor caso se aguanta lo mismo que antes; en el normal se sale en cuanto hay resultado.
  Start-Sleep -Milliseconds 700

  # Un reintento, y solo si el formulario SIGUE abierto: eso prueba que el clic no llego al
  # boton. A diferencia de la pulsacion por el puente, un clic fallido no consume el formulario,
  # asi que reintentar es seguro y no deja una corrida ambigua.
  $r = [Busca]::EsperarDesenlace($interAntes, 15 + $PausaResultado)
  if ((($r -split '\|')[0]) -eq 'NADA' -and [Busca]::HayCIM()) {
    "Nada cambio y el formulario sigue abierto: el clic no dio en el boton. Un reintento."
    $r2 = Invoke-Buscar
    if ($r2 -like 'ERROR*') { $r2; exit 1 }
    "'Buscar Ahora' pulsado de nuevo en $(($r2 -split '\|')[1])."
  }

  # LA ESPERA DE VERDAD: hasta minuto y medio para que Azul traiga algo. Solo se entra si la
  # sonda de arriba no reconocio nada todavia.
  #
  # Antes esto NO se podia condicionar, y hay un aviso viejo en este mismo sitio que lo explica:
  # con EsperarResultado, la sonda de 15 s devolvia "OK" en cuanto la firma de pantalla cambiaba,
  # y el formulario cerrandose YA la cambiaba -- asi que "OK" no significaba que Azul hubiera
  # traido nada. Condicionar sobre eso dejaba el presupuesto real en 15 s y los clientes lentos
  # salian como "no encontrado" (paso el 04/09/2026 con CLIENTE EJEMPLO UNO).
  #
  # Con EsperarDesenlace esa trampa desaparece: no vuelve por un cambio de pantalla cualquiera,
  # sino solo cuando reconoce el marco de interaccion o una rejilla con primera fila. 'NADA' es
  # de verdad "todavia no hay nada", asi que seguir esperando es lo correcto, y volver antes
  # cuando si hay algo es justo lo que se busca.
  if ((($r -split '\|')[0]) -eq 'NADA') { $r = [Busca]::EsperarDesenlace($interAntes, 90) }
  $desenlace = ($r -split '\|')[0]
  "Despues: $desenlace $((($r -split '\|', 2)[1]).Trim())"
  ""

  # Con 'NADA' NO se sale aqui. Se deja caer a las guardas de abajo, que miran el marco y las
  # filas y saben decir cual de los tres desenlaces malos es (marco heredado del cliente
  # anterior, sin resultado, o rejilla a medias). Un ERROR generico en este punto se llevaria
  # por delante esos mensajes, que son los que dicen que hacer.
  if ($desenlace -eq 'NADA') {
    "AVISO: se agoto el plazo sin reconocer ni marco de interaccion nuevo ni rejilla."
  }

  # Marca de medicion. Aisla el costo del bloque de busqueda -- que es donde viven las
  # esperas de resultado y los sondeos de marco -- del resto de la corrida.
  [Busca]::Nota("buscar_cuenta: JAB tras la busqueda = " + [Busca]::Llamadas())
}
# ---- encontro o no encontro -------------------------------------------------
# Que el formulario se cierre NO es encontrar: se cierra igual cuando no hay coincidencia.
# La senal buena es que Azul abra un marco de interaccion NUEVO, con al cliente cargado.
# Tres desenlaces, no dos:
#   1. coincidencia unica  -> Azul carga al cliente y abre un marco de interaccion
#   2. varias coincidencias -> abre una rejilla para elegir. ES un resultado, no un fallo:
#      aqui se para y se muestra, porque elegir cual es el cliente no me toca a mi.
#   3. ninguna             -> ni marco ni filas
$interDespues = [Busca]::Interaccion()
$filas   = [Busca]::FilasConDatos()
# Azul REUTILIZA el marco de interaccion: el nombre puede ser identico al de otro cliente
# cargado antes ("Inicio de Interaccion [2] - 682000001" salio igual para dos clientes
# distintos). Asi que "existe el marco" es la condicion, y si ademas el nombre no cambio se
# avisa, en vez de dar por bueno un marco que podria ser el del cliente anterior.
#
# 04/09/2026: esto era un AVISO y la corrida seguia con codigo 0. Alineado ahora con la rama
# de rejilla de mas abajo, donde la carga solo cuenta si el marco CAMBIO de nombre. El aviso
# suelto no servia de nada en el lote: lote.ps1 solo mira el codigo de salida, daba la
# busqueda por buena y leia las lineas del cliente ANTERIOR bajo el nombre del cliente
# pedido. Un marco con el mismo nombre y sin filas no prueba ninguna carga: se aborta.
#
# 04/09/2026, segundo arreglo. "$cargado = existe el marco" a secas era demasiado flojo: la
# rama de mas abajo hace "if ($cargado) {...} elseif ($filas -gt 0) {...}", asi que un marco
# heredado del cliente anterior GANABA sobre la rejilla. En la corrida del 04/09 habia rejilla
# (1 fila, tabla 1x5) y aun asi se declaro CLIENTE CARGADO sin pasar por la seleccion de fila
# ni por la comprobacion contra la razon social del Excel. Salio bien de casualidad: el marco
# resulto ser nuevo. Es justo la condicion del segundo cliente de un lote encadenado.
#
# Ahora "cargado" significa lo mismo aqui que en la rama de rejilla (ver mas abajo, tras
# 'Seleccionar'): marco que existe Y que CAMBIO de nombre. Con eso, marco viejo + rejilla cae
# solo en la rama de rejilla, que si selecciona la fila y si comprueba la razon social.
# El caso "marco viejo y sin filas" no cambia: lo sigue atajando la guarda de aqui debajo.
$hayMarco = ($interDespues -ne "")
$cargado  = ($hayMarco -and $interDespues -ne $interAntes)
if ($hayMarco -and $interDespues -eq $interAntes -and $filas -le 0) {
  "ERROR: ya habia un marco de interaccion con este mismo nombre antes de buscar, y Azul no"
  "       trajo filas. Nada prueba que se haya cargado la cuenta '$Cuenta': lo que se ve"
  "       seria del cliente anterior. No se sigue."
  "       marco antes:   '$interAntes'"
  "       marco despues: '$interDespues'"
  "Captura: $(Save-Prueba -Etiqueta 'marco_preexistente')"
  exit 1
}

# ---- 5. hasta aqui ----------------------------------------------------------
# Resumen de una linea. El volcado completo de marcos y filas era andamio de exploracion:
# ahora solo se imprime cuando algo falla, que es cuando de verdad hace falta verlo.
""
# Solo se fotografia lo que sale mal. La imagen que importa -- la vista general del cliente --
# la saca azul_fast.ps1 y esa si va al documento.
if (-not $cargado -and $filas -le 0) {
  "Captura: $(Save-Prueba -Etiqueta 'sinresultado')"
  [Busca]::Marcos().TrimEnd()
  # Se calcula AQUI, que es la unica rama que lo usa. Antes se calculaba siempre: Resultado()
  # recorre el arbol buscando la rejilla y lee hasta 3 filas x 12 columnas, asi que toda
  # corrida que salia BIEN pagaba ~72 llamadas al puente por un texto que se tiraba.
  $resumenRejilla = [Busca]::Resultado(3).TrimEnd()
  if ($resumenRejilla -ne "") { $resumenRejilla }
}
""

if ($cargado) {
  "CLIENTE CARGADO: $interDespues"
} elseif ($filas -gt 0) {
  "VARIOS RESULTADOS: $filas. Azul no carga al cliente solo; hay que elegir la cuenta."
  if ($NoEntrar) {
    "Se para aqui por -NoEntrar."
  } else {
    # Se elige la PRIMERA fila, como indico Dorian. Primero por el puente; si no marca, con un
    # clic real sobre la celda. La prueba de que quedo marcada no es la llamada: es que Azul
    # habilite 'Seleccionar', que esta apagado mientras no hay fila elegida.
    $s = [Busca]::SeleccionarFila(0)
    if ($s -like 'ERROR*') { $s; exit 1 }
    if ($s -eq 'NO') {
      "El puente no marco la fila; se pulsa de verdad."
      $c = Invoke-ClicGeo -Geometria { [Busca]::GeoCelda(0, 1) } -Que "la primera fila"
      if ($c -like 'ERROR*') { $c; exit 1 }
      Start-Sleep -Milliseconds 700
    }
    "Primera fila marcada."

    # Comprobacion contra la lista de ordenes: la fila elegida tiene que hablar del cliente
    # pedido. Es barata y ataja el error mas caro de todos, que es entrar al cliente
    # equivocado sin que nada lo delate hasta que ya esta escrito en el documento.
    if ($razon -ne "") {
      $texto = ([Busca]::FilaTexto(0)).ToUpper()
      # A la palabra que sirve de clave se le quitan los signos de los extremos. Sin esto, una
      # razon social como "EJEMPLO, ASESORIA E INTEGRADORES" deja la clave en
      # 'EJEMPLO,' con la coma pegada, y Azul escribe la misma empresa sin coma: la guarda
      # rechazaba a un cliente correcto. Se recortan solo los extremos, no el interior, para no
      # romper palabras que llevan guion dentro.
      $signos = [char[]]@(',', '.', ';', ':', '"', "'", '(', ')', '[', ']', '-', '/')
      $clave = (($razon -split '\s+') |
        ForEach-Object { $_.Trim($signos) } |
        Where-Object { $_.Length -ge 4 } |
        Select-Object -First 1)
      if ($clave -and ($texto -notlike "*$($clave.ToUpper())*")) {
        "ERROR: la primera fila no menciona '$clave', que es la razon social de la lista."
        "  fila: $($texto.Trim())"
        "  No se entra a un cliente que no coincide. Usa -Forzar si de verdad es el correcto."
        if (-not $Forzar) { exit 1 }
        "Se continua por -Forzar."
      } else {
        "La fila coincide con la razon social de la lista ($razon)."
      }
    } else {
      # La guarda mas fuerte del sistema es CONDICIONAL, y hasta hoy su ausencia era muda.
      # Con el libro de Excel vivo no se notaba, porque la razon social siempre venia; con un
      # CSV exportado a mano, una columna mal elegida la deja vacia y la corrida seguia
      # igual de callada, sin nada que impidiera entrar al cliente equivocado.
      "AVISO: sin razon social, no se puede comprobar que la fila sea del cliente pedido."
      "       La guarda contra entrar al cliente equivocado NO se aplica en esta corrida."
    }

    $c = Invoke-ClicGeo -Geometria { [Busca]::GeoBotonResultados('Seleccionar') } -Que "'Seleccionar'"
    if ($c -like 'ERROR*') { $c; exit 1 }
    "'Seleccionar' pulsado con clic real en $(($c -split '\|')[1])."

    # Aqui lo unico que se espera es el marco de interaccion. La rejilla ya no es un desenlace:
    # es lo que se acaba de dejar atras, y esperar a que la pantalla entera se asiente era pagar
    # segundos por una senal que no dice nada mas que esta.
    $e = [Busca]::EsperarInteraccion($interAntes, 90)
    if ((($e -split '\|')[0]) -eq 'NADA') {
      "ERROR: 90 s tras 'Seleccionar' y no aparecio ningun marco de interaccion nuevo."
      "Captura: $(Save-Prueba -Etiqueta 'seleccionar_sin_efecto')"
      exit 1
    }
    $interDespues = (($e -split '\|', 2)[1])
    $cargado = ($interDespues -ne "" -and $interDespues -ne $interAntes)
        if ($cargado) { "CLIENTE CARGADO: $interDespues" }
    else { "AVISO: se pulso 'Seleccionar' pero no aparecio marco de interaccion nuevo." }
  }
} else {
  "SIN RESULTADO. El formulario se cerro pero Azul no cargo ningun cliente ni trajo filas."
  "Con 'ID de CF' = $Cuenta no hubo coincidencia. Comprueba el numero en el Excel, o si"
  "esa cuenta va en otro campo del formulario."
  exit 1
}

# ---- 6. abrir la pestana Suscripciones --------------------------------------
# Ultimo paso: con el cliente dentro, la pestana 'Suscripciones' ensena sus lineas. Hay
# clientes que no piden elegir cuenta y se llega aqui directo; por eso este bloque no depende
# de si hubo rejilla o no, solo de que el cliente este cargado.
if (-not $cargado) {
  ""
  "No se abre 'Suscripciones' porque no hay cliente cargado."
  exit 1
}

# ---- TOPE DE TRAMO (Etapa 5.4, 08/09/2026) ----------------------------------
# Reloj de pared que envuelve TODO este tramo: desde que la busqueda confirma una
# sola cuenta (aqui arriba, $cargado ya en true) hasta que la tabla de Suscripciones esta
# lista para leer. Capa NUEVA por fuera: no reemplaza ni recorta Init() (5 s) ni el margen
# de CommitCF -- esos ya vencieron mucho antes de llegar aqui y siguen intactos; son los que
# atajan un puente lento o un commit que no prendio, un problema distinto al de este tramo.
# Igual que el resto del archivo, el reloj se SONDEA entre pasos -- no puede cortar una
# llamada JAB que ya este en curso, la misma limitacion que ya tienen todos los demas topes
# de este archivo (docs/trampas.md, trampa 5: ninguna espera fija sirve, y ninguna espera,
# fija o no, puede interrumpir una llamada bloqueada a medias).
# Si se pasa del tope: NO se asume que la tabla esta lista. Se para, no se lee nada, se deja
# para revisar a mano -- mismo idioma que cualquier otra guarda de este archivo (ERROR +
# captura + exit 1). $topeTramo tiene que quedar asignado ANTES de que se llame a cualquiera
# de las dos funciones de abajo.
#
# 20 s y no 10 (Dorian, 11/09/2026). Los 10 se pusieron cuando se creia que lo que se estaba
# atajando era Azul tardando en contestar. MEDIDO el 11/09 sobre un cliente que salio BIEN,
# a mano quieta y con el cache de pestana ya puesto: este tramo gasto 40.604 llamadas JAB y
# 9 segundos. Es decir que el tope viejo dejaba UN segundo de margen a un cliente sano, y de
# ahi salian los tres abortos del historial. El gasto no es Azul contestando: es el puente
# recorriendo el arbol entero, y lo sigue haciendo el bucle de abajo en CADA vuelta.
# El numero esta en una sola variable a proposito: antes los mensajes repetian "10 s" en seis
# sitios y cambiar el tope los dejaba mintiendo.
$TRAMO_SEG = 20
$topeTramo = (Get-Date).AddSeconds($TRAMO_SEG)
function Test-TramoVencido { (Get-Date) -ge $topeTramo }
function Get-TramoRestanteSeg { [Math]::Max(0, [int][Math]::Ceiling(($topeTramo - (Get-Date)).TotalSeconds)) }

""
# Ahora se llega aqui EN CUANTO Azul abre el marco de interaccion, que puede ser un pelo antes
# de que sus pestanas esten colocadas. Se espera a que la pestana exista y ADEMAS este quieta:
# pulsarla mientras la tira de pestanas todavia se mueve manda el clic a la posicion vieja.
if (-not [Busca]::EsperarPestanaExiste('Suscripciones', (Get-TramoRestanteSeg))) {
  "ERROR: la pestana 'Suscripciones' no llego a quedarse quieta dentro del tope de tramo de $TRAMO_SEG s."
  "Captura: $(Save-Prueba -Etiqueta 'pestana_sin_aparecer')"
  [Busca]::Nota("tramo busqueda->Suscripciones: EXCEDIDO tope de $TRAMO_SEG s (la pestana no aparecio)")
  exit 1
}
if ([Busca]::PestanaSeleccionada('Suscripciones')) {
  "La pestana 'Suscripciones' ya estaba abierta."
} else {
  # Donde estaba la pestana justo antes de pulsarla, para poder comparar si el clic no surte
  # efecto: si despues esta en otro sitio, es que se movio y el clic salio a la posicion vieja.
  $geoAntes = [Busca]::GeoPestana('Suscripciones')
  $c = Invoke-ClicGeo -Geometria { [Busca]::GeoPestana('Suscripciones') } -Que "la pestana 'Suscripciones'"
  if ($c -like 'ERROR*') { $c; exit 1 }
  "Pestana 'Suscripciones' pulsada con clic real en $(($c -split '\|')[1])."
  # La confirmacion la da Azul: la pestana pasa al estado 'selected'. Se pide como maximo lo
  # que quede del tope de tramo (Get-TramoRestanteSeg), nunca mas -- por eso ya no se
  # pide un "40" fijo aqui: ese tope propio de EsperarPestana() sigue existiendo tal cual en
  # su codigo, pero en este punto del archivo el tramo es siempre el mas corto de
  # los dos, asi que es el que manda.
  if (-not [Busca]::EsperarPestana('Suscripciones', [Math]::Min(6, (Get-TramoRestanteSeg)))) {
    # Un solo reintento, y SOLO si la pestana se ha movido desde donde se pulso: eso prueba que
    # el clic salio a la posicion vieja. Si no se ha movido, el clic llego bien y el problema es
    # otro, asi que no se vuelve a pulsar. La geometria se relee, de modo que el segundo clic va
    # otra vez a 'Suscripciones' y a ningun otro sitio.
    $geoAhora = [Busca]::GeoPestana('Suscripciones')
    if ($geoAhora -like 'OK|*' -and $geoAhora -ne $geoAntes) {
      "La pestana se habia movido de $(($geoAntes -split '\|')[1]) a $(($geoAhora -split '\|')[1]); se vuelve a pulsar."
      $c2 = Invoke-ClicGeo -Geometria { [Busca]::GeoPestana('Suscripciones') } -Que "la pestana 'Suscripciones' (reintento)"
      if ($c2 -like 'OK|*') { "Reintento pulsado en $(($c2 -split '\|')[1])." }
    }
    if (-not [Busca]::EsperarPestana('Suscripciones', (Get-TramoRestanteSeg))) {
      "ERROR: la pestana 'Suscripciones' no quedo seleccionada dentro del tope de tramo de $TRAMO_SEG s."
      "Captura: $(Save-Prueba -Etiqueta 'pestana_sin_efecto')"
      [Busca]::Nota("tramo busqueda->Suscripciones: EXCEDIDO tope de $TRAMO_SEG s (pestana)")
      exit 1
    }
  }
}

# Las lineas tardan en pintarse despues de que la pestana se selecciona: se espera a que la
# tabla aparezca y DEJE DE CRECER, igual que hace azul_fast.ps1 con el arbol de atributos.
$resumen = ""; $previo = -1; $estables = 0
for ($w = 0; $w -lt 60; $w++) {
  # Primera linea del cuerpo: el tope de tramo tambien corta este bucle, no solo el de 60
  # vueltas que ya habia. 60 vueltas de 400 ms son 24 s mas lo que cueste cada llamada, asi
  # que con el tramo en 20 s este 'break' sigue siendo el que sale primero cuando haga falta.
  # OJO: ResumenSuscripciones() recorre el arbol entero en CADA vuelta y sin cache, asi que
  # este bucle no es de espera barata -- es, junto con la busqueda de la pestana, lo que se
  # come el tramo. Es el siguiente candidato a arreglar, igual que se arreglo la pestana.
  if (Test-TramoVencido) { break }
  $resumen = [Busca]::ResumenSuscripciones(8)
  if ($resumen -ne "" -and ($resumen -match 'Suscripciones: (\d+) filas')) {
    $n = [int]$Matches[1]
    if ($n -eq $previo) { $estables++; if ($estables -ge 3) { break } } else { $estables = 0; $previo = $n }
  }
  [Busca]::Espera(400)
}

# Se comprueba el tope SOLO si no hubo estabilizacion real ($estables -lt 3). Sin esa
# condicion, una tabla que SI estabilizo a tiempo pero cruzo el reloj durante la ultima
# llamada a ResumenSuscripciones() (que no tiene tope propio y recorre el arbol entero sin
# cache) se descartaria igual -- justo el caso de cliente lento-pero-real que mas importa
# dejar pasar. Con esta condicion, una estabilizacion de verdad siempre gana.
if ($estables -lt 3 -and (Test-TramoVencido)) {
  "ERROR: el tramo busqueda->Suscripciones paso de $TRAMO_SEG s. No se da por lista la tabla:"
  "       se marca para revisar a mano en vez de adivinar."
  "Captura: $(Save-Prueba -Etiqueta 'tramo_tabla')"
  [Busca]::Nota("tramo busqueda->Suscripciones: EXCEDIDO tope de $TRAMO_SEG s (tabla)")
  exit 1
}

""
if ($resumen -eq "") {
  "AVISO: la pestana esta abierta pero no encuentro la tabla de Suscripciones."
  "Mira la captura: puede que este cliente no tenga lineas, o que haya que bajar el scroll."
} else {
  $resumen.TrimEnd()
}
""
"ALTO. Cliente cargado y pestana de Suscripciones abierta."
"No se leyeron los detalles de cada linea y no se toco Word."
# Total de la corrida. Se compara contra otra corrida del mismo cliente para ver si un
# cambio de rendimiento sirvio, sin depender de lo cargado que estuviera Azul.
"Llamadas JAB de arbol: $([Busca]::Llamadas())"
[Busca]::Nota("buscar_cuenta: JAB total = " + [Busca]::Llamadas())
