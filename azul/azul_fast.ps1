param(
  # Numero de cuenta y razon social, tal como vienen en la lista de ordenes. NO deciden que se
  # lee -- eso lo gobierna lo que Azul tenga cargado -- sino que quedan en la cabecera del CSV
  # para que el archivo diga DE QUIEN es.
  #
  # Hasta el 05/09/2026 ese dato salia del titulo del panel del cliente, que se abria para la
  # foto. Sin foto no se abre ese panel, asi que sin esto el CSV en disco quedaria anonimo, y
  # con el se queda sin poder reintentar el volcado a Word a mano ni auditar una corrida vieja.
  [string]$Cuenta = "",
  [string]$Razon = "",
  # No escribir en Word al terminar. Por defecto SI se escribe: Word siempre esta abierto y
  # el resultado siempre va al final, asi que el segundo comando sobraba.
  [switch]$SinWord,
  # Pausas ADICIONALES para no pedirle datos a Azul mas rapido de lo que aguanta una persona
  # (Dorian, 07/09/2026: a mano aguanta horas, con el script se caia en minutos). Se SUMAN a
  # las esperas que ya habia; ninguna las sustituye. En segundos.
  [int]$PausaLineas = 4,   # entre cerrar el detalle de una linea y seleccionar la siguiente
  [int]$PausaLectura = 7,  # entre abrir el detalle de una linea y empezar a leerlo
  # Donde esta el enlace "Cliente:" de la banda de arriba, en coordenadas RELATIVAS a la
  # esquina superior izquierda de la ventana de Azul. Va como parametro y no clavado porque
  # ese enlace es contenido de un navegador incrustado y el puente no lo ve, asi que no hay
  # forma de preguntarle donde esta.
  #
  # 500 y no 529, medido el 04/09/2026 sobre dos capturas con cliente cargado: el enlace
  # ARRANCA siempre en x~491 y se extiende hacia la derecha segun el largo del nombre. Con
  # 'CLIENTE EJE...' llegaba hasta x~568 y 529 caia dentro; con 'ACME' termina en x~520 y 529
  # caia 9 px FUERA, sobre gris. Por eso la foto se perdia con los clientes de nombre corto.
  # 500 cae dentro del primer caracter del nombre sea cual sea su largo.
  [int]$ClienteX = 500,
  [int]$ClienteY = 120
)
$ErrorActionPreference='Stop'
# NOTA: este archivo debe permanecer en ASCII puro.
# PowerShell 5.1 lee .ps1 sin BOM como ANSI y rompe los acentos, lo que hacia
# que comparaciones como "Plan Movil" nunca coincidieran. Usar escapes \uXXXX.
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

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
  public int accessibleSelection; public int accessibleText; public int accessibleInterfaces;
}
[StructLayout(LayoutKind.Sequential)]
public struct ATI { public long caption; public long summary; public int rowCount; public int columnCount; public long ctx; public long tbl; }
[StructLayout(LayoutKind.Sequential)]
public struct MSG { public IntPtr hwnd; public uint message; public IntPtr wParam; public IntPtr lParam; public uint time; public int px; public int py; }

public static class Azul {
  const string D="WindowsAccessBridge-32.dll";
  const int NB=512, MAXA=256, MAXTD=32;

  // Etiquetas de Azul con acentos, en escapes para sobrevivir la codificacion
  const string L_PLAN = "Plan M\u00F3vil";   // "Plan Movil" con o acentuada
  const string L_COMP = "Compromiso";
  const string L_MPE  = "MPE";
  const string L_FF   = "Fecha Final del";
  const string L_DUR  = "Duraci\u00F3n del";  // "Duraci\u00F3n del Compromiso" = Plan Forzoso
  const string L_MOD  = "Modelo";
  const string L_MAR  = "Marca";
  const string MULT   = "m\u00FAltiples ";        // multiples

  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void Windows_run();
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleContextFromHWND(IntPtr h,out int vm,out long ac);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleContextInfo(int vm,long ac,out ACI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern long getAccessibleChildFromContext(int vm,long ac,int idx);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleTableInfo(int vm,long ac,out ATI i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool getAccessibleActions(int vm,long ac,IntPtr a);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern bool doAccessibleActions(int vm,long ac,IntPtr a,out int f);
  // Seleccion por el modelo real de la tabla. El indice es de CELDA, no de fila:
  // para la fila r hay que pasar r*columnas. Esto si dispara los listeners de Azul.
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void addAccessibleSelectionFromContext(int vm,long ac,int i);
  [DllImport(D,CallingConvention=CallingConvention.Cdecl)] static extern void clearAccessibleSelectionFromContext(int vm,long ac);
  [DllImport("user32.dll")] static extern bool PeekMessage(out MSG m,IntPtr h,uint a,uint b,uint c);
  [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG m);
  [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG m);
  // Front() y RealClick() vivian aqui sin usarse, del enfoque viejo por capturas y clics.
  // Ahora que si hacen falta (para la banda de "Cliente:", que el puente no ve) se mudaron
  // a captura.ps1, que es donde vive todo lo que toca la maquina. Este archivo se queda
  // solo con el puente: lo que unicamente lee.
  static bool Habilitado(long ac){
    ACI i; if(!Inf(ac,out i)) return false;
    return i.states_en_US!=null && i.states_en_US.IndexOf("enabled")>=0;
  }

  static int vm; static long root;
  static StringBuilder log=new StringBuilder();
  static void L(string s){ log.AppendLine(s); }

  // progreso en vivo a disco: sin esto la corrida es una caja negra
  public static string PROG="";   /* la fija PowerShell segun donde este el script */
  // Pausas ADICIONALES (ms), fijadas desde PowerShell tras el Add-Type. Se suman a las
  // esperas que ya habia en el bucle de lectura; no sustituyen ninguna.
  public static int PAUSA_LINEAS=4000;
  public static int PAUSA_LECTURA=7000;
  static void P(string s){
    try{ System.IO.File.AppendAllText(PROG, DateTime.Now.ToString("HH:mm:ss")+"  "+s+"\r\n"); }catch{}
  }

  static void Pump(int ms){
    var e=DateTime.Now.AddMilliseconds(ms); MSG m;
    while(DateTime.Now<e){
      while(PeekMessage(out m,IntPtr.Zero,0,0,1)){ TranslateMessage(ref m); DispatchMessage(ref m); }
      System.Threading.Thread.Sleep(3);
    }
  }
  // Contador de llamadas al puente. Inf y Kid son los dos puntos por los que pasa TODO
  // recorrido del arbol, asi que contarlos aqui mide el costo real sin instrumentar nada mas.
  // Es un incremento de un long: no cambia el comportamiento y no pesa de forma medible.
  // Sirve para comparar una corrida contra otra sin cronometro: el reloj depende de lo
  // cargado que este Azul, el numero de llamadas no.
  static long jab=0;
  public static long Llamadas(){ return jab; }
  static bool Inf(long ac,out ACI i){ jab++; return getAccessibleContextInfo(vm,ac,out i); }
  static long Kid(long ac,int k){ jab++; return getAccessibleChildFromContext(vm,ac,k); }

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

  static List<long> Find(long start,Func<ACI,long,bool> pred,int cap,bool firstOnly){
    var res=new List<long>(); var st=new Stack<long>(); st.Push(start); int n=0; ACI i;
    while(st.Count>0 && n<cap){
      long ac=st.Pop(); n++;
      if(!Inf(ac,out i)) continue;
      if(pred(i,ac)){ res.Add(ac); if(firstOnly) return res; }
      if(i.role_en_US=="table") continue;
      for(int k=i.childrenCount-1;k>=0;k--){ long c=Kid(ac,k); if(c!=0) st.Push(c); }
    }
    return res;
  }
  static long First(long s,Func<ACI,long,bool> p){ var r=Find(s,p,60000,true); return r.Count>0?r[0]:0; }
  static string Cel(long tbl,int idx){
    long c=Kid(tbl,idx); if(c==0) return "";
    ACI i; if(!Inf(c,out i)) return "";
    return i.name==null?"":i.name.Trim();
  }

  // En el arbol de atributos la columna 0 son nodos "tree" cuyo texto NO viene en
  // name sino en description, envuelto en HTML: <html>Fecha Final del Compromiso</html>
  static string Strip(string s){
    if(s==null) return "";
    var sb=new StringBuilder(); bool tag=false;
    foreach(char c in s){
      if(c=='<') tag=true;
      else if(c=='>') tag=false;
      else if(!tag) sb.Append(c);
    }
    return sb.ToString().Replace("&nbsp;"," ").Trim();
  }
  // texto de una celda: primero name, y si viene vacio se usa description sin HTML
  static string Txt(long tbl,int idx){
    long c=Kid(tbl,idx); if(c==0) return "";
    ACI i; if(!Inf(c,out i)) return "";
    string n=(i.name==null?"":i.name.Trim());
    if(n.Length>0) return n;
    return Strip(i.description);
  }
  // Los marcos internos son hijos DIRECTOS de un desktop pane. Cacheamos los panes
  // una vez: asi cada sondeo cuesta ~10 llamadas JAB en vez de recorrer todo el arbol.
  // Esto importa porque cada llamada JAB es IPC sincrono y se bloquea mientras Azul carga.
  static List<long> panes = new List<long>();
  const string DET_PREFIX = "Detalles del Producto Asignado";

  static long DetFrame(){
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.StartsWith(DET_PREFIX)) return c;
      }
    }
    return 0;
  }
  // recorrido completo: solo como red de seguridad, nunca dentro de un bucle de espera
  static long DetFrameFull(){
    return First(root,(i,ac)=> i.role_en_US=="internal frame" && i.name!=null && i.name.StartsWith(DET_PREFIX));
  }
  static bool CerrarDetalle(){
    long d=DetFrame(); if(d==0) d=DetFrameFull(); if(d==0) return false;
    long b=First(d,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()=="Cerrar");
    if(b==0) return false;
    Click(b);
    // Sondeo en vez de esperar 700 ms fijos: en cuanto el marco desaparece, seguimos.
    // DetFrame() solo recorre los desktop panes cacheados, asi que preguntar es barato.
    // El presupuesto de tiempo es el mismo de antes; lo normal es salir a los ~160 ms.
    for(int w=0; w<9 && DetFrame()!=0; w++) Pump(80);
    return true;
  }
  // ---- panel de la vista general del cliente ----
  // El sistema ya NO lo abre ni lo fotografia: desde el 05/09/2026 esa captura la hace Dorian
  // a mano y la pega en el Word. Lo que queda aqui es lo unico que sigue haciendo falta:
  // SABER si hay uno abierto y poder CERRARLO.
  //
  // Hace falta porque ese marco reemplaza al de interaccion y deja la tabla de Suscripciones
  // fuera del alcance del puente (docs/mapa-datos.md seccion 5). Antes lo abria y lo cerraba
  // el propio sistema; ahora lo abre una persona, que es mucho mas probable que lo deje
  // abierto. Ver Close-PanelClienteAbierto, mas abajo.
  //
  // El titulo lo gobierna todo: mientras carga el marco se llama "Formulario" y solo cuando
  // termina pasa a "Cuenta: <nombre>". De ahi el prefijo.
  const string CTA_PREFIX = "Cuenta: ";

  // Solo por los desktop panes cacheados, sin red de seguridad por el arbol entero.
  // El caso NORMAL es que no haya panel abierto, y el recorrido completo es caro justo en ese
  // caso: First() visita el arbol entero antes de poder decir que no esta. Con la comprobacion
  // ahora en el arranque de cada cliente, esa red se pagaria en TODAS las corridas para
  // contestar "no hay". Se prefiere no verlo en el caso raro -- y entonces la corrida falla
  // como fallaba antes de existir esta guarda -- a pagarlo en todas.
  static long MarcoCuentaPanes(){
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.StartsWith(CTA_PREFIX)) return c;
      }
    }
    return 0;
  }
  public static string MarcoCuenta(){
    long f=MarcoCuentaPanes(); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.name==null?"":i.name;
  }
  // "x,y,w,h" del marco. La X de su barra de titulo se calcula de aqui en vez de clavarla:
  // el puente no expone los botones de la barra, pero si la geometria del marco.
  public static string GeoMarcoCuenta(){
    long f=MarcoCuentaPanes(); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.x+","+i.y+","+i.width+","+i.height;
  }
  // Espera al TITULO, no al reloj: mientras carga, el marco se llama "Formulario" y se ve
  // como un formulario en blanco con boton "Crear". Solo cuando termina de cargar el titulo
  // pasa a "Cuenta: <nombre>".
  public static string EsperarMarcoCuenta(int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      long f=MarcoCuentaPanes();
      if(f!=0){ ACI i; if(Inf(f,out i) && i.name!=null && i.name.Length>0) return i.name; }
      Pump(400);
    }
    return MarcoCuenta();
  }
  public static bool EsperarSinMarcoCuenta(int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      if(MarcoCuenta().Length==0) return true;
      Pump(400);
    }
    return MarcoCuenta().Length==0;
  }
  // ---- marco de interaccion ----
  // REGLAS.md Ciclo 3, 09/09/2026. Mismo mecanismo que el panel del cliente: el puente ve
  // el marco y su geometria, pero no los botones de su barra de titulo, asi que la X se
  // calcula igual (x+ancho-21, y+9) en vez de clavarla.
  // Acentuada por escape, no literal -- ver la nota de ASCII puro al principio del archivo.
  // Mismo literal que F_INT en buscar_cuenta.ps1, ya probado contra Azul real el 09/09/2026.
  const string INT_PREFIX = "Inicio de Interacci\u00F3n";
  static long MarcoInteraccionPanes(){
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.StartsWith(INT_PREFIX)) return c;
      }
    }
    return 0;
  }
  public static string MarcoInteraccion(){
    long f=MarcoInteraccionPanes(); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.name==null?"":i.name;
  }
  public static string GeoMarcoInteraccion(){
    long f=MarcoInteraccionPanes(); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.x+","+i.y+","+i.width+","+i.height;
  }
  public static bool EsperarSinMarcoInteraccion(int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      if(MarcoInteraccion().Length==0) return true;
      Pump(400);
    }
    return MarcoInteraccion().Length==0;
  }
  // Espera bombeando mensajes, no con Start-Sleep: el puente necesita que el hilo despache.
  public static void Espera(int ms){ Pump(ms); }
  // Deja que PowerShell escriba en progreso.txt por el mismo canal que el resto de la corrida.
  public static void Nota(string s){ P(s); }

  static string Csv(string s){
    if(s==null) s="";
    if(s.IndexOf('"')>=0) s=s.Replace("\"","\"\"");
    return "\""+s+"\"";
  }

  // Arranque del puente y cacheo de los desktop panes. Se saco de Run() para que la captura
  // de la vista del cliente pueda usarse sola, sin releer todas las lineas de la cuenta.
  public static bool Init(IntPtr hwnd){
    Windows_run();
    // Sondeo por condicion, no espera por reloj: se pregunta por lo que de verdad hace falta,
    // que es que el puente conteste con el contexto raiz de la ventana.
    //
    // 5 s y no 800 ms (05/09/2026). El tope es techo, no coste: el bucle sale en cuanto el
    // puente contesta, asi que con Azul sano no se paga nada. Se sube por el Azul DEGRADADO,
    // que es el caso que importa: el 04/09/2026 a las 13:24, con Azul vivo pero quemando un
    // nucleo al 100%, un cliente murio aqui con "sin contexto raiz. El puente no ve a Azul".
    // Con 800 ms, un puente lento no es un puente ausente y se trataba igual.
    DateTime tope=DateTime.Now.AddMilliseconds(5000);
    while(!getAccessibleContextFromHWND(hwnd,out vm,out root) || root==0){
      if(DateTime.Now>=tope) return false;
      Pump(50);
    }
    // cachear los desktop panes ANTES de cualquier espera
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
    // Un cache VACIO no falla, y ese es el problema: Init devuelve true, la corrida sigue, y
    // DetFrame() se queda sin donde mirar -- todo cae al recorrido completo del arbol, en cada
    // linea, que es exactamente lo que el cache existe para evitar. Desde el 05/09/2026
    // MarcoCuentaPanes() ya no tiene red por el arbol entero, asi que la guarda del panel de
    // cliente abierto tampoco veria nada. No se aborta (la lectura puede salir bien, solo que
    // lenta), pero deja de ser silencioso.
    if(panes.Count==0) P("AVISO: cache de panes VACIO. La corrida va a recorrer el arbol entero en cada busqueda de marco, y la guarda del panel de cliente no vera nada.");
    return true;
  }

  public static string Run(IntPtr hwnd){
    log.Clear();
    var sw=System.Diagnostics.Stopwatch.StartNew();
    if(!Init(hwnd)) return "ERROR: sin contexto raiz";

    // un detalle abierto tapa la tabla de Suscripciones en el arbol
    CerrarDetalle();

    long sub=0; int SR=0,SC=0;
    var tablas=Find(root,(i,ac)=> i.role_en_US=="table",60000,false);
    P("tablas encontradas: "+tablas.Count);
    foreach(long t in tablas){
      ATI ti;
      if(!getAccessibleTableInfo(vm,t,out ti)){ P("   tabla sin info"); continue; }
      long c0=Kid(t,0);
      string r0="(sin celda0)";
      if(c0!=0){ ACI i0t; if(Inf(c0,out i0t)) r0=i0t.role_en_US; }
      P("   tabla "+ti.rowCount+"x"+ti.columnCount+" celda0="+r0);
      if(ti.columnCount<10||ti.rowCount<1) continue;
      if(c0==0) continue;
      ACI i0; if(!Inf(c0,out i0)) continue;
      if(i0.role_en_US=="radio button"){ sub=t; SR=ti.rowCount; SC=ti.columnCount; break; }
    }
    if(sub==0){ P("ERROR: sin tabla de Suscripciones"); return "ERROR: no encuentro la tabla de Suscripciones. Deja Azul en la pestana Suscripciones."; }
    P("tabla Suscripciones "+SR+"x"+SC);

    long btn=First(root,(i,ac)=> i.role_en_US=="push button" && i.name!=null && i.name.Trim()=="Ver Productos Asignados");
    if(btn==0) return "ERROR: no encuentro el boton 'Ver Productos Asignados'.";

    // ---- estado de la linea: columna 7 de la tabla de Suscripciones (Activa/Cancelado) ----
    // Las canceladas no se consideran. Se descartan AQUI, antes de abrir nada, asi que
    // ni se les toca la subventana de detalle; se listan aparte para poder cuadrar.
    // Un valor que no sea Activa ni Cancelado NO se excluye: excluir por algo que no
    // entiendo romperia el cuadre. Entra a la corrida y sale como REVISAR (mas abajo).
    var rows=new List<int>(); var nums=new List<string>(); var els=new List<string>();
    var canceladas=new List<string>();
    int filasConNumero=0;
    for(int r=0;r<SR;r++){
      string num=Cel(sub,r*SC+1);
      bool dig=num.Length>=6; foreach(char ch in num) if(!char.IsDigit(ch)) dig=false;
      if(!dig) continue;
      filasConNumero++;
      string el=(SC>7)?Cel(sub,r*SC+7):"";
      if(el.StartsWith("Cancel",StringComparison.OrdinalIgnoreCase)){ canceladas.Add(num); continue; }
      rows.Add(r); nums.Add(num); els.Add(el);
    }
    P("filas con numero: "+filasConNumero+"  canceladas: "+canceladas.Count+"  a leer: "+rows.Count);

    // tol: la ventana no es un muro. Hasta 9 dias pasada la frontera la linea sigue
    // siendo renovable; a los 10 ya no. Decision del usuario, 02/09/2026.
    DateTime hoy=DateTime.Today, lim=hoy.AddMonths(3), tol=lim.AddDays(10), unAnio=hoy.AddYears(1);
    var filasSi=new List<string>(); var filasNo=new List<string>(); var filasRev=new List<string>();
    int ySi=0,zNo=0,wRev=0,sinForz=0;
    bool abortado=false;
    bool primeraSeleccion=true;

    try{
      for(int n=0;n<rows.Count;n++){
        string plan="",disp="",mpe="",ffin="",forz="",est="REVISAR",nota=""; bool esSim=false;
        // Un reintento por linea: los fallos pasajeros (el detalle no abre, los
        // atributos no cargan) suelen recuperarse al segundo intento. Los fallos
        // de datos (ambiguedad, sin Compromiso) NO se reintentan: son deterministas.
        for(int intento=1;intento<=2;intento++){
          // Pausa ADICIONAL entre selecciones (Dorian, 07/09/2026). Antes de fijar 'tope':
          // no le resta presupuesto al tope de 40 s que ya existia, se lo suma por delante.
          if(!primeraSeleccion) Pump(PAUSA_LINEAS);
          primeraSeleccion=false;
          plan=""; disp=""; mpe=""; ffin=""; forz=""; est="REVISAR"; nota=""; esSim=false;
          bool transitorio=false;
          // Tope de reloj por intento. Los contadores de vueltas no acotan nada cuando
          // Azul agoniza: cada llamada JAB es IPC sincrono y se queda colgada segundos.
          // Sin esto, una sola linea se comio 135 s de una corrida de 295 s (02/09/2026).
          DateTime tope=DateTime.Now.AddSeconds(40);
          bool porTiempo=false;
          try{
            // guarda: si la cuenta cambio bajo nuestros pies, abortar en vez de girar en vano
            if(Cel(sub,rows[n]*SC+1)!=nums[n]){
              nota="abortado: la cuenta cambio durante la corrida"; abortado=true; goto finIntento;
            }
            P("-> linea "+nums[n]+" (fila "+rows[n]+")"+(intento>1?"  [reintento]":""));

            // Seleccionar por el modelo real de la tabla (indice de CELDA = fila*columnas).
            // doAccessibleActions sobre el radio solo cambia el estado visual del widget
            // y NO habilita "Ver Productos Asignados"; esta via si.
            clearAccessibleSelectionFromContext(vm,sub);
            Pump(80);
            addAccessibleSelectionFromContext(vm,sub,rows[n]*SC);

            // Sondeo en vez de esperar 950 ms fijos. No se adivina el momento: se pregunta
            // por la MISMA condicion que hay que comprobar de todos modos -- que quedo
            // marcada la fila pedida y que el boton se habilito. Sondear una condicion es
            // mas seguro que una espera fija, porque la fija tambien puede quedarse corta y
            // encima no verifica nada. Lo normal es salir a los ~160 ms. (La estabilidad del
            // arbol de atributos, mas abajo, NO se toca: ahi no hay condicion fiable que
            // preguntar y por eso se espera a que deje de crecer.)
            // El estado del boton se guarda al salir del bucle en vez de volver a
            // preguntarlo dos veces mas abajo: son la misma respuesta y ya se tenia.
            //
            // 2500 ms y no 950 (05/09/2026). Es un techo, no un coste: con Azul sano se sale a
            // los ~160 ms y no cambia nada. Se sube porque agotarlo tiene DOS salidas malas, y
            // la segunda es peor que perder tiempo:
            //
            //   - sin 'marcada' -> "no se marco la fila esperada", se marca transitorio y se
            //     REINTENTA la linea entera: Pump(1200) mas volver a seleccionar mas otro
            //     tope. Un tope corto sale mas caro que uno largo.
            //   - con 'marcada' pero sin 'habilitado' -> "boton deshabilitado", y eso NO es
            //     transitorio: no se reintenta. La linea acaba en A REVISAR con un mensaje que
            //     manda a dar un clic manual que no hacia falta. Con Azul cargado, eso es un
            //     REVISAR FALSO sobre una linea buena -- y una linea que se pierde en silencio
            //     es justo lo que este proyecto no se permite.
            bool marcada=false, habilitado=false;
            DateTime topeSel=DateTime.Now.AddMilliseconds(2500);
            while(true){
              long r2=Kid(sub,rows[n]*SC);
              ACI ri;
              marcada = r2!=0 && Inf(r2,out ri)
                        && ri.states_en_US!=null && ri.states_en_US.IndexOf("checked")>=0;
              habilitado = Habilitado(btn);
              if(marcada && habilitado) break;
              if(DateTime.Now>=topeSel) break;
              Pump(80);
            }
            P("   seleccionada="+marcada+"  boton_enabled="+habilitado);
            if(!marcada){ nota="no se marco la fila esperada"; transitorio=true; goto finIntento; }
            if(!habilitado){ nota="boton deshabilitado: haz un clic manual en cualquier linea una vez y reintenta"; goto finIntento; }
            bool okB=Click(btn);
            P("   boton="+okB);
            if(!okB){ nota="no abrio Ver Productos Asignados"; transitorio=true; goto finIntento; }

            // Pausa ADICIONAL entre abrir el detalle y empezar a leerlo (Dorian, 07/09/2026).
            // Se le suma a 'tope' para no comerle presupuesto al sondeo de abajo: es aparte,
            // no en vez de.
            Pump(PAUSA_LECTURA);
            tope=tope.AddSeconds(PAUSA_LECTURA/1000.0);

            // Se pregunta ANTES de esperar: el bucle viejo dormia 500 ms de entrada aunque
            // el detalle ya estuviera ahi. Mismo presupuesto total (12 s), grano mas fino.
            long det=DetFrame();
            for(int w=0;w<80 && det==0 && DateTime.Now<tope;w++){ Pump(150); det=DetFrame(); }
            if(det==0) det=DetFrameFull();   // red de seguridad, una sola vez
            P("   detalle="+(det!=0));
            if(det==0){
              porTiempo=(DateTime.Now>=tope);
              nota=porTiempo?"se agoto el tiempo esperando el detalle":"el detalle no aparecio";
              transitorio=true; goto finIntento;
            }

            // ---- el detalle abierto debe ser el de ESTA linea ----
            // el titulo la identifica: "Detalles del Producto Asignado: Producto Movil (8910000001)"
            // sin esto se podria leer una subventana vieja y atribuir sus datos a otra linea
            ACI di;
            string titulo = (Inf(det,out di) && di.name!=null) ? di.name : "";
            for(int w=0;w<16 && titulo.IndexOf(nums[n])<0 && DateTime.Now<tope;w++){
              Pump(400);
              titulo = (Inf(det,out di) && di.name!=null) ? di.name : "";
            }
            P("   titulo="+titulo);
            if(titulo.IndexOf(nums[n])<0){
              porTiempo=(DateTime.Now>=tope);
              nota=porTiempo?"se agoto el tiempo esperando el titulo":"el detalle abierto no corresponde a esta linea";
              transitorio=true; CerrarDetalle(); goto finIntento;
            }

            long at=0; int AR=0,AC=0,prevAR=-1,estable=0;
            for(int w=0;w<60 && DateTime.Now<tope;w++){
              if(at==0){
                int bk=0;
                foreach(long d in Find(det,(i,ac)=> i.role_en_US=="table",20000,false)){
                  ATI t2; if(!getAccessibleTableInfo(vm,d,out t2)) continue;
                  int kk=t2.rowCount*t2.columnCount;
                  if(kk>bk){ bk=kk; at=d; AR=t2.rowCount; AC=t2.columnCount; }
                }
              } else { ATI t3; if(getAccessibleTableInfo(vm,at,out t3)){ AR=t3.rowCount; AC=t3.columnCount; } }
              // esperar a que el arbol DEJE de crecer: leerlo a medio cargar da datos incompletos
              if(at!=0 && AR>5 && AC>=3){
                if(AR==prevAR){ estable++; if(estable>=3) break; } else estable=0;
                prevAR=AR;
              }
              Pump(250);
            }
            P("   attr AR="+AR+" AC="+AC);
            if(at==0||AR<=0||AC<3){
              porTiempo=(DateTime.Now>=tope);
              nota=porTiempo?"se agoto el tiempo esperando los atributos":"los atributos no cargaron";
              transitorio=true; CerrarDetalle(); goto finIntento;
            }

            int maxR=Math.Min(AR,3000);
            string[] nm=new string[maxR];
            for(int r=0;r<maxR;r++) nm[r]=Txt(at,r*AC);

            // ---- unicidad: si hay ambiguedad NO se elige, se marca REVISAR ----
            int cPlan=0,cMPE=0,cFF=0,cMod=0,cDur=0;
            int rPlan=-1,rMPE=-1,rFF=-1,rDur=-1;
            for(int r=0;r<maxR;r++){
              string s2=nm[r];
              if(s2==L_PLAN){ cPlan++; if(rPlan<0) rPlan=r; }
              else if(s2==L_MPE){ cMPE++; if(rMPE<0) rMPE=r; }
              else if(s2.StartsWith(L_FF)){ cFF++; if(rFF<0) rFF=r; }
              else if(s2.StartsWith(L_DUR)){ cDur++; if(rDur<0) rDur=r; }
              else if(s2==L_MOD){ string v=Txt(at,r*AC+2); if(v!="N/A"&&v.Length>0) cMod++; }
            }
            if(cMPE>1){ nota=MULT+"MPE"; CerrarDetalle(); goto finIntento; }
            if(cFF>1){ nota=MULT+"Fecha Final del Compromiso"; CerrarDetalle(); goto finIntento; }
            if(cDur>1){ nota=MULT+"Duracion del Compromiso"; CerrarDetalle(); goto finIntento; }
            if(cPlan>1){ nota=MULT+L_PLAN; CerrarDetalle(); goto finIntento; }
            if(cMod>1){ nota=MULT+"dispositivos"; CerrarDetalle(); goto finIntento; }
            if(cPlan==0){ nota="sin nodo "+L_PLAN; CerrarDetalle(); goto finIntento; }

            plan=(AC>=4)?Txt(at,rPlan*AC+3):"";

            // ---- Compromiso debe ser hijo inmediato de Plan Movil ----
            int rComp=-1;
            for(int r=rPlan+1;r<Math.Min(rPlan+3,maxR);r++){ if(nm[r]==L_COMP){ rComp=r; break; } }

            // dispositivo: Marca + Modelo, leidos tal cual
            for(int r=0;r<maxR;r++){
              if(nm[r].StartsWith(L_MAR)){ string v=Txt(at,r*AC+2); if(v!="N/A"&&v.Length>0&&disp.Length==0) disp=v; }
              else if(nm[r]==L_MOD){ string v=Txt(at,r*AC+2); if(v!="N/A"&&v.Length>0) disp=(disp+" "+v).Trim(); }
            }
            if(disp.Length==0) disp="N/A";

            // Linea ACTIVA sin bloque Compromiso = SIM/eSIM sin contrato. No es una duda,
            // es un caso conocido y renovable. Se marca junto al numero para distinguirla,
            // sin agregar columna. Ojo: esto NO relaja las demas causas de REVISAR.
            if(rComp<0){ est="SI"; esSim=true; nota="SIM/eSIM sin contrato"; CerrarDetalle(); goto finIntento; }

            // ---- contencion: MPE y Fecha Final DENTRO del bloque Compromiso ----
            int fin=rComp+12;
            if(rMPE<0){ nota="sin MPE"; CerrarDetalle(); goto finIntento; }
            if(rFF<0){ nota="sin Fecha Final del Compromiso"; CerrarDetalle(); goto finIntento; }
            if(rMPE<=rComp||rMPE>fin){ nota="MPE fuera del bloque Compromiso"; CerrarDetalle(); goto finIntento; }
            if(rFF<=rComp||rFF>fin){ nota="Fecha Final fuera del bloque Compromiso"; CerrarDetalle(); goto finIntento; }

            // MPE se reporta tal cual lo entrega Azul, incluidos 0 y -1 (decision del usuario)
            mpe=Txt(at,rMPE*AC+2);
            ffin=Txt(at,rFF*AC+2);
            // Plan Forzoso (12/24/36 meses) = "Duracion del Compromiso", y solo si cae
            // DENTRO del bloque Compromiso, igual que MPE y Fecha Final. Si Azul no trae
            // la etiqueta se deja vacio y se cuenta al final: vacio no es adivinar.
            if(rDur>rComp && rDur<=fin) forz=Txt(at,rDur*AC+2);
            CerrarDetalle();

            if(ffin.Length==0){ nota="Fecha Final vacia"; goto finIntento; }
            string ff=ffin.Split(' ')[0];
            DateTime dt;
            if(!DateTime.TryParseExact(ff,new string[]{"d/M/yyyy","dd/MM/yyyy"},System.Globalization.CultureInfo.InvariantCulture,0,out dt)){
              nota="fecha ilegible: "+ffin; ffin=""; goto finIntento;
            }
            ffin=dt.ToString("dd/MM/yyyy");
            // Hacia atras no hay limite; hacia adelante son 3 meses. La frontera efectiva
            // es una sola: hoy+3m. Una activa vencida hace dos anios SI es renovable.
            if(dt<tol){
              est="SI";
              if(dt<hoy) nota="ya vencido el "+ffin;
              else if(dt>lim) nota="a "+((int)(dt-lim).TotalDays)+" dias de la ventana, dentro de la tolerancia";
            } else {
              est="NO";
              // matiz informativo, no cambia el bucket
              nota=(dt>=unAnio)?"renovada recientemente":"fuera de la ventana";
            }
          } catch(Exception ex){ nota="error: "+ex.Message; transitorio=true; }
          finIntento:
          if(abortado) break;
          // Un fallo por tiempo no se reintenta: si Azul tardo 40 s en no responder,
          // no se va a componer en 1.2 s, y el reintento duplica la espera para nada.
          if(porTiempo){ P("   se agoto el tiempo en "+nums[n]+", no se reintenta"); break; }
          if(!transitorio || intento==2) break;
          P("   fallo pasajero, reintentando "+nums[n]);
          CerrarDetalle();
          Pump(1200);
        }
        // Estado de linea distinto de Activa: dos casos (Dorian, 09/09/2026).
        // "Suspendida" se clasifica IGUAL que Activa: la fecha/Compromiso de arriba ya
        // decidio est/esSim, exactamente como para una linea Activa. Solo se marca para
        // que quede trazable en la salida -- no se toca est ni esSim.
        // Cualquier OTRO estado que no sea Activa ni Suspendida sigue sin clasificarse: no
        // se excluye (excluir por un valor desconocido romperia el cuadre), va a REVISAR
        // con el valor crudo. Nunca se adivina.
        bool esSuspendida=false;
        if(!abortado && !els[n].StartsWith("Activ",StringComparison.OrdinalIgnoreCase)){
          if(els[n].StartsWith("Suspend",StringComparison.OrdinalIgnoreCase)){
            esSuspendida=true;
          } else {
            nota="estado de linea no reconocido: '"+els[n]+"'"+(nota.Length>0?" | "+nota:"");
            est="REVISAR"; esSim=false;
          }
        }
        // Estado y Nota ya no son columnas: Estado solo decide en que tabla cae la linea,
        // y el matiz que hay que ver va junto al numero, sin inventar columnas nuevas.
        // Se combinan con " | " -- una linea puede ser Suspendida Y SIM a la vez.
        string numOut=nums[n];
        var tags=new List<string>();
        if(esSim) tags.Add("SIM");
        if(esSuspendida) tags.Add("Suspendida");
        if(est=="REVISAR" && nota.Length>0) tags.Add(nota);
        if(tags.Count>0) numOut=nums[n]+" ("+string.Join(" | ",tags)+")";
        P("   = "+est+"  plan="+plan+"  disp="+disp+"  mpe="+mpe+"  forz="+forz+"  ff="+ffin+"  nota="+nota);
        string fila=Csv(numOut)+","+Csv(plan)+","+Csv(disp)+","+Csv(mpe)+","+Csv(forz)+","+Csv(ffin);
        // una SIM sin contrato no tiene Plan Forzoso: eso es correcto, no un fallo de
        // lectura. Contarla aqui haria sonar una alarma falsa.
        if(forz.Length==0 && est!="REVISAR" && !esSim) sinForz++;
        if(est=="SI"){ ySi++; filasSi.Add(fila); }
        else if(est=="NO"){ zNo++; filasNo.Add(fila); }
        else { wRev++; filasRev.Add(fila); }
        if(abortado){
          for(int m=n+1;m<rows.Count;m++){
            wRev++;
            filasRev.Add(Csv(nums[m]+" (no leida: corrida abortada)")+",\"\",\"\",\"\",\"\",\"\"");
          }
          break;
        }
      }
    } finally {
      for(int g=0;g<4 && CerrarDetalle();g++){}
    }

    // Un solo encabezado y las filas agrupadas. Los separadores empiezan con '#', que es
    // justo lo que agregar_a_word.ps1 descarta al filtrar por ^(Numero,|"): el paso a Word
    // sigue sirviendo sin tocarlo, y ahora recibe las renovables primero.
    L("Numero,Plan,Dispositivo,MPE,PlanForzoso,FechaExpiracion");
    L("# RENOVABLES ("+filasSi.Count+")");
    foreach(string s in filasSi) L(s);
    L("# NO RENOVABLES ("+filasNo.Count+")");
    foreach(string s in filasNo) L(s);
    if(filasRev.Count>0){
      L("# A REVISAR ("+filasRev.Count+") - el motivo va junto al numero");
      foreach(string s in filasRev) L(s);
    }
    // las canceladas salen como filas normales (solo el numero) para que el paso a Word
    // las lea igual que los otros bloques, sin un formato aparte que mantener
    L("# CANCELADAS EXCLUIDAS ("+canceladas.Count+")");
    foreach(string s in canceladas) L(Csv(s)+",\"\",\"\",\"\",\"\",\"\"");
    L("#");
    L("# CUADRE: "+rows.Count+" activas leidas + "+canceladas.Count+" canceladas excluidas = "+filasConNumero+" filas de la tabla");
    L("# "+ySi+" renovables, "+zNo+" no renovables, "+wRev+" a revisar"+(sinForz>0?("  ("+sinForz+" sin Plan Forzoso legible)"):""));
    L("# renovable = vence antes del "+tol.ToString("dd/MM/yyyy")+" = ventana al "+lim.ToString("dd/MM/yyyy")+" + 10 dias de tolerancia, sin limite hacia atras");
    L("# corrida "+hoy.ToString("dd/MM/yyyy")+", "+(sw.ElapsedMilliseconds/1000.0).ToString("0.0")+" s, "+jab+" llamadas JAB");
    P("JAB total de la lectura: "+jab);
    return log.ToString();
  }
}
'@
# Aqui vive lo que toca la maquina: capturar, traer al frente y hacer clic real.
. (Join-Path $PSScriptRoot 'captura.ps1')

# Cierra el marco "Cuenta: ..." con la X de su barra de titulo. El puente no expone los
# botones de la barra, pero si la geometria del marco, asi que la X se calcula a partir de
# ella en vez de clavar un par de numeros que dejarian de servir si el marco se moviera.
function Close-MarcoCuenta {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $geo = [Azul]::GeoMarcoCuenta()
  if ($geo -eq "") { return $true }
  $g = $geo.Split(',')
  $r = Get-AzulRect -Hwnd $Hwnd
  # la X es el ultimo boton de la barra, pegado al borde derecho del marco
  $cx = [int]$g[0] - $r.Left + [int]$g[2] - 21
  $cy = [int]$g[1] - $r.Top + 9
  [void](Invoke-AzulClick -Hwnd $Hwnd -X $cx -Y $cy)
  $ok = [Azul]::EsperarSinMarcoCuenta(15)
  [Azul]::Nota("cierre del panel del cliente = $ok")
  return $ok
}

# Cierra el marco "Inicio de Interaccion ..." con la X de su barra de titulo. Misma cuenta
# geometrica que Close-MarcoCuenta, apuntando al marco de interaccion en vez de al del
# cliente. REGLAS.md Ciclo 3, 09/09/2026: confirmado que cerrarlo sin haber escrito nada en
# sus campos (Razon 1/2, Resultado) no pide guardar ni muestra ningun aviso.
function Close-MarcoInteraccion {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $geo = [Azul]::GeoMarcoInteraccion()
  if ($geo -eq "") { return $true }
  $g = $geo.Split(',')
  $r = Get-AzulRect -Hwnd $Hwnd
  $cx = [int]$g[0] - $r.Left + [int]$g[2] - 21
  $cy = [int]$g[1] - $r.Top + 9
  [void](Invoke-AzulClick -Hwnd $Hwnd -X $cx -Y $cy)
  $ok = [Azul]::EsperarSinMarcoInteraccion(15)
  [Azul]::Nota("cierre del marco de interaccion = $ok")
  return $ok
}

# GUARDA DE ARRANQUE. Un panel de cliente abierto ("Cuenta: <nombre>") REEMPLAZA al marco de
# interaccion y deja la tabla de Suscripciones fuera del alcance del puente (docs/mapa-datos.md
# seccion 5). Con el ahi, la corrida muere con "no encuentro la tabla de Suscripciones", que no
# dice nada del motivo real.
#
# Hasta el 05/09/2026 esto no hacia falta como guarda suelta: el unico que abria ese panel era
# el sistema, para la foto del cliente, y lo cerraba siempre en su bloque finally. Desde que la
# foto la saca Dorian a mano, quien lo abre es EL -- y puede dejarlo abierto, o Azul puede
# tardar en cerrarlo. O sea que el cuidado que Get-VistaCliente tenia al empezar sigue
# haciendo falta justo cuando Get-VistaCliente deja de existir: por eso se movio aqui ANTES de
# borrar nada.
#
# Se pregunta solo por los desktop panes cacheados. El caso normal es que NO haya panel, y con
# el recorrido completo del arbol como red de seguridad ese caso -- el de todas las corridas --
# pagaria una visita al arbol entero para contestar "no hay". Es justo lo que prohibe
# docs/trampas.md. Si algun dia un panel no colgara de un pane cacheado, la guarda no lo veria
# y la corrida fallaria como fallaba antes: mal mensaje, pero nunca datos equivocados.
function Close-PanelClienteAbierto {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $titulo = [Azul]::MarcoCuenta()
  if ($titulo -eq "") { return $true }

  Write-Host "Habia un panel de cliente abierto ($titulo). Se cierra antes de leer."
  [Azul]::Nota("arranque: panel de cliente abierto - $titulo - se cierra")
  $ok = Close-MarcoCuenta -Hwnd $Hwnd
  if (-not $ok) {
    Write-Host "AVISO: no se pudo cerrar el panel del cliente. La tabla de Suscripciones puede"
    Write-Host "       seguir tapada. Cierralo a mano con la X de su barra de titulo."
    [Azul]::Nota("arranque: NO se pudo cerrar el panel del cliente")
  }
  return $ok
}

# Abre la vista general del cliente, la fotografia y vuelve a cerrarla.
# Devuelve @{ Cliente = "<nombre>"; Imagen = "<ruta png>" }; los campos van vacios si fallo.
# Nunca lanza: que falte la foto no puede costar la corrida.
function Get-VistaCliente {
  param(
    [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
    [Parameter(Mandatory=$true)][int]$X,
    [Parameter(Mandatory=$true)][int]$Y
  )
  $res = @{ Cliente = ""; Imagen = "" }
  try {
    # Un panel de cliente abierto de antes haria que la espera de abajo diera por bueno un
    # marco viejo: la foto seria del cliente equivocado. Se cierra primero.
    if ([Azul]::MarcoCuenta().Length -gt 0) {
      [Azul]::Nota("captura: habia un panel de cliente abierto de antes, se cierra")
      [void](Close-MarcoCuenta -Hwnd $Hwnd)
    }
    [Azul]::Nota("captura: clic en el enlace Cliente en ($X,$Y)")
    [void](Invoke-AzulClick -Hwnd $Hwnd -X $X -Y $Y)
    try {
      # Se espera al TITULO, no al reloj: mientras carga, el marco se llama "Formulario" y
      # se ve como un formulario en blanco con boton "Crear". Una espera fija de 3 s
      # fotografiaba justo eso, y parece una imagen valida.
      $titulo = [Azul]::EsperarMarcoCuenta(40)
      if ($titulo -eq "") {
        throw "el panel del cliente no abrio en 40 s (el clic no dio en el enlace, o Azul no respondio)"
      }
      $res.Cliente = $titulo.Substring(8).Trim()
      [Azul]::Nota("captura: panel abierto - $titulo")

      # El titulo cambia antes de que terminen de pintarse los campos: se espera ademas a
      # que la ventana deje de cambiar.
      $swEstable = [Diagnostics.Stopwatch]::StartNew()
      $estabilizo = Wait-PantallaEstable -Hwnd $Hwnd -SegundosTope 30
      $swEstable.Stop()
      [Azul]::Nota("captura: espera de pantalla estable = $([Math]::Round($swEstable.Elapsed.TotalSeconds,2)) s")
      if (-not $estabilizo) {
        [Azul]::Nota("captura: la pantalla no se estabilizo en 30 s, se captura igual")
      }

      $destino = Join-Path $PSScriptRoot ("salidas\capturas\cliente_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".png")
      [void](Save-AzulShot -Hwnd $Hwnd -Ruta $destino)
      if (-not (Test-CapturaValida -Ruta $destino)) {
        Remove-Item $destino -Force -ErrorAction SilentlyContinue
        throw "la captura salio en blanco o en negro"
      }
      $res.Imagen = $destino
      [Azul]::Nota("captura: guardada en $destino")
    } finally {
      # Se cierra pase lo que pase: dejar el panel abierto rompe la siguiente corrida,
      # porque tapa la tabla de Suscripciones.
      [void](Close-MarcoCuenta -Hwnd $Hwnd)
    }
  } catch {
    [Azul]::Nota("captura: fallo - " + $_.Exception.Message)
    Write-Host "AVISO: no se pudo capturar la vista del cliente: $($_.Exception.Message)"
  }
  return $res
}

# Con la ventana minimizada, Swing no crea la subventana de detalle y JAB no la ve.
# Puede estar en segundo plano, pero NO minimizada. Get-AzulHwnd hace las dos guardas.
try { $hwnd = Get-AzulHwnd } catch { "ERROR: $($_.Exception.Message)"; exit 1 }
[Azul]::PROG = Join-Path $PSScriptRoot 'progreso.txt'
[Azul]::PAUSA_LINEAS  = $PausaLineas  * 1000
[Azul]::PAUSA_LECTURA = $PausaLectura * 1000

# Init() aqui y no dentro de Run() porque cerrar el panel exige un clic real, y eso vive en
# PowerShell. Run() vuelve a llamar a Init(), que es idempotente: cuesta un recacheo de panes.
# Si Init falla, no se hace nada y se deja que Run() lo reporte como lo reportaba siempre.
if ([Azul]::Init($hwnd)) { [void](Close-PanelClienteAbierto -Hwnd $hwnd) }

$out = [Azul]::Run($hwnd)
$csvPath = Join-Path $PSScriptRoot "azul_renovables.csv"
$out | Out-File -FilePath $csvPath -Encoding utf8

# ---- vista general del cliente ----
# Va DESPUES de leer las lineas, a proposito: el marco del cliente REEMPLAZA al de la
# interaccion y deja la tabla de Suscripciones fuera del alcance del puente. Si se hiciera
# primero y el regreso fallara, se perderia la corrida entera. Aqui, lo peor que pasa es
# que falte la imagen, y el CSV ya esta en disco.
$imagen = (Get-VistaCliente -Hwnd $hwnd -X $ClienteX -Y $ClienteY).Imagen
if ($imagen -eq "") { Write-Host "       Las lineas SI se leyeron y estan en el CSV." }

# ---- cerrar el marco de interaccion ----
# REGLAS.md Ciclo 3, 09/09/2026: sin este cierre, "Inicio de Interaccion" queda listado en
# el panel Abrir Ventanas de Azul y el cliente sigue cargado para el siguiente -- justo lo
# que el sistema debe evitar. Va AL FINAL, despues de la foto: no cambia lo que ya se leyo
# ni lo que ya se guardo, y si fallara, el CSV y la imagen ya estan a salvo. En un try
# aparte -- a diferencia del resto de este bloque, que ya esta probado -- porque es la
# primera vez que este paso concreto corre contra Azul real.
try {
  $cerroInteraccion = Close-MarcoInteraccion -Hwnd $hwnd
  if (-not $cerroInteraccion) {
    Write-Host "AVISO: no se pudo cerrar 'Inicio de Interaccion'. El cliente puede seguir"
    Write-Host "       cargado para el siguiente. Cierralo a mano con la X de su barra de titulo."
  }
} catch {
  [Azul]::Nota("cierre interaccion: fallo - " + $_.Exception.Message)
  Write-Host "AVISO: fallo al intentar cerrar 'Inicio de Interaccion': $($_.Exception.Message)"
}

# ---- de quien es este CSV ----
# Tres lineas de cabecera. La cuenta y la razon social vienen de la lista de ordenes, por
# parametro; la imagen viene de la vista general del cliente, arriba.
#
# Siguen siendo INERTES para el parser de agregar_a_word.ps1, que solo abre bloque con
# '# NOMBRE (n)' y solo toma como fila lo que empieza con comilla.
#
# No desaparecieron con la foto porque son lo unico que dice, mirando el archivo en disco, a
# que cliente pertenece y donde esta su imagen. Sin ellas, reintentar el volcado a Word a
# mano seria a ciegas.
if ($Cuenta -ne "" -or $Razon -ne "" -or $imagen -ne "") {
  $cab = @()
  if ($Cuenta -ne "") { $cab += "# CUENTA: $Cuenta" }
  if ($Razon  -ne "") { $cab += "# CLIENTE: $Razon" }
  if ($imagen -ne "") { $cab += "# IMAGEN: $imagen" }
  ($cab + (Get-Content $csvPath -Encoding UTF8)) | Out-File -FilePath $csvPath -Encoding utf8
}

$out
if ($imagen -ne "") { "Imagen: $imagen" }
"CSV: $csvPath"

# ---- paso a Word ----
# Word siempre esta abierto y el resultado siempre va al final, detras de un salto de
# pagina, asi que el segundo comando sobraba: la corrida escribe sola.
#
# La unica guarda que se conserva es la que de verdad importa: si la corrida quedo
# INCOMPLETA no se escribe. Meter una tabla incompleta en el documento de un cliente es peor
# que esperar. Ojo con la distincion: una linea ambigua (multiples MPE, fecha ilegible) es un
# RESULTADO legitimo y si va al documento, en su bloque A REVISAR. Lo que no va es una
# corrida que se rompio a medias.
if (-not $SinWord) {
  # La comprobacion vive en revisar_csv.ps1 desde que lote.ps1 tambien la necesita: corriendo
  # con -SinWord este bloque no se ejecuta, y era justo la corrida desatendida la que se
  # quedaba sin guarda.
  $problemas = (& (Join-Path $PSScriptRoot 'revisar_csv.ps1') -Csv $csvPath).Problemas

  ""
  if ($problemas.Count -gt 0) {
    "NO se escribio en Word: la corrida quedo incompleta."
    foreach ($m in $problemas) { "  - $m" }
    "  Revisa el CSV y, si te sirve asi, corre agregar_a_word.ps1 a mano."
  } else {
    try {
      & (Join-Path $PSScriptRoot 'agregar_a_word.ps1') -Csv $csvPath -Cuenta $Cuenta -Razon $Razon -Imagen $imagen
    } catch {
      "AVISO: no se pudo escribir en Word: $($_.Exception.Message)"
      "       El CSV esta completo; puedes correr agregar_a_word.ps1 cuando quieras."
    }
  }
}
