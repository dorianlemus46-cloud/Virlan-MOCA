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
  # esquina superior izquierda de la ventana de Azul. En 0 -- lo normal -- NO se usa una
  # posicion fija: se saca de donde esta HOY la lupita del telefono (Find-IconoLupa en
  # captura.ps1), que si se sabe encontrar mirando la ventana.
  #
  # Ese enlace es contenido de un navegador incrustado y el puente no lo ve. Hasta el
  # 18/09/2026 aqui habia 500,120, medido el 04/09/2026 (el enlace ARRANCA en x~491 y crece
  # hacia la derecha con el nombre, asi que 500 cae dentro del primer caracter). Ese dia la
  # barra entera se corrio 55 pixeles a la derecha y el clic cayo delante de la etiqueta:
  # un cliente se quedo sin foto. La lupita se corrio exactamente lo mismo (791 -> 846) y el
  # enlace tambien (491 -> 546): van en bloque, a 300 pixeles uno del otro medidos las dos
  # veces. Por eso se pica a 291 de la lupita, que son los mismos 9 pixeles dentro del
  # nombre que daba el 500.
  #
  # Con los dos mayores que 0 se pica ahi sin mirar, como salida de emergencia.
  [int]$ClienteX = 0,
  [int]$ClienteY = 0,
  # AVERIGUACION: volcar al progreso todos los campos del detalle de la PRIMERA linea leida,
  # no solo los cinco que el lector usa. Sirve para ver que mas trae Azul en esa tabla sin
  # adivinar -- por ejemplo si el nombre comercial del telefono esta en algun campo. No cambia
  # nada de lo que se lee ni de lo que se escribe.
  [switch]$VolcarCampos,
  # SOLO LEER lo que Dorian ya tiene abierto, sin tocar su pantalla despues (15/09/2026). Se
  # salta los tres pasos que mueven ventanas: cerrar un panel de cliente al arrancar, la foto
  # (que abre y cierra el panel del cliente) y cerrar la interaccion al final. Esa interaccion
  # la abrio el, no el sistema, y cerrarla puede pedir 'Descartar' sobre algo que no es nuestro
  # (REGLAS.md 4.1). Implica -SinWord: una consulta suelta no va a ninguna base.
  [switch]$SoloLeer,
  # Leer SOLO estas lineas (numeros separados por coma) y volcar la ficha completa de cada una
  # al progreso. Vacio = todas las lineas, lo de siempre.
  [string]$Numeros = ""
)
if ($SoloLeer) { $SinWord = [switch]$true }
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
  const string L_SIMEQ= "SIM y Equipos";
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
  // Para encontrar los avisos que Azul abre FUERA de su ventana principal. Ver PulsarDescartar.
  delegate bool EnumProc(IntPtr h,IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb,IntPtr p);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
  // Front() y RealClick() vivian aqui sin usarse, del enfoque viejo por capturas y clics.
  // Ahora que si hacen falta (para la banda de "Cliente:", que el puente no ve) se mudaron
  // a captura.ps1, que es donde vive todo lo que toca la maquina. Este archivo se queda
  // solo con el puente: lo que unicamente lee.
  static bool Habilitado(long ac){
    ACI i; if(!Inf(ac,out i)) return false;
    return i.states_en_US!=null && i.states_en_US.IndexOf("enabled")>=0;
  }

  static int vm; static long root;
  // La ventana principal, guardada en Init(). Solo se usa para saber de que proceso es, y asi
  // poder mirar las demas ventanas de ESE Azul cuando se busca un aviso (ver PulsarDescartar).
  static IntPtr hwndRaiz=IntPtr.Zero;
  static StringBuilder log=new StringBuilder();
  static void L(string s){ log.AppendLine(s); }

  // progreso en vivo a disco: sin esto la corrida es una caja negra
  public static string PROG="";   /* la fija PowerShell segun donde este el script */
  // Pausas ADICIONALES (ms), fijadas desde PowerShell tras el Add-Type. Se suman a las
  // esperas que ya habia en el bucle de lectura; no sustituyen ninguna.
  public static int PAUSA_LINEAS=4000;
  public static int PAUSA_LECTURA=7000;
  // Volcado de los campos del detalle de la PRIMERA linea que se lea. Es una herramienta de
  // averiguacion, no de produccion: el lector se queda con cinco campos de esa tabla (Plan
  // Movil, MPE, Fecha Final, Duracion, Marca y Modelo) y los demas ni se nombran. Cuando hace
  // falta saber QUE MAS trae Azul ahi -- por ejemplo si el nombre comercial del telefono esta
  // en algun campo, y no solo el codigo del fabricante que da "Modelo" -- esto lo ensena sin
  // tener que adivinar. Solo imprime al progreso; no cambia ni una decision de lectura.
  public static bool VOLCAR=false;
  static bool volcado=false;
  static bool volcadoEquipo=false;
  // Numeros separados por coma. Si trae algo, SOLO se leen esas lineas y de cada una se vuelca
  // la ficha entera al progreso. Para consultas sueltas de Dorian; vacio = todas, como siempre.
  public static string SOLO_NUMS="";
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

  // El marco del panel del cliente MIENTRAS CARGA. Azul lo abre llamandose "Formulario" y solo
  // al terminar lo renombra a "Cuenta: <nombre>"; como el resto de los marcos, numera las
  // instancias ("Formulario [2]"), asi que se compara por prefijo y no por igualdad.
  const string FORM_PREFIX = "Formulario";
  static long MarcoFormularioPanes(){
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.StartsWith(FORM_PREFIX)) return c;
      }
    }
    return 0;
  }

  // Un marco interno por el prefijo de su titulo, solo por las zonas cacheadas. Generico
  // porque ya van tres prefijos: "Cuenta: ", "Formulario" y "Contacto: ".
  static long MarcoPorPrefijo(string pref){
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US=="internal frame" && ci.name!=null && ci.name.StartsWith(pref)) return c;
      }
    }
    return 0;
  }

  // La ficha de Contacto.
  //
  // MEDIDO el 11/09/2026 con la prueba que deja el fallo: cuando la banda "Cuenta:" de arriba
  // viene VACIA, pulsar el enlace del cliente NO abre la vista general de la cuenta, abre esta
  // ficha, titulada "Contacto: <nombre>". El sistema solo conocia "Cuenta: ", asi que se
  // quedaba los 50 s esperando un panel que tenia abierto delante. Y como el cierre tambien
  // buscaba solo "Cuenta: ", la ficha quedaba abierta para siempre: en el segundo cliente ya
  // habia dos, y el tercero fallo distinto porque la suya ya estaba abierta y el clic no abrio
  // nada nuevo.
  //
  // No es la vista que quiere la base -- trae datos de contacto, no el resumen de la cuenta --
  // asi que no se fotografia: se reconoce, se cierra y el bloque se queda con su hueco.
  const string CONT_PREFIX = "Contacto:";
  public static string MarcoContacto(){
    long f=MarcoPorPrefijo(CONT_PREFIX); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.name==null?"":i.name;
  }
  public static string GeoMarcoContacto(){
    long f=MarcoPorPrefijo(CONT_PREFIX); if(f==0) return "";
    ACI i; if(!Inf(f,out i)) return "";
    return i.x+","+i.y+","+i.width+","+i.height;
  }
  public static bool EsperarSinMarcoContacto(int seg){
    DateTime tope=DateTime.Now.AddSeconds(seg);
    while(DateTime.Now<tope){
      if(MarcoPorPrefijo(CONT_PREFIX)==0) return true;
      Pump(400);
    }
    return MarcoPorPrefijo(CONT_PREFIX)==0;
  }

  // El panel del cliente, esperado en DOS FASES en vez de un solo reloj de 40 s.
  //
  // El 11/09/2026 la foto fallo 3 veces de 8 con "el panel no abrio en 40 s", y Dorian confirmo
  // que el panel SI se estaba cargando. El reloj unico no distinguia las dos causas y trataba
  // igual a la peor y a la mejor: al clic que no dio en el enlace le regalaba 40 s de espera
  // inutil, y al panel que estaba cargando de verdad le cortaba a los 40 s.
  //
  // Fase corta: que APAREZCA un marco, ya sea "Formulario" (cargando) o directamente
  // "Cuenta: ". Que aparezca es la prueba de que el clic dio en el enlace. Si no aparece
  // ninguno, el clic no dio, y se dice asi para que quien llama reintente el clic en vez de
  // esperar de brazos cruzados.
  //
  // Fase larga: con el marco ya abierto y cargando, esperar a que el titulo pase al nombre del
  // cliente. Aqui se puede esperar mas que antes SIN riesgo de premiar un fallo, porque ya se
  // sabe que hay algo cargandose de verdad.
  //
  // Devuelve el titulo, o "SINMARCO" / "CARGANDO" segun en que fase se quedo. Ningun titulo de
  // verdad puede confundirse con esos dos: todos empiezan por "Cuenta: ".
  // Devuelve el titulo si abrio la vista de cuenta, o uno de estos tres: "CONTACTO" si lo que
  // abrio fue la ficha de contacto, "SINMARCO" si no abrio nada, "CARGANDO" si abrio y no
  // termino. Ningun titulo de verdad puede confundirse con ellos: todos empiezan por "Cuenta: ".
  // CORREGIDO el 11/09/2026 con lo que Dorian vio en pantalla: la ficha de Contacto NO es el
  // final del camino, es el panel a medio cargar. La version anterior salia en el instante en
  // que la veia, asi que el sistema empezaba a cerrar todo mientras Azul todavia estaba trayendo
  // la informacion del cliente, y la foto no se tomaba nunca. Dos cosas lo respaldan: que el
  // titulo de esa ficha traia el nombre de la EMPRESA en dos de los seis casos, no el de una
  // persona, y que quien mira la pantalla vio la carga en curso.
  //
  // Ahora, ver una ficha de Contacto abre un plazo propio (segContacto) durante el que se sigue
  // esperando la vista de cuenta. Ese plazo es corto a proposito: si se usara el largo entero,
  // cada cliente que de verdad se quede en la ficha de Contacto costaria casi un minuto de
  // espera inutil.
  public static string EsperarPanelCliente(int segCorto,int segLargo,int segContacto){
    // Fase corta: que aparezca ALGO. Los tres titulos posibles valen igual como prueba de que
    // el clic dio en el enlace; distinguirlos es asunto de la fase larga.
    DateTime t1=DateTime.Now.AddSeconds(segCorto);
    bool abierto=false;
    while(true){
      string n=MarcoCuenta(); if(n.Length>0) return n;
      if(MarcoPorPrefijo(CONT_PREFIX)!=0 || MarcoFormularioPanes()!=0){ abierto=true; break; }
      if(DateTime.Now>=t1) break;
      Pump(300);
    }
    if(!abierto) return "SINMARCO";

    P("captura: el panel abrio y sigue cargando; se espera al nombre del cliente hasta "+segLargo+" s");
    DateTime t2=DateTime.Now.AddSeconds(segLargo);
    DateTime topeCont=DateTime.MaxValue;
    while(DateTime.Now<t2){
      string n=MarcoCuenta(); if(n.Length>0) return n;
      if(MarcoPorPrefijo(CONT_PREFIX)!=0){
        if(topeCont==DateTime.MaxValue){
          topeCont=DateTime.Now.AddSeconds(segContacto);
          P("captura: hay una ficha de Contacto abierta; se le dan "+segContacto+" s mas por si la vista de cuenta todavia esta cargando");
        }
        if(DateTime.Now>=topeCont) return "CONTACTO";
      }
      Pump(400);
    }
    return "CARGANDO";
  }

  // Que marcos internos hay ABIERTOS ahora mismo, por su titulo. Solo para diagnostico.
  //
  // Existe por una pregunta que el 11/09/2026 no se podia contestar: la foto del cliente fallo
  // tres veces de ocho con "el panel no abrio en 40 s", y con lo que se anotaba no habia forma
  // de distinguir las dos causas posibles -- que el clic no diera en el enlace y no se abriera
  // nada, o que el panel SI abriera y Azul no lo terminara de cargar. El segundo caso se ve
  // como un formulario en blanco cuyo titulo sigue siendo "Formulario", y EsperarMarcoCuenta()
  // espera el prefijo "Cuenta: ", asi que no lo reconoce y se agota igual.
  //
  // Mismo coste que MarcoCuentaPanes(): solo las zonas cacheadas, sin recorrer el arbol.
  public static string MarcosAhora(){
    var sb=new StringBuilder();
    for(int p=0;p<panes.Count;p++){
      ACI pi; if(!Inf(panes[p],out pi)) continue;
      int kc=pi.childrenCount;
      for(int k=0;k<kc;k++){
        long c=Kid(panes[p],k); if(c==0) continue;
        ACI ci; if(!Inf(c,out ci)) continue;
        if(ci.role_en_US!="internal frame") continue;
        if(sb.Length>0) sb.Append(" | ");
        sb.Append(ci.name==null?"(sin nombre)":ci.name.Trim());
      }
    }
    return sb.Length==0?"(ninguno)":sb.ToString();
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
  // ---- el aviso que Azul levanta al cerrar una subventana ----
  // REGLAS.md seccion 4.1, 11/09/2026. 'Descartar' es la unica excepcion a la lista de botones
  // prohibidos, y lo es bajo condiciones muy estrechas. Dos de ellas las garantiza quien llama,
  // no este codigo: solo se invoca desde el cierre de un marco que nuestra propia navegacion
  // abrio, y solo cuando ese marco NO se fue con la X. Las otras tres las garantiza esto:
  // se exige que el boton viva dentro de un cuadro de aviso de verdad, se deja por escrito lo
  // que decia ese aviso ANTES de pulsar, y no se pulsa nada cuyo nombre no sea exactamente
  // 'Descartar'.
  //
  // Que lo descartado sea el borrador de interaccion y no un dato del cliente no se deduce del
  // arbol: se sostiene en que el sistema es de solo lectura y nunca escribe en ningun campo, de
  // modo que el unico cambio pendiente que puede existir es el que creo la propia navegacion.
  const string BTN_DESC = "Descartar";

  // Roles que SI son un cuadro de aviso. Deliberadamente no incluye "internal frame": si el
  // boton apareciera dentro del marco del cliente en vez de dentro de un aviso, esto se niega a
  // pulsar y lo reporta, que es lo que corresponde cuando no se entiende lo que hay en pantalla.
  static bool EsAviso(string r){
    if(r==null) return false;
    return r=="dialog" || r=="alert" || r=="option pane" || r=="file chooser";
  }

  // Busca el boton y, de paso, el cuadro de aviso mas cercano que lo contiene y los roles del
  // camino. Los roles se devuelven para que una corrida que NO encuentre el cuadro deje dicho
  // que habia en su lugar: es la unica forma de aprender la forma real de este aviso sin
  // instrumentar Azul. No se entra en las tablas, que es donde vive casi todo el costo del
  // arbol y donde con seguridad no hay botones de aviso.
  static long BuscaDescartar(long ac,long avisoArriba,string caminoArriba,int prof,out long aviso,out string camino){
    aviso=0; camino="";
    if(prof>40) return 0;
    ACI i; if(!Inf(ac,out i)) return 0;
    string rol=(i.role_en_US==null?"?":i.role_en_US);
    string cam=(caminoArriba.Length==0?rol:(caminoArriba+" > "+rol));
    long av=avisoArriba;
    if(EsAviso(rol)) av=ac;
    if(rol=="push button" && i.name!=null && i.name.Trim()==BTN_DESC){ aviso=av; camino=cam; return ac; }
    if(rol=="table") return 0;
    int kc=i.childrenCount;
    for(int k=0;k<kc;k++){
      long c=Kid(ac,k); if(c==0) continue;
      long a2; string c2;
      long r=BuscaDescartar(c,av,cam,prof+1,out a2,out c2);
      if(r!=0){ aviso=a2; camino=c2; return r; }
    }
    return 0;
  }

  // Junta el texto legible de un nodo y su descendencia. Es para la constancia escrita del
  // aviso, asi que poco y corto: un cuadro de aviso tiene un mensaje y dos o tres botones.
  static string TextoAviso(long ac,int prof){
    if(prof>6) return "";
    ACI i; if(!Inf(ac,out i)) return "";
    var sb=new StringBuilder();
    if(i.name!=null && i.name.Trim().Length>0) sb.Append(i.name.Trim()+" / ");
    string d=Strip(i.description);
    if(d.Length>0) sb.Append(d+" / ");
    int kc=i.childrenCount;
    for(int k=0;k<kc;k++){ long c=Kid(ac,k); if(c==0) continue; sb.Append(TextoAviso(c,prof+1)); }
    return sb.ToString();
  }

  // Lo que devuelve es un codigo para que PowerShell decida que escribir en pantalla:
  //   SINAVISO      no hay ningun boton 'Descartar' en el arbol; lo que bloquea es otra cosa
  //   SINCUADRO:..  hay boton pero no dentro de un cuadro de aviso; NO se pulsa, y se dice que habia
  //   SINTEXTO      el cuadro esta pero no se le pudo leer nada; NO se pulsa, porque la regla
  //                 exige dejar constancia de lo que decia antes de tocarlo
  //   APAGADO       el boton esta ahi pero deshabilitado
  //   PULSADO / NOPULSO
  // Las ventanas de primer nivel VISIBLES del mismo proceso que la principal, ella aparte.
  // Un cuadro de aviso modal de Swing es una de estas: no cuelga del arbol de la ventana
  // principal, asi que desde 'root' no se ve por mucho que se recorra.
  static List<IntPtr> VentanasHermanas(){
    var l=new List<IntPtr>();
    if(hwndRaiz==IntPtr.Zero) return l;
    uint mio; GetWindowThreadProcessId(hwndRaiz,out mio);
    if(mio==0) return l;
    EnumProc cb=delegate(IntPtr h,IntPtr p){
      if(h!=hwndRaiz && IsWindowVisible(h)){
        uint pid; GetWindowThreadProcessId(h,out pid);
        if(pid==mio) l.Add(h);
      }
      return true;
    };
    EnumWindows(cb,IntPtr.Zero);
    return l;
  }

  // El intento sobre UN arbol. Lo de fuera decide sobre cuantos arboles se intenta.
  static string PulsarDesde(long raiz){
    long aviso; string camino;
    long b=BuscaDescartar(raiz,0,"",0,out aviso,out camino);
    if(b==0) return "SINAVISO";
    if(aviso==0){ P("aviso al cerrar: hay boton 'Descartar' pero no dentro de un cuadro de aviso. Camino: "+camino); return "SINCUADRO:"+camino; }
    string txt=TextoAviso(aviso,0).Trim();
    if(txt.Length>400) txt=txt.Substring(0,400);
    if(txt.Length==0){ P("aviso al cerrar: cuadro encontrado pero sin texto legible. Camino: "+camino); return "SINTEXTO"; }
    P("aviso al cerrar, texto completo: "+txt);
    if(!Habilitado(b)) return "APAGADO";
    bool ok=Click(b);
    P("'Descartar' pulsado = "+ok);
    if(!ok) return "NOPULSO";
    Pump(700);
    return "PULSADO";
  }

  // Primero la ventana principal, que es donde estaba mirando hasta el 17/09/2026. Si ahi no
  // hay nada, las demas ventanas del MISMO Azul: Dorian viene viendo el aviso y pulsandolo a
  // mano, y el registro no tenia ni una sola linea de "aviso al cerrar", lo que solo encaja
  // con que el boton estuviera donde no se buscaba.
  //
  // Esto NO afloja ninguna de las condiciones de la regla 4.1: sigue sin pulsarse nada que no
  // se llame exactamente 'Descartar', sigue exigiendose que viva dentro de un cuadro de aviso
  // y sigue escribiendose el texto del cuadro antes de tocarlo. Lo unico que cambia es donde
  // se busca. Y quien llama sigue siendo el mismo: el cierre de un marco que abrio nuestra
  // propia navegacion, cuando la X no basto.
  public static string PulsarDescartar(){
    string r=PulsarDesde(root);
    if(r!="SINAVISO") return r;
    var vs=VentanasHermanas();
    if(vs.Count==0) return "SINAVISO";
    P("aviso al cerrar: en la ventana principal no hay 'Descartar'; se miran "+vs.Count+" ventana(s) mas del mismo Azul");
    foreach(IntPtr h in vs){
      int vm2; long r2;
      if(!getAccessibleContextFromHWND(h,out vm2,out r2) || r2==0) continue;
      if(vm2!=vm) continue;   // otra maquina virtual Java: sus manejadores no valen para esta
      string s=PulsarDesde(r2);
      if(s!="SINAVISO"){ P("aviso al cerrar: estaba en una ventana aparte, no en la principal"); return s; }
    }
    return "SINAVISO";
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
    hwndRaiz=hwnd;
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
      if(SOLO_NUMS.Length>0 && (","+SOLO_NUMS+",").IndexOf(","+num+",")<0) continue;
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

            // Volcado de averiguacion, solo de la primera linea: la tabla de atributos entera,
            // campo por campo. Va aqui y no antes porque aqui ya se espero a que el arbol deje
            // de crecer, asi que lo que se vuelca es la tabla completa y no una a medio cargar.
            if((VOLCAR && !volcado) || SOLO_NUMS.Length>0){
              volcado=true;
              P("   ---- VOLCADO de campos del detalle: "+maxR+" filas x "+AC+" columnas ----");
              for(int r=0;r<maxR;r++){
                string v2=(AC>2)?Txt(at,r*AC+2):"";
                string v3=(AC>3)?Txt(at,r*AC+3):"";
                P("   ["+r+"]  "+nm[r]+"  |  "+v2+"  |  "+v3);
              }
              P("   ---- fin del volcado ----");
            }

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

            // dispositivo: Marca + Modelo, leidos tal cual. Es el respaldo, no el resultado:
            // "MOTOROLA XT2421-7" es el codigo del fabricante y no dice nada a quien vende.
            string marca="", modelo=""; int rMar=-1;
            for(int r=0;r<maxR;r++){
              if(nm[r].StartsWith(L_MAR)){ string v=Txt(at,r*AC+2); if(v!="N/A"&&v.Length>0&&marca.Length==0){ marca=v; rMar=r; } }
              else if(nm[r]==L_MOD){ string v=Txt(at,r*AC+2); if(v!="N/A"&&v.Length>0) modelo=v; }
            }
            disp=(marca+" "+modelo).Trim();

            // El modelo COMERCIAL, que es lo que Dorian pidio el 11/09/2026.
            //
            // MEDIDO ese dia volcando los 130 campos del detalle de una linea: el nombre de
            // venta del equipo NO es el valor de ningun campo. Es el NOMBRE de la fila que
            // agrupa al equipo -- "MOTOROLA MOTO G04S AZUL ATT" -- y de ella cuelgan "Marca
            // (Fabricante)" = MOTOROLA y "Modelo" = XT2421-7. El lector venia leyendo las dos
            // ramas y tirando el tronco, que es el unico que trae el nombre comercial.
            //
            // Esa fila se distingue de un campo en que sus DOS columnas de valor vienen vacias,
            // y se reconoce por la marca que ya se leyo: empieza por ella. Atarlo a la marca y
            // no a una fila fija es lo que lo hace sobrevivir a que Azul reordene la tabla, y
            // la reordena: el numero de campos depende de lo que la linea tenga contratado.
            //
            // Si no aparece, disp se queda con Marca + Modelo. Un telefono sin nombre de venta
            // en Azul -- o una linea que es solo SIM -- no es un error, es un dato que no esta.
            string origenDisp="";   // "marca", "forma" o vacio si no se hallo: va al progreso
            if(marca.Length>0){
              for(int r=0;r<maxR;r++){
                if(!nm[r].StartsWith(marca)) continue;
                if(nm[r].Length<=marca.Length) continue;   // la fila es la marca sola, no el nombre
                if(Txt(at,r*AC+2).Length>0) continue;      // trae valor: es un campo, no el equipo
                if(AC>3 && Txt(at,r*AC+3).Length>0) continue;
                disp=nm[r].Trim(); origenDisp="marca"; break;
              }
            }

            // Respaldo por FORMA, 15/09/2026. Dorian vio que muchos iPhone salian como "APPLE
            // A3295": el nombre de venta si esta en la columna de etiquetas, pero no empieza por
            // la marca, y la regla de arriba lo exige. Tambien pasaba con HONOR y SAMSUNG.
            //
            // Aqui no se mira el texto sino la forma de la fila del equipo: sube desde "Marca
            // (Fabricante)" y es la primera con la columna de valor vacia, escrita toda en
            // mayusculas y con mas de una palabra. Los campos vacios que hay en medio
            // ("EquiposOrderID", "Codigo de desbloqueo otorgado") llevan minusculas, y "SIM" es
            // una sola palabra. No se sube mas alla de "SIM y Equipos", la seccion del equipo.
            //
            // La ULTIMA columna NO se exige vacia, a diferencia de la regla por marca. MEDIDO el
            // 15/09/2026 con el volcado de un HONOR que salia "HONOR ALT-LX3": la fila del equipo
            // era "HONOR X7C BDL EARBUDS X7I GREEN ATT" con "Precio de equipo" en la ultima
            // columna, y exigirla vacia era justo lo que la tiraba. En el MOTOROLA del 11/09
            // venia vacia; las dos formas caben aqui.
            if(origenDisp.Length==0 && rMar>0){
              for(int r=rMar-1;r>=0;r--){
                string s3=nm[r].Trim();
                if(s3==L_SIMEQ) break;
                if(s3.IndexOf(' ')<0) continue;
                if(s3!=s3.ToUpperInvariant()) continue;
                bool letra=false; foreach(char ch in s3){ if(char.IsLetter(ch)){ letra=true; break; } }
                if(!letra) continue;
                if(Txt(at,r*AC+2).Length>0) continue;
                disp=s3; origenDisp="forma"; break;
              }
            }
            if(disp.Length==0) disp="N/A";
            if(origenDisp!="marca"){
              P("   nombre del equipo: "+(origenDisp=="forma" ? "por forma, no empieza por la marca" : "NO hallado, queda Marca + Modelo")+" -> "+disp);
              // Averiguacion: la seccion del equipo de la primera linea en que la marca no basto,
              // con la columna 1, que el volcado general no ensena. Solo imprime.
              if(VOLCAR && !volcadoEquipo && rMar>=0){
                volcadoEquipo=true;
                P("   ---- VOLCADO de la seccion del equipo: filas "+Math.Max(0,rMar-40)+" a "+Math.Min(maxR-1,rMar+1)+" ----");
                for(int r=Math.Max(0,rMar-40);r<=Math.Min(maxR-1,rMar+1);r++){
                  P("   ["+r+"]  "+nm[r]+"  |  "+(AC>1?Txt(at,r*AC+1):"")+"  |  "+(AC>2?Txt(at,r*AC+2):"")+"  |  "+(AC>3?Txt(at,r*AC+3):""));
                }
                P("   ---- fin del volcado ----");
              }
            }

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
#
# El aviso que Azul levanta a veces al cerrar una subventana. Hasta el 11/09/2026 lo despachaba
# Dorian a mano: mientras no se pulsaba, la subventana no se iba, el cliente seguia cargado para
# el siguiente y todos sus recorridos del arbol costaban hasta nueve veces mas.
#
# Se llama SOLO desde el cierre de un marco y SOLO cuando la X no lo cerro. Esas dos condiciones
# son parte de la regla, no una optimizacion, pero tambien son lo que hace que esto no cueste
# nada: en las corridas en las que Azul no pregunta, el marco se va con la X y aqui no se entra.
# Cuando si se entra, se paga un recorrido del arbol, que es justo la situacion en la que no
# pagarlo sale mucho mas caro.
function Resolver-AvisoDescartar {
  param([string]$Que = "una subventana")
  $r = [Azul]::PulsarDescartar()
  if ($r -eq 'PULSADO') {
    Write-Host "Azul pidio confirmacion al cerrar $Que y se pulso 'Descartar'."
    return $true
  }
  # Antes esto se iba en silencio. Si la subventana no se cerro y ademas no hay ningun aviso a
  # la vista, eso es justo lo que hay que poder leer despues en el registro, no un hueco.
  if ($r -eq 'SINAVISO') {
    [Azul]::Nota("aviso al cerrar ${Que}: no hay ningun boton 'Descartar' en ninguna ventana de Azul")
    return $false
  }
  Write-Host "AVISO: al cerrar $Que hay un aviso que NO se pulso ($r). Queda para revisar a mano."
  return $false
}

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
  # Mismo presupuesto de 15 s que antes, partido en dos: lo que tarda un cierre normal, y el
  # resto para despues de despachar el aviso.
  $ok = [Azul]::EsperarSinMarcoCuenta(4)
  if (-not $ok) {
    [void](Resolver-AvisoDescartar -Que "el panel del cliente")
    $ok = [Azul]::EsperarSinMarcoCuenta(11)
  }
  [Azul]::Nota("cierre del panel del cliente = $ok")
  return $ok
}

# Cierra la ficha "Contacto: ..." con la X de su barra de titulo. Misma cuenta geometrica que
# Close-MarcoCuenta. Hace falta desde el 11/09/2026: hasta entonces nada sabia cerrar esta
# ficha y se quedaba abierta toda la corrida, tapando la tabla y encareciendo cada recorrido
# del arbol del cliente siguiente.
function Close-MarcoContacto {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $geo = [Azul]::GeoMarcoContacto()
  if ($geo -eq "") { return $true }
  $g = $geo.Split(',')
  $r = Get-AzulRect -Hwnd $Hwnd
  $cx = [int]$g[0] - $r.Left + [int]$g[2] - 21
  $cy = [int]$g[1] - $r.Top + 9
  [void](Invoke-AzulClick -Hwnd $Hwnd -X $cx -Y $cy)
  $ok = [Azul]::EsperarSinMarcoContacto(4)
  if (-not $ok) {
    [void](Resolver-AvisoDescartar -Que "la ficha de Contacto")
    $ok = [Azul]::EsperarSinMarcoContacto(11)
  }
  [Azul]::Nota("cierre de la ficha de Contacto = $ok")
  return $ok
}

# Cierra el marco "Inicio de Interaccion ..." con la X de su barra de titulo. Misma cuenta
# geometrica que Close-MarcoCuenta, apuntando al marco de interaccion en vez de al del
# cliente.
#
# La nota del Ciclo 3 del 09/09/2026 decia que cerrarlo sin haber escrito nada en sus campos
# (Razon 1/2, Resultado) no pide guardar ni muestra ningun aviso. Eso quedo desmentido el
# 11/09/2026: Dorian viene despachando a mano un aviso que si aparece. De ahi el paso nuevo.
function Close-MarcoInteraccion {
  param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
  $geo = [Azul]::GeoMarcoInteraccion()
  if ($geo -eq "") { return $true }
  $g = $geo.Split(',')
  $r = Get-AzulRect -Hwnd $Hwnd
  $cx = [int]$g[0] - $r.Left + [int]$g[2] - 21
  $cy = [int]$g[1] - $r.Top + 9
  [void](Invoke-AzulClick -Hwnd $Hwnd -X $cx -Y $cy)
  $ok = [Azul]::EsperarSinMarcoInteraccion(4)
  if (-not $ok) {
    [void](Resolver-AvisoDescartar -Que "'Inicio de Interaccion'")
    $ok = [Azul]::EsperarSinMarcoInteraccion(11)
  }
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
    # Lo mismo con la ficha de Contacto. Sin esto, la ficha del cliente ANTERIOR sigue abierta y
    # el clic de este no abre nada nuevo: medido el 11/09/2026, asi fallo el tercer cliente de
    # tres, y el sintoma enganaba porque parecia que el clic no habia dado en el enlace.
    if ([Azul]::MarcoContacto().Length -gt 0) {
      [Azul]::Nota("captura: habia una ficha de Contacto abierta de antes, se cierra")
      [void](Close-MarcoContacto -Hwnd $Hwnd)
    }
    [Azul]::Nota("captura: clic en el enlace Cliente en ($X,$Y)")
    [void](Invoke-AzulClick -Hwnd $Hwnd -X $X -Y $Y)
    try {
      # Se espera al TITULO, no al reloj: mientras carga, el marco se llama "Formulario" y
      # se ve como un formulario en blanco con boton "Crear". Una espera fija de 3 s
      # fotografiaba justo eso, y parece una imagen valida.
      # 10 s para que el panel APAREZCA, 60 para que termine de cargar, y 15 mas de cortesia
      # desde que se ve una ficha de Contacto. En el caso bueno no cambia nada: las tres fases
      # salen en cuanto se cumple su condicion. Los numeros subieron el 11/09/2026 porque el
      # sistema iba a cerrar el panel mientras Azul aun cargaba.
      $titulo = [Azul]::EsperarPanelCliente(10, 60, 15)

      if ($titulo -eq 'SINMARCO') {
        # No se abrio NINGUN marco: el clic no dio en el enlace. Reintentar es seguro justo por
        # eso -- no hay nada abierto que un segundo clic pueda duplicar -- y es lo que antes no
        # se hacia nunca, porque el reloj unico no sabia distinguir este caso.
        [Azul]::Nota("captura: no aparecio ningun panel; el clic no dio en el enlace. Un reintento.")
        [void](Invoke-AzulClick -Hwnd $Hwnd -X $X -Y $Y)
        $titulo = [Azul]::EsperarPanelCliente(10, 60, 15)
      }

      if ($titulo -eq 'SINMARCO' -or $titulo -eq 'CARGANDO') {
        # PRUEBAS DEL FALLO (11/09/2026). Este fallo no dejaba nada detras, y por eso llevaba
        # dias sin diagnosticar. La lista de marcos dice cual de las dos causas fue, y la foto
        # ensena que habia en pantalla.
        #
        # Va en su propio try: un diagnostico que rompa la corrida seria peor que no tenerlo.
        $porque = if ($titulo -eq 'SINMARCO') { "no se abrio ningun panel ni al reintentar: el clic no da en el enlace" }
                  else { "el panel abrio pero Azul no lo termino de cargar en 50 s" }
        try {
          [Azul]::Nota("captura: marcos abiertos al fallar = " + [Azul]::MarcosAhora())
          $dirP = Join-Path $PSScriptRoot 'salidas\busquedas'
          if (-not (Test-Path -LiteralPath $dirP)) { [void](New-Item -ItemType Directory -Path $dirP -Force) }
          $pru = Join-Path $dirP ("panel_no_abrio_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".png")
          [void](Save-AzulShot -Hwnd $Hwnd -Ruta $pru)
          [Azul]::Nota("captura: prueba del fallo guardada en $pru")
        } catch {
          [Azul]::Nota("captura: no se pudo dejar prueba del fallo - " + $_.Exception.Message)
        }
        throw "no se pudo abrir el panel del cliente: $porque"
      }
      if ($titulo -eq 'CONTACTO') {
        # Se espero y lo que Azul dejo abierto es la ficha de Contacto. Se fotografia esa: en
        # pantalla esta la informacion del cliente, y una foto vale mas que un hueco. Queda
        # escrito de cual ficha es, para que la imagen del documento se pueda rastrear.
        $quien = ([Azul]::MarcoContacto()) -replace '^Contacto:\s*', ''
        $res.Cliente = $quien.Trim()
        [Azul]::Nota("captura: la vista que quedo abierta es la ficha de Contacto ($quien). Se fotografia esa.")
        Write-Host "       Azul dejo la ficha de Contacto para este cliente; la foto es de esa ficha."
      } else {
        $res.Cliente = $titulo.Substring(8).Trim()
        [Azul]::Nota("captura: panel abierto - $titulo")
      }

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
      # porque tapa la tabla de Suscripciones. Las dos clases de panel, no solo la de cuenta:
      # la ficha de Contacto estuvo quedandose abierta hasta el 11/09/2026 justo por faltar
      # esta linea, y cada una que se acumulaba encarecia todos los recorridos siguientes.
      [void](Close-MarcoCuenta -Hwnd $Hwnd)
      [void](Close-MarcoContacto -Hwnd $Hwnd)
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
[Azul]::VOLCAR        = [bool]$VolcarCampos
[Azul]::SOLO_NUMS     = (($Numeros -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join ','

# Init() aqui y no dentro de Run() porque cerrar el panel exige un clic real, y eso vive en
# PowerShell. Run() vuelve a llamar a Init(), que es idempotente: cuesta un recacheo de panes.
# Si Init falla, no se hace nada y se deja que Run() lo reporte como lo reportaba siempre.
if (-not $SoloLeer -and [Azul]::Init($hwnd)) { [void](Close-PanelClienteAbierto -Hwnd $hwnd) }

$out = [Azul]::Run($hwnd)
$csvPath = Join-Path $PSScriptRoot "azul_renovables.csv"
$out | Out-File -FilePath $csvPath -Encoding utf8

# ---- vista general del cliente ----
# Va DESPUES de leer las lineas, a proposito: el marco del cliente REEMPLAZA al de la
# interaccion y deja la tabla de Suscripciones fuera del alcance del puente. Si se hiciera
# primero y el regreso fallara, se perderia la corrida entera. Aqui, lo peor que pasa es
# que falte la imagen, y el CSV ya esta en disco.
#
# Y antes de eso: la foto solo sirve si el cliente va a entrar a la base. Abrir el panel del
# cliente, esperar a que la pantalla se estabilice y volver a cerrarlo cuesta bastante, y en un
# cliente sin nada que renovar ese gasto se tira. Se usa la MISMA regla que decide en lote.ps1
# (0 renovables y 0 a revisar = no entra), para que las dos no se puedan separar.
#
# Ademas es una senal a simple vista para Dorian: si Azul abre el panel del cliente, ese
# cliente va al documento. Si no lo abre, paso de largo.
#
# La foto se salta solo cuando los dos numeros se leyeron de verdad. Si el CSV quedo sin el
# bloque -- corrida rota -- se hace lo de siempre: mejor una foto de mas que perder la prueba.
$nRen = -1; $nRev = 0
foreach ($ln in ($out -split "`r?`n")) {
  if     ($ln -match '^#\s+RENOVABLES\s*\((\d+)\)') { $nRen = [int]$Matches[1] }
  elseif ($ln -match '^#\s+A REVISAR\s*\((\d+)\)')  { $nRev = [int]$Matches[1] }
}
$valeLaFoto = -not ($nRen -eq 0 -and $nRev -eq 0)

# Del centro de la lupita del telefono al punto donde se pica el enlace "Cliente:". Medido
# dos veces con la barra en dos sitios distintos y salio igual: ver el comentario de -ClienteX.
$ENLACE_DESDE_LUPA_X = 291
$ENLACE_DESDE_LUPA_Y = 1

$imagen = ""
if (-not $SoloLeer) {
  if ($valeLaFoto) {
    # Donde picar el enlace "Cliente:" no se recuerda: se saca de donde esta HOY la lupita.
    # Ver el comentario de -ClienteX. Con -ClienteX/-ClienteY a mano se pica ahi sin mirar.
    $cx = $ClienteX; $cy = $ClienteY
    if (-not ($cx -gt 0 -and $cy -gt 0)) {
      $lupa = Find-IconoLupa -Hwnd $hwnd
      if ($lupa.Ok) {
        $cx = $lupa.X - $ENLACE_DESDE_LUPA_X
        $cy = $lupa.Y - $ENLACE_DESDE_LUPA_Y
        [Azul]::Nota("captura: lupita en ($($lupa.X),$($lupa.Y)), enlace Cliente en ($cx,$cy)")
      } else {
        $cx = 0; $cy = 0
        [Azul]::Nota("captura: sin lupita no se sabe donde esta el enlace -- $($lupa.Por)")
      }
    }
    if ($cx -gt 0 -and $cy -gt 0) {
      $imagen = (Get-VistaCliente -Hwnd $hwnd -X $cx -Y $cy).Imagen
    } else {
      Write-Host "AVISO: no se encontro la lupita, asi que no se sabe donde esta el enlace del cliente."
      Write-Host "       No se pica a ciegas: se deja el hueco de la foto."
    }
    if ($imagen -eq "") { Write-Host "       Las lineas SI se leyeron y estan en el CSV." }
  } else {
    [Azul]::Nota("captura: 0 renovables y 0 a revisar, no se fotografia; se cierra el cliente")
    Write-Host "       Sin nada que renovar: no se toma foto, se cierra el cliente."
  }
}

# ---- cerrar el marco de interaccion ----
# REGLAS.md Ciclo 3, 09/09/2026: sin este cierre, "Inicio de Interaccion" queda listado en
# el panel Abrir Ventanas de Azul y el cliente sigue cargado para el siguiente -- justo lo
# que el sistema debe evitar. Va AL FINAL, despues de la foto: no cambia lo que ya se leyo
# ni lo que ya se guardo, y si fallara, el CSV y la imagen ya estan a salvo. En un try
# aparte -- a diferencia del resto de este bloque, que ya esta probado -- porque es la
# primera vez que este paso concreto corre contra Azul real.
try {
  if ($SoloLeer) {
    [Azul]::Nota("solo leer: la interaccion se deja abierta, es de Dorian")
  } else {
    $cerroInteraccion = Close-MarcoInteraccion -Hwnd $hwnd
    if (-not $cerroInteraccion) {
      Write-Host "AVISO: no se pudo cerrar 'Inicio de Interaccion'. El cliente puede seguir"
      Write-Host "       cargado para el siguiente. Cierralo a mano con la X de su barra de titulo."
    }
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
