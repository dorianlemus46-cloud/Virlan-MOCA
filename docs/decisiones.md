# Decisiones — por qué las cosas son como son

> Ábrelo **cuando quieras cambiar un diseño**, para saber qué se rompió la última vez
> que se intentó lo contrario. Cada entrada dice qué se decidió, por qué, y cuándo.

---

## Por qué Java Access Bridge, y no otra cosa

**Decidido: 31/08 – 01/09/2026.** En producción.

Azul es Java Swing sobre Siebel. Se probaron y descartaron dos alternativas:

- **UI Automation / `pywinauto`** — la ventana es un `SunAwtFrame` y expone
  **0 descendientes** por UI Automation. No hay nada que leer. Descartado.
- **Capturas de pantalla + OCR + clics por coordenadas** — fue el primer método y
  funcionaba, pero depende de píxeles: cualquier cambio de tamaño, tema o zoom lo
  rompe. Sobrevive como **plan B** en `azul\respaldo-capturas\`, por si el puente se
  cae.

El puente **lee la interfaz como texto**, sin capturas y sin mover el mouse. Resultado
verificable: 19 líneas leídas con **resultado idéntico en dos corridas independientes**
— que es la prueba de que no está adivinando.

Requisito de entorno. En una máquina nueva se hace una vez, y está en `INSTALAR.md`:

```powershell
& "C:\Program Files (x86)\Java\jre1.8.0_451\bin\jabswitch.exe" -enable
```

La versión del JRE cambia de una máquina a otra; lo que no cambia es que sea el de **32
bits**, el de `Program Files (x86)`. Crea `.accessibility.properties` en el perfil del
usuario (`$env:USERPROFILE`). **Hay que reiniciar Azul** después, no hace falta cerrar
sesión de Windows. Se revierte con `jabswitch -disable`.

---

## Por qué clics y teclado reales para escribir

**Decidido: 03/09/2026.** Es el hallazgo que gobierna todo el diseño.

**El Java Access Bridge LEE esta aplicación perfectamente, pero NO la conduce.**
Costó tres corridas descubrirlo, y es lo que menos se deduce mirando el código:

- `setTextContents` devuelve **éxito** y deja el campo vacío, o lo llena sin que Siebel
  se entere. Una búsqueda con el número **a la vista** devolvía cero resultados.
- `doAccessibleActions` sobre `Buscar Ahora` **dispara el botón** —el formulario se
  cierra, se ve en el árbol— pero **Siebel no ejecuta la consulta**.
- Las pestañas del cliente son `page tab` con geometría y estado, pero con `act=no`:
  no tienen acción de accesibilidad.

**Conclusión operativa:** el puente sirve para **leer y para saber dónde está todo**;
escribir y pulsar va con **teclado y mouse reales** (SendInput con scancode Unicode).

**La consecuencia de diseño importa tanto como el hallazgo:** la geometría se le
pregunta al puente **en cada corrida** y el clic se da encima. **Nunca se clavan
coordenadas de algo que el puente pueda ver.** Las dos únicas excepciones son las dos
posiciones Chromium de `REGLAS.md` §5, donde no hay nada que preguntar.

---

## Por qué no se usan los botones `Exportar`

**Decidido: 01/09/2026.**

Azul tiene botones `Exportar` que darían los datos de golpe. **No se usan.** Se
mantiene el ciclo autorizado: seleccionar línea → `Ver Productos Asignados` → leer →
`Cerrar`.

El motivo es la regla de seguridad: `Exportar` no está dentro de ningún ciclo que
Dorian haya autorizado, y el proyecto entero se sostiene sobre no salirse de esos
ciclos. Está en la lista de nombres que nunca se pulsan (`REGLAS.md` §4).

---

## Por qué un CSV en medio

**Decidido: sin fecha registrada.** En producción desde el principio.

Lo que **sí** está registrado sobre el CSV:

- Es el **contrato entre las piezas**: `azul_fast.ps1` lo escribe, `agregar_a_word.ps1`
  y `revisar_csv.ps1` lo leen, y `lote.ps1` encadena a los dos. Cada paso puede
  reintentarse solo — si la captura falla, **las líneas no se pierden**: el CSV ya está
  escrito y se reintenta solo la foto con `-SoloCaptura`.
- Deja **rastro auditable** de cada corrida, que es lo que permite cuadrar
  (`activas leídas + canceladas excluidas = filas de la tabla`).
- Los bloques van separados por líneas `#` en vez de por una columna Estado, y el
  cliente y la imagen viajan en dos líneas de cabecera `# CLIENTE:` / `# IMAGEN:`.

> ⚠ **La elección del formato CSV como tal no aparece razonada en ninguno de los
> archivos existentes.** Lo de arriba describe la función que cumple, no el motivo por
> el que se eligió CSV frente a otra cosa. **Pendiente de que Dorian lo confirme** si
> quiere que quede registrado.

**Lo que sí está razonado son las columnas** — 6, y **sin columna Estado ni Nota**
(02/09/2026): Dorian las quitó explícitamente. El Estado es razonamiento interno, y el
bloque en el que cae la línea ya lo dice; el matiz va **entre paréntesis junto al
número** (`(SIM)`, `(múltiples MPE)`) en vez de abrir una columna nueva. Dorian pide
**tablas angostas**.

---

## Por qué las líneas sin nodo `Compromiso` pasaron de REVISAR a RENOVABLE

**Cambiado: 02/09/2026, por decisión de Dorian** (confirmado por él el 04/09/2026).

**Antes:** una línea sin nodo `Compromiso` iba a `REVISAR`, con el criterio de que sin
fecha límite no se puede clasificar y no se adivina.

**Ahora:** va a **RENOVABLE**, y se marca `(SIM)` junto al número.

**El motivo registrado:** una línea **activa** sin bloque `Compromiso` no es un dato
que falte — es una **SIM/eSIM sin contrato**, y por tanto sí es renovable. Es un caso
*conocido*, no una incógnita. En palabras de la memoria `azul-regla-renovacion`:
*"una línea activa sin bloque Compromiso es casi siempre una SIM sola, sin contrato, y
sí es renovable. Se marca `(SIM)` junto al número. Antes esto era REVISAR; ya no."*
Y en `PROYECTO-AZUL.md`: *"Sin nodo `Compromiso` ya no es `REVISAR`: es una SIM/eSIM
sin contrato y **sí es renovable**."*

**El principio que lo justifica, y que es lo importante:**
*"Sin datos" no siempre es duda.* REVISAR es para **ambigüedad real o fallo de
lectura**, no para ausencias explicables. Por eso el cambio **no relajó las demás
causas de REVISAR** —múltiples MPE, múltiples Fecha Final, fecha ilegible, atributos
que no cargaron, estado de línea desconocido—: esas siguen siendo ambigüedad o fallo,
y ahí no se adivina.

*Nota de procedencia: los archivos de la época registran el cambio como "corrección del
02/09/2026", sin citar la instrucción. Dorian confirmó el 04/09/2026 que fue decisión
suya.*

---

## Por qué las líneas `Suspendida` se tratan igual que `Activa`

**Decidido: 09/09/2026, por Dorian.**

**Antes:** una línea con estado distinto de `Activa` o `Cancelado` —incluida
`Suspendida`— entraba a `REVISAR` con el valor crudo, sin clasificar. Es lo que pasó en
la primera corrida real que trajo una línea `Suspendida` (cuenta 636000002, 09/09/2026):
salió como *"estado de linea no reconocido: 'Suspendida'"*.

**Ahora:** `Suspendida` sigue la **misma regla de fecha/Compromiso** que `Activa` —
mismo cálculo de `fin`, misma ventana de 3 meses + 10 días, mismo caso SIM/eSIM sin
Compromiso— y se marca `(Suspendida)` junto al número para que quede trazable, igual
que `(SIM)`. Cualquier OTRO estado que no sea `Activa`, `Suspendida` ni `Cancelado`
sigue yendo a `REVISAR` sin cambios: esto no relaja esa guarda, solo saca a
`Suspendida` de ella.

**El motivo:** Dorian confirmó que una línea suspendida también se puede renovar — no
es una ambigüedad de lectura, es un estado de negocio conocido, igual que ya lo era la
SIM sin contrato.

> ⚠ **Escrito y verificado sin Azul delante. Dorian pidió explícitamente no probarlo
> contra Azul real antes de este cambio** (09/09/2026, mismo día). Sigue sin verse una
> línea `Suspendida` clasificada como RENOVABLE contra una corrida real — lo único que
> se vio contra Azul real fue el comportamiento VIEJO (REVISAR), antes de este cambio.

---

## Por qué las bases volvieron a ser automáticas, y qué cambió para que el correlativo aguante

**Decidido: 09/09/2026, por Dorian.** Revierte la entrada de abajo
("Por qué el sistema dejó de nombrar las bases"), que estuvo vigente tres días.

**Lo que se pedía:** no tener que abrir Word ni nombrar nada. Una carpeta dentro del proyecto,
documentos `BASE 034 PS1 VIRLAN.docx`, `035`, `036`… y el sistema haciendo el resto.

**El problema real no era nombrar: era que el sistema DEDUCÍA en vez de recordar.** Miraba qué
archivos había en la carpeta para decidir el número siguiente, y esa carpeta era toda su
memoria. En cuanto Dorian renombraba una base terminada, el número quedaba libre y la corrida
siguiente lo reutilizaba. Mientras dedujera, no tenía arreglo — y por eso el 06/09 la salida
fue quitarle la potestad de nombrar.

**Ahora recuerda.** El número sale de un registro propio (`azul\bases\registro_bases.json`,
que escribe `azul\bases.ps1`) con cinco reglas que son el diseño entero:

1. **El número se aparta antes de que exista el documento.** Primero se anota que el 034 está
   tomado, después se crea el `.docx`. Si la creación falla, ese número **se quema**. Quemar un
   número no cuesta nada; reutilizar uno es lo que rompió el sistema la vez pasada.
2. **El contador nunca retrocede.** Ni borrando ni moviendo archivos. Si la carpeta trajera un
   número más alto que el registro, gana el más alto y se dice en voz alta.
3. **Jamás se sobrescribe.** Es la red de debajo. *Medido el 09/09/2026: con la regla 2 puesta,
   la 3 no llega a dispararse nunca. Se conserva igual — cuesta cuatro líneas y cubre el caso
   que no se nos ocurrió.*
4. **El registro se escribe entero o no se escribe.** Va a un temporal, se relee para
   comprobarlo, y solo entonces se renombra encima del bueno. Un registro roto **aborta y no se
   reinicia solo**: reiniciarlo repartiría números ya usados, que es justo el fallo a evitar.
5. **Si la base abierta desapareció del disco, se planta y pregunta.** No empieza otra con el
   mismo número. Es `REGLAS.md` §6 aplicada aquí.

**Por qué la memoria no habla con Word ni con Azul.** `bases.ps1` solo sabe de números, nombres
y de si un archivo está o no está. Esa separación es lo que permite probarlo entero sin Word
abierto y sin Azul delante, que es exactamente lo que no se podía hacer con la versión vieja.

**Dos cosas que trae de regalo.** Como el registro guarda quién está dentro de la base abierta,
relanzar una lista **ya no duplica clientes**: se saltan antes de buscarlos en Azul, que es lo
que hacía el `-Reanudar` que se quitó el 06/09. Y como una base son 10 clientes pero una tirada
son tres a cinco, **una base cruza varias corridas**: el lote continúa la que quedó a medias.

**Lo que se pierde.** El documento ya no lo elige una persona, así que la garantía deja de ser
"Dorian miró cuál era" y pasa a ser "la memoria lo recuerda". Si el registro se corrompiera y
se arreglara mal a mano, el sistema podría escribir en una base equivocada. Por eso el registro
aborta en vez de adivinar, y por eso es texto que se lee con los ojos.

> **El número de arranque es el 034, elegido por Dorian el 09/09/2026** porque la serie vieja
> del escritorio tiene ahí un hueco. La carpeta nueva y la vieja son independientes: cuando el
> contador pase por 035 y 036 producirá nombres que ya existen en el escritorio. No se pisa
> nada, pero conviene saberlo.

---

## ~~Por qué el sistema dejó de nombrar las bases~~

**Decidido: 06/09/2026, por Dorian. Superado el 09/09/2026** — ver la entrada de arriba. El
diagnóstico de por qué fallaba **sigue siendo correcto y sigue rigiendo**; lo que cambió es que
ahora hay una memoria de verdad, así que la cura ya no tiene que ser dejar de nombrar.

**Antes:** `New-Base` creaba el documento y lo guardaba de entrada como
`BASE NNN PS1 FAST.docx`; `Get-RutaBase` decidía el `NNN` mirando qué archivos existían en
la carpeta — **su única memoria**. Y Dorian bautiza las bases *al final*, con los 10 dentro.

**Ese desalineo tenía una consecuencia real:** en cuanto renombraba un documento terminado,
ese número quedaba libre y la corrida siguiente lo reutilizaba. No pisaba el archivo
renombrado, pero el correlativo dejaba de ser único.

**Ahora:** el sistema **no crea ni nombra documentos**. Escribe en el que Dorian tenga activo
en Word. Mientras el sistema nombrara, el problema no tenía arreglo; sin nombrar, **no se
mitiga: deja de existir**.

**El precio, y cómo se paga.** Ahora se puede escribir en el documento equivocado, que es un
error que no se ve hasta que ya está escrito. Por eso `Get-DocumentoDestino` hace cuatro
comprobaciones **antes del primer cliente** — no al ir a escribir, porque leer un cliente
cuesta dos o tres minutos y descubrirlo después sería tirarlos: Word abierto con al menos un
documento (y **no se crea ninguno**), `-Documento` obligatorio si hay más de uno abierto
(`REGLAS.md` §6, *"preguntar cuál"*), ruta en disco, y que no sea una copia de
autorrecuperación (trampa 13, que ahora pesa **más** que antes).

Dos consecuencias de que el documento sea suyo y no del sistema: `Activate()` va en **cada**
cliente —si pincha otra ventana de Word a media corrida, el `ActiveDocument` deja de ser el
nuestro— y al terminar **se guarda pero no se cierra**.

---

# Decisiones superadas

> No se borran. Explican por qué el código estuvo como estuvo, y **borrarlas invita a
> reproponer lo que ya se descartó**. Ninguna de las tres rige.

## ~~Por qué la foto del cliente va al final de la corrida~~

**Decidido: 02/09/2026. Superado el 06/09/2026** — la foto la saca Dorian a mano.

El marco de la vista del cliente **reemplaza al de la interacción** y tapa la tabla de
Suscripciones. Si la captura fuera al principio y el regreso fallara, **se perdería la
lectura entera**. Yendo al final, si fallaba la foto las tablas ya estaban leídas.

**El hecho que lo motivaba sigue siendo verdad, y ahora sostiene otra guarda:** un panel de
cliente abierto tapa la tabla, así que `azul_fast.ps1` comprueba y lo cierra al arrancar cada
cliente. Es el mismo peligro; lo que cambió es **quién** deja el panel abierto.

---

## ~~Por qué se lee el Excel abierto y no una ruta en disco~~

**Decidido: 03/09/2026. Superado el 06/09/2026** — ahora es un CSV (`docs/mapa-datos.md` §1).

El nombre del libro trae el número de descarga y **cambia en cada bajada**
(`ListaOrdenes (59).xls PNDIENE PS1.xls`). Una ruta fija quedaría obsoleta al día
siguiente. Se localizaba por la hoja **`resultData`** y por los encabezados.

**Sigue siendo cierto**, y por eso el CSV va a una ruta fija que Dorian sobrescribe cada
mañana, en vez de al nombre que le ponga Excel.

Por lo mismo, `excel_ordenes.ps1` lista **todos** los libros y hojas abiertos y no solo
el activo — la misma razón por la que `agregar_a_word.ps1` comprueba qué documento es
el activo antes de escribir. **Ese diagnóstico se conserva:** cuando la exportación a CSV
salga rara, es lo que dice qué hay de verdad en el libro del que salió.

---

## ~~Por qué 10 clientes por base, y por qué la base se guarda al crearla~~

**Decidido: 03/09/2026. Superado el 06/09/2026.** El **corte automático** se fue; **la regla
de los 10 no**, y vive ahora en `ARRANQUE.md`, que es donde la aplica quien la aplica.

- **10 clientes por documento.** Se guardaban como `BASE 034 PS1 FAST.docx`, `035`, `036`…
  en `C:\Users\Dell\Desktop\BASES DE DATOS PS1`. El sufijo **`FAST`** marcaba que las
  escribía el sistema. **Ese nombrado se fue entero** (ver la entrada de arriba).
- **Un cliente sin nada que renovar NO entra al documento** y **no cuenta para los
  10**; queda anotado en el progreso con el motivo. Sí entra si tiene renovables > 0
  **o** líneas en *A REVISAR* — porque *A REVISAR* no es "no renovable", es "no se pudo
  leer", y descartarlo **tiraría renovaciones buenas en silencio**.
  **Esto NO cambió.**
- **La base se guarda al crearla**, no al llegar a 10: `agregar_a_word.ps1` se niega a
  escribir en un documento sin ruta en disco, y así una caída a mitad **no se lleva lo
  ya escrito**. **El motivo sigue vigente y ahora se cumple solo**: el documento que abre
  Dorian ya tiene ruta, y el lote guarda cliente a cliente.
- Si un cliente falla, se anota y **se sigue con el siguiente**. **Esto NO cambió.** Lo que
  se fue es `-Reanudar`: sin él, repetir la lista entera duplicaría en el documento los
  clientes que sí salieron, así que para reintentar hay que exportar una lista con solo las
  cuentas que fallaron. Cuáles fueron está en `salidas\lote_progreso.csv`, que **se conserva**.

---

## Por qué volvió la captura automática del cliente

**Decidido: 08/09/2026, por Dorian.** Reactiva lo que la entrada de abajo
("~~Por qué la foto del cliente va al final de la corrida~~") da por superado.

Se quitó el 06/09/2026 solo por tiempo, no porque fallara: el mecanismo (clic en el
enlace `Cliente:` en `(500,120)`, esperar el título del marco, esperar a que la
pantalla se estabilice, fotografiar, cerrar) ya estaba probado contra Azul real antes
de quitarse. Dorian decidió que prefiere recuperar el tiempo de pegar la foto a mano
en vez de mantener ese ahorro.

**Qué se restauró, y qué no.** Se trajo de vuelta `Get-VistaCliente` y
`[Azul]::EsperarMarcoCuenta` tal como estaban en el último commit antes de Etapa 3
(`git diff 5ddc851 5fb8607`), con la misma coordenada. No se restauraron
`-SinCaptura`/`-SoloCaptura`: la captura corre siempre, sin excepción — cambia el
significado de la prueba "cliente suelto, sin foto" que documentaba `ESTADO.md`, que se
corrigió en consecuencia.

**Qué NO se tocó.** `Close-PanelClienteAbierto`/`Close-MarcoCuenta` —la guarda que
Etapa 3 dejó viva para cerrar un panel que quedara abierto— sigue exactamente igual:
`Get-VistaCliente` la reutiliza para su propio cierre, no hay dos mecanismos. Sigue
haciendo falta por el mismo motivo de siempre, solo que ahora protege contra que un
tope de tiempo mate la corrida a media captura, no contra que Dorian se distraiga.

**Efecto colateral esperado, no un bug:** como `lote.ps1` nunca pasó
`-SinCaptura`/`-SinFoto` (ni antes ni ahora), la base de 10 clientes vuelve a fotografiar
a cada uno automáticamente, no solo las corridas sueltas. Es exactamente el punto del
cambio: dejar de pegar 10 fotos a mano por base.

---

## Por qué el Word no se guarda solo

**Decidido: 02/09/2026.**

`agregar_a_word.ps1` escribe al final del documento y **no lo guarda**: Dorian revisa y
guarda. Un `Ctrl+Z` deshace lo agregado. El salto de página existe porque se procesa un
cliente tras otro en el mismo documento y las tablas no deben encimarse.

Esto convive con la regla de arriba (la base **del lote** sí se guarda al crearla, para
que tenga ruta en disco).

---

## Por qué ASCII puro en los `.ps1`

**Decidido: 01/09/2026**, después del fallo. Ver trampa 1 en `docs/trampas.md` y la
entrada correspondiente en `docs/bitacora.md`.

No es una preferencia de estilo: PowerShell 5.1 lee los `.ps1` sin BOM como ANSI y el
fallo es **silencioso** — no hay error, simplemente todas las líneas acaban en REVISAR.

---

## Por qué las esperas son por condición y no por reloj

**Decidido: 02/09 – 03/09/2026.**

Tres esperas fijas del lector pasaron a sondeo por condición (−32% de tiempo). La
excepción es **la estabilidad del árbol de atributos, que no se tocó y no se debe
tocar**: ahí no hay condición fiable que preguntar, por eso se espera a que deje de
crecer.

El caso que lo demostró: una espera de 3 segundos —que fue lo que se pidió al
principio— fotografiaba un **formulario en blanco** con botón `Crear`. Parecía una
imagen válida y no lo era.

---

*Procedencia: `PROYECTO-AZUL.md` (§ Decisiones tomadas, Vista general del cliente,
Búsqueda de la cuenta), `azul\CONTEXTO-sesion-03-09-2026.md` (§3, §4.6),
`azul\CONTEXTO-sesion-02-09-2026.md` (§2), `azul\LEEME.md`, y las memorias
`azul-lectura-tecnica`, `azul-regla-renovacion`, `azul-vista-cliente`,
`azul-flujo-por-cliente`.*
