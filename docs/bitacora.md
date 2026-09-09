# Bitácora de incidentes

> Ábrelo **cuando algo se rompa**, para ver si ya pasó antes.
> **Una entrada por incidente**, no por sesión. Solo lo que se rompió y cómo se
> arregló. Lo que nunca falló no va aquí.
>
> Formato: **síntoma → causa → arreglo**. Al final, dónde quedó la guarda.
> Cada guarda de `docs/trampas.md` viene de una de estas entradas: **no se quitan.**

---

## 2026-09-09 · Un bloque de UNA sola línea salía en Word sin encabezados

**Síntoma.** Encontrado al ejercitar `probar_word.ps1` por primera vez de verdad. En el
documento, los bloques con **una sola línea** —muy comunes: un cliente con una renovable, o
con una sola cancelada— salían como una tabla de **una fila con los datos y sin la fila de
encabezados**. Con dos líneas o más nunca falló, y por eso aguantó sin verse: nadie mira una
tabla de una fila y echa de menos el rótulo.

**Causa.** `ConvertFrom-Csv` devuelve un **objeto suelto y no un array** cuando le llega una
sola línea, y el `.Count` de ese objeto llega **vacío**, no `1`. La tabla se creaba con
`$null + 1 = 1` fila: la de encabezados, que después se llevaba por delante la fila de datos.
Medido aislado el mismo día: con una línea `.Count` sale vacío, con dos sale `2`.

**Arreglo.** Envolver en `@()` el resultado de `ConvertFrom-Csv`, para que sea siempre un
array y su cuenta sea siempre la real.

**Lo que no fue.** No tuvo nada que ver con las bases automáticas que se escribían ese día;
el fallo llevaba ahí desde que existe el bloque de tablas. Salió porque `probar_word.ps1` se
corrió por primera vez con Word abierto, que es lo que antes tenía prohibido.

→ `azul/agregar_a_word.ps1`, y la guarda nueva en `azul/diagnostico/probar_word.ps1`:
comprueba que **cada** tabla empiece por su fila de encabezados y traiga sus filas de datos.
Tres de los cuatro bloques del CSV sintético traen una sola línea a propósito.

---

## 2026-09-04 · El lote daba por fallidas dos búsquedas que salieron perfectas

**Síntoma.** Primera corrida real de `lote.ps1` (2 clientes, carpeta de prueba). Los dos
clientes se buscaron **de punta a punta y bien** —formulario, `Buscar Ahora`, rejilla,
`Seleccionar`, cliente cargado, pestaña Suscripciones con sus líneas—, y aun así el lote
escribió `no se pudo abrir el cliente` en los dos y saltó al siguiente. Nunca llamó al
lector. Final: `0 escrito(s), 2 fallo(s)`. En `lote_progreso.csv`, el detalle salió
truncado: `"buscar_cuenta.ps1 salio con "`, sin número.

**Causa.** El tope de tiempo recién puesto cambió `& $ps32 ... | Out-Host` + `$LASTEXITCODE`
por `Start-Process -PassThru` + `$p.ExitCode`. **`Start-Process -PassThru` entrega el objeto
sin conservar el handle del proceso**, así que cuando el hijo termina .NET ya no puede leer
su código y `ExitCode` devuelve `$null`. Y `$null -ne 0` es verdadero: todo hijo, saliera
como saliera, se contaba como fallo. Reproducido aislado el 04/09/2026 con `exit 0`,
`exit 1` y `exit 7`: los tres devolvían `$null`.

**Arreglo.** Tocar `$p.Handle` inmediatamente después de `Start-Process`, mientras el
proceso vive, para que .NET cachee el handle. Con esa línea, los mismos tres casos
devuelven 0, 1 y 7. Además, si `ExitCode` volviera a venir `$null` se registra en voz alta
y se devuelve `-998` —nunca 0—: dar por buena una búsqueda que no se pudo comprobar es la
vía a leer un cliente bajo el nombre de otro.

**Lo que no fue.** No fue Azul, ni el puente, ni el tope de tiempo. El tope no llegó a
dispararse (ningún hijo pasó de 300 s) y `omserver` no apareció en toda la corrida.

**Confirmado en produccion.** Corrida de las 12:48 del mismo dia: el progreso registro
`buscar_cuenta.ps1 salio con 1`, con numero, y el cliente que salio bien llego entero a Word.

→ `azul/lote.ps1`, función `Invoke-Hijo`

---
## 2026-08-31 · Todas las líneas salían `sin nodo Plan Móvil`

**Síntoma.** El lector no encontraba ni un solo atributo. Todas las líneas vacías.

**Causa.** En el árbol del detalle, el nombre del atributo **no está en `name`**: viene
en **`description`, envuelto en HTML** (`<html>Fecha Final del Compromiso</html>`).
Se estaba buscando en el campo equivocado.

**Arreglo.** Leer `description`, quitar las etiquetas HTML, y tomar el valor de `name`
de la columna 2.

→ `docs/mapa-datos.md` §4 · `docs/trampas.md` (otras condiciones)

---

## 2026-08-31 · `Ver Productos Asignados` seguía deshabilitado tras seleccionar la fila

**Síntoma.** El radio de la fila quedaba `checked` en el árbol, pero el botón no se
habilitaba y no había forma de abrir el detalle.

**Causa.** `doAccessibleActions` sobre el radio **cambia el widget pero no dispara la
lógica de Siebel** que habilita el botón. Un clic real habría funcionado, pero las
celdas de tabla reportan bounds `-1`: no hay coordenadas que usar.

**Arreglo.** `addAccessibleSelectionFromContext(vm, tabla, fila * columnas)`.
**El índice es de celda, no de fila.**

→ `docs/mapa-datos.md` §3 · `docs/trampas.md` (otras condiciones)

---

## 2026-08-31 · Una corrida tardaba 11 minutos con 20 s de CPU

**Síntoma.** Tiempos absurdos con la máquina casi ociosa: se estaba esperando, no
calculando.

**Causa.** Se recorría el árbol **dentro del bucle de espera**. Cada llamada JAB es IPC
síncrono que se bloquea mientras Azul carga; 30-40 recorridos por línea.

**Arreglo.** Cachear contextos y **no recorrer el árbol dentro de un bucle de espera**.
El `Find` no desciende dentro de nodos `table` (ahí están las miles de celdas).

→ `docs/trampas.md` (otras condiciones)

---

## 2026-08-31 · El detalle se leía a medias: 20 filas en vez de 120

**Síntoma.** Atributos faltantes de forma intermitente, distintos en cada corrida.

**Causa.** **El árbol de atributos carga perezosamente.** Leerlo apenas abre devuelve
solo parte.

**Arreglo.** Esperar a que `rowCount` **se estabilice**, no a que supere un umbral.
Esta espera **no se debe convertir en sondeo por condición**: no hay condición fiable
que preguntar. Se dejó intacta al optimizar el lector el 02/09.

→ `docs/trampas.md` (otras condiciones) · `docs/decisiones.md` (esperas por condición)

---

## 2026-09-01 · Todo se iba a REVISAR sin decir por qué

**Síntoma.** El lector clasificaba **todas** las líneas como REVISAR. Sin error, sin
mensaje, sin nada raro en el log.

**Causa.** **Encoding.** PowerShell 5.1 lee los `.ps1` sin BOM como ANSI:
`"Plan Móvil"` se convertía en `"Plan MÃ³vil"` y la comparación fallaba **en silencio**.

**Arreglo.** El archivo queda en **ASCII puro**, con los acentos como escapes
`\uXXXX`: `"Plan Móvil"` se escribe `"Plan M\u00F3vil"`.
`diagnostico\verificar.ps1` comprueba que no haya
bytes > 127 antes de cada corrida.

→ `docs/trampas.md` trampa 1 · `docs/decisiones.md` (por qué ASCII)

---

## 2026-09-02 · Azul se cayó a media corrida — cuenta sin leer

**Síntoma.** Corrida interrumpida. La cuenta con la línea `4770000001` quedó sin leer.

**Causa.** Caída de Azul, ajena al sistema.

**Arreglo.** Lo leído se conservó en `salidas\incompletas\`:
`2026-09-02_azul-se-cayo_falta-4770000001.csv` y `..._progreso.txt`.

**⚠ Pendiente:** esa cuenta **hay que releerla completa**. Sigue sin hacerse.

→ `ESTADO.md`

---

## 2026-09-02 · La foto del cliente salió de un formulario en blanco

**Síntoma.** Una captura que **parecía válida** y no lo era: un formulario vacío, con
campos obligatorios en blanco y un botón `Crear`.

**Causa.** Una **espera fija de 3 segundos** —que fue lo que se pidió al principio— no
alcanza a que el marco cargue. **Ninguna espera fija alcanza.**

**Arreglo.** Se espera al **título `Cuenta:`** del marco *y*, además, a que la pantalla
**deje de cambiar** (dos comparaciones iguales seguidas), con tope de reloj. Mismo
principio que la guarda del árbol de atributos.

→ `docs/trampas.md` trampa 5 · `docs/mapa-datos.md` §5

---

## 2026-09-02 · Las dos tablas de Word se fundían en una de 20 filas

**Síntoma.** Los títulos de bloque desaparecían y RENOVABLES / NO RENOVABLES quedaban
pegadas en una sola tabla.

**Causa.** `agregar_a_word.ps1` asignaba `Range.Text` a párrafos, y eso **borra la
marca de fin de párrafo**. Nunca se había visto porque la versión de "una tabla por
bloque" **jamás se había ejecutado contra Word**.

**Arreglo.** Todo pasa por `Add-Parrafo`, que usa `MoveEnd` para dejar la marca **fuera
del rango**.

→ `docs/trampas.md` trampa 12

---

## 2026-09-02 · El paso a Word tomó el CSV del día anterior

**Síntoma.** Sin argumentos, `agregar_a_word.ps1` escribió a partir de un CSV viejo.

**Causa.** Tomaba **el CSV más nuevo de `salidas\`** y eligió uno del día anterior.
Falló por casualidad (formato viejo); **con uno válido habría metido el cliente
equivocado en el documento sin decir nada.**

**Arreglo.** Toma `azul_renovables.csv`, que es el de la última corrida, y **avisa con
la fecha** si el CSV no es reciente.

→ `docs/trampas.md` (guardas)

---

## 2026-09-02 · Se escribió en una copia de autorrecuperación de Word

**Síntoma.** El bloque del cliente acabó en un documento que no era el de trabajo.

**Causa.** Se escribía a ciegas en `ActiveDocument`, y **eso no siempre es el documento
que uno cree**: Word había levantado una copia de autorrecuperación
(`BASE 031 PS1 (Recuperado automáticamente).docx`).

**Arreglo.** `agregar_a_word.ps1` **se niega** si el documento activo es autorrecuperado
o si nunca se ha guardado en disco. `-Forzar` lo salta.

**Secuela resuelta:** aquel documento quedó abierto sin guardar, con 26 páginas y tres
bloques del mismo cliente (CLIENTE EJEMPLO DOS) de tres corridas de prueba. **Dorian
lo dio por descartado el 03/09/2026**: era de pruebas.

→ `docs/trampas.md` trampa 13

---

## 2026-09-03 · El clic caía fuera del botón que se acababa de medir

**Síntoma.** Se medía un botón por el puente, se pulsaba, y el clic no daba en él.

**Causa.** **`Front()` des-maximizaba la ventana.** Llamaba a
`ShowWindow(SW_RESTORE)` **siempre**, y sobre una ventana maximizada eso la mueve y la
redimensiona — justo **entre medir y pulsar**.

**Arreglo.** Existe **`Frente()`**, que solo restaura si la ventana está **de verdad
minimizada**. **Usar siempre `Frente()`.**

**Consecuencia abierta:** el clic clavado de `Cliente:` en `(529,120)` se midió con la
ventana a 1366x768, **antes** de descubrir esto. Hay que confirmar que sigue acertando.

→ `docs/trampas.md` trampa 7 · `REGLAS.md` §5 · `ESTADO.md`

---

## 2026-09-03 · El tecleo se truncó y la búsqueda devolvió otro cliente

**Síntoma.** De `507000002` entraron solo `507`. **La búsqueda no falló: devolvió un
cliente distinto.** Ese es el peligro — un número truncado sigue siendo un número
válido.

**Causa.** El campo no siempre acepta el tecleo completo a la primera.

**Arreglo.** Se **relee el campo** después de escribir, se espera activamente y se
reintenta **hasta 3 veces**. Si al final no cuadra, **no se busca**.

→ `docs/trampas.md` trampa 8

---

## 2026-09-03 · Criterios viejos en el formulario devolvían otro cliente

**Síntoma.** La búsqueda por cuenta devolvía un cliente que no correspondía.

**Causa.** El formulario CIM **conserva criterios de búsquedas anteriores**, que se
combinan con el número de cuenta.

**Arreglo.** Se limpian todos los campos antes de escribir y **se reporta qué se
borró**. Ojo: **un campo con solo espacios cuenta** — hay que mirar `v.Length`, no
`v.Trim().Length`.

→ `docs/trampas.md` trampa 9

---

## 2026-09-03 · Se reportó "cuenta encontrada" dos veces sin haberla encontrado

**Síntoma.** El sistema daba la búsqueda por buena cuando no había resultados.

**Causa.** Se estaba usando **"el formulario se cerró"** como señal de éxito. **El
formulario se cierra igual sin coincidencias.**

**Arreglo.** Se distinguen **tres desenlaces** por señales reales: marco
`Inicio de Interacción [N] - <id>` (única), botón contador `N Registro/s` (varias), o
ninguna de las dos (sin resultados). Con rejilla, además se comprueba la **razón social
contra el Excel** antes de `Seleccionar`.

→ `docs/mapa-datos.md` §2

---

## 2026-09-03 · Azul ignoraba el clic en el ícono del teléfono

**Síntoma.** El clic en el ícono no abría el formulario de búsqueda.

**Causa.** Había **una rejilla de resultados abierta**. Con ella abierta, Azul ignora
ese clic.

**Arreglo.** Cerrar la rejilla antes de volver a buscar.

→ `docs/trampas.md` trampa 10

---

## 2026-09-03 · El marco de búsqueda no se encontraba en la segunda vuelta

**Síntoma.** Funcionaba la primera vez y fallaba la segunda.

**Causa.** **Azul numera las instancias de sus marcos**: la segunda vez el formulario se
llama `Encontrar Comunicante Búsqueda CIM [2]`. Se buscaba por igualdad exacta.

**Arreglo.** Buscar por **prefijo**. Relacionado: Azul **reutiliza el marco de
interacción**, así que su nombre puede ser idéntico para dos clientes distintos —
"existe el marco" es la condición, y que el nombre no cambie se avisa aparte.

→ `docs/trampas.md` trampas 6 y 11

---

*Procedencia: `azul\CONTEXTO-sesion-02-09-2026.md` (§2, §3, §5),
`azul\CONTEXTO-sesion-03-09-2026.md` (§3, §5, §7), `PROYECTO-AZUL.md`, `azul\LEEME.md`,
`azul\salidas\incompletas\`, y las memorias `azul-lectura-tecnica`, `azul-vista-cliente`.
Las fechas de las entradas del 31/08 y 01/09 son las de su registro, no necesariamente
las del fallo.*


