# Bitácora de incidentes

> Ábrelo **cuando algo se rompa**, para ver si ya pasó antes.
> **Una entrada por incidente**, no por sesión. Solo lo que se rompió y cómo se
> arregló. Lo que nunca falló no va aquí.
>
> Formato: **síntoma → causa → arreglo**. Al final, dónde quedó la guarda.
> Cada guarda de `docs/trampas.md` viene de una de estas entradas: **no se quitan.**

---

## 2026-09-30 · Azul se degrada con el uso — el puente nunca soltaba lo que pedía

**Síntoma.** Documentado desde el 04/09: misma cuenta, mismo código, 13 renovables a las
13:07 y 2 a las 13:24 con un núcleo de Azul al 100%; el 5º cliente de un bloque colapsa;
el 17/09 Azul murió a los 32 min "con 1.7 GB".

**Causa (la más probable, no demostrada).** Java Access Bridge fija en el heap de Azul cada
objeto que devuelve hasta que el cliente llama `releaseJavaObject` (documentación de Oracle,
API de JAB). Ningún archivo del proyecto la declaraba ni la llamaba, y cada cliente hace
decenas de miles de llamadas; el heap de Azul es de solo 512 MB (`max-heap-size` del JNLP
que sirve el proveedor, no se puede subir desde aquí).

**Lo que no es.** El 1.7 GB: medido el 30/09, Azul **nace** con ~1716 MB de memoria privada
(la reserva del proceso de 32 bits, con el Chromium embebido de JxBrowser). No es una fuga
ni un aviso por sí solo.

**Arreglo.** `releaseJavaObject` en `azul_fast.ps1` y `buscar_cuenta.ps1` sobre todo lo que
se descarta, más una memoria de la tabla de atributos para no pedir cada celda 4-5 veces.
Detalle en `ESTADO.md`. `-SinLiberar` lo apaga.

**Resultado.** BASE 035 completa con Azul recién abierto (31 min, 20 clientes, 0 a revisar,
0 cortes, 0 fallos de lectura). El CPU de Azul fue igual en aumento por tercios
(63% → 70% → 78%), sin llegar al 100%: **no está demostrado que el arreglo quite la
degradación**, solo que no rompe nada. Falta la corrida A/B con `-SinLiberar`.

**Guarda.** En este proyecto, todo objeto que devuelva el puente se suelta al terminar con
él. En `buscar_cuenta.ps1`, lo que devuelven `Find`/`MarcoPanes` se suelta solo (vía
`Anota`/`Sal`): ahí **no** se hace `Rel` a mano sobre eso, o se soltaría dos veces.

## 2026-09-29 · "sin contexto raíz" en la prueba de 3 clientes — no se reprodujo al diagnosticar

**Síntoma.** Reportado por correo: la prueba de 3 clientes falló en los tres con
`ERROR: sin contexto raiz. El puente no ve a Azul`, incluso con Azul recién abierto (PID
nuevo, 2 min de vida) y después de aplicar `jabswitch -enable` y reiniciar Azul dos veces
(PID 16800 y luego 13384). `AtBroker.exe` no estaba corriendo en ese momento.

**Diagnóstico repetido en esta sesión, con el mismo PID 13384 todavía abierto.** Todo de
solo lectura, sin tocar Azul:

- Las dos copias de `WindowsAccessBridge-32.dll` (`C:\Windows\SysWOW64\` y el `bin\` del
  JRE) son **idénticas**: mismo tamaño (178816 bytes), misma versión de archivo
  (8.0.4510.10) y mismo SHA256. No hay ninguna otra copia en el disco.
- `.accessibility.properties` existe en el perfil, con
  `assistive_technologies=com.sun.java.accessibility.AccessBridge`.
- `jp2launcher` (PID 13384) tiene cargadas `JavaAccessBridge-32.dll` y
  `JAWTAccessBridge-32.dll` del JRE de 32 bits.
- `jp2launcher` y el PowerShell que carga el puente corren en la **misma sesión de
  Windows** (SessionId 2) y **ninguno de los dos está elevado**.
- `AtBroker.exe` **sigue sin estar corriendo**, igual que en el reporte original.
- Con todo lo anterior sin cambiar, `diagnostico\jab_now.ps1` contra ese mismo PID 13384
  **sí obtuvo el contexto raíz**: 1020 nodos, el marco
  `Búsqueda: Contacto y Suscripción` visible y el botón `Ver Productos Asignados`
  habilitado. `diagnostico\verificar.ps1` dio `TODO OK`.

**Causa.** No identificada. El puente respondió con el mismo proceso de Azul, la misma
sesión y la misma ausencia de `AtBroker.exe` que cuando fallaba, así que ninguna de las
condiciones comprobadas explica el cambio. El candidato más probable, sin poder
confirmarlo, es un estado transitorio del lado del puente o de la JVM en el momento del
fallo original — quizás ligado a cuánto llevaba Azul abierto o a su carga en ese instante
(ver la degradación medida el 04/09/2026 en la sección de arriba) — que se resolvió solo
o con alguno de los reinicios, y no una causa de configuración permanente de esta máquina.

**Arreglo.** Ninguno aplicado: no hizo falta, el puente ya respondía al diagnosticar.

**Lo que no fue.** DLL duplicada o desactualizada, Java de 64 bits, elevación desigual
entre procesos, y sesión de Windows distinta — las cuatro descartadas con evidencia.
`AtBroker.exe` sigue sin poder confirmarse ni descartarse como causa: estaba apagado
tanto cuando fallaba como cuando funcionó.

**Confirmación.** Con Dorian avisado, se corrió `lote.ps1 -Limite 3 -Lista ".\azul\Lista _Ordenes.csv"`
contra Azul real, mismo PID 13384. El puente respondió de punta a punta, sin un solo
`ERROR: sin contexto raiz` en 9.5 min y 112 723 llamadas JAB combinadas: **2 escritos, 1
fallo, 0 saltados.** Dos clientes se buscaron, cargaron, leyeron y escribieron bien en
`BASE 034 PS1 VIRLAN.docx` (queda con 2 de 10, guardada y cerrada). El tercero falló **por
diseño, no por el puente**: la rejilla trajo 5 resultados y la primera fila no mencionaba su
razón social; la guarda de razón social lo detectó y no entró al cliente equivocado.

**⚠ Pendiente:** revisar ese cliente a mano en Azul (su cuenta está en
`azul\salidas\lote_progreso.csv` como `FALLO BUSQUEDA`) para ver cuál de
las 5 filas de la rejilla es la correcta, y solo entonces reintentar con `-Forzar` si hace
falta. Si `sin contexto raiz` vuelve a aparecer en una corrida futura, repetir esta misma
batería y además registrar si `AtBroker.exe` aparece corriendo cuando funciona y ausente
cuando falla, para saber si es la pieza que importa o una coincidencia — sigue sin poder
confirmarse ni descartarse.

→ `ESTADO.md`

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


