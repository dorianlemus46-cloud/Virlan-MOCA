# ESTADO

> **Este archivo se SOBRESCRIBE.** No se crea uno nuevo por sesión.
>
> Al informar, distinguir **siempre** lo que está probado contra Azul real de lo que solo
> está escrito. Esa distinción importa más que cualquier otra cosa en este proyecto.

**Última actualización:** 09/09/2026 — **las bases automáticas vuelven, ahora con memoria
real** (`azul\bases.ps1`, carpeta `azul\bases`, numeración desde la 034). Revierte la etapa 4
de la simplificación sin traer de vuelta el fallo que la motivó. Antes de eso: simplificación
completa de seis etapas, la **etapa 5.4** (tope de tramo de 10 s) y la **restauración de la
captura automática del cliente**. Nada de esto se ha ejercitado contra Azul real.

> **Lo del 09/09/2026 sí está probado en todo lo que no necesita Azul**, y bastante a fondo:
> ver *"Bases automáticas"* justo debajo. Lo que falta es verlo con clientes saliendo de Azul.

**El alcance cambió.** Se acabó el objetivo de 30 clientes encadenados y desatendidos. Ahora
son **bloques cortos con Dorian presente**. Todo lo que existía solo para sostener el lote
largo se quitó. El motivo no fue el código: `omsserver` se cae, y Azul se degrada con el uso
—medido el 04/09: la misma cuenta dio 13 renovables a las 13:07 y 2 renovables con 13 a
revisar a las 13:24—, así que una corrida de 95-105 minutos cruzaría ese punto y los últimos
clientes saldrían peor que los primeros sin que nada lo indicara.

---

# 🔴 SIN EJERCITAR CONTRA AZUL — lo primero que hay que probar

**Nada de esta lista se ha visto funcionar contra Azul real.** Está escrito y verificado con
lo que se puede verificar sin Azul (ASCII, parseo, compilación del C#, y las pruebas que no
tocan la aplicación), y nada más.

Va en orden: **primero lo que no toca Azul**, después lo que sí. No saltes el orden.

### Sin Azul — se puede probar en cualquier momento

| # | Qué | Cómo | Si falla |
|---|---|---|---|
| 1 | **Los cuatro gates** | `verificar.ps1`, `probar_captura.ps1`, `probar_lista.ps1`, `probar_bases.ps1` | No corras nada más hasta que digan `TODO OK` |
| 2 | **El hueco `[imagen]` en Word** *(etapa 1)* | `diagnostico\probar_word.ps1` — **ya no hace falta cerrar Word**: usa su propio Word y su propio documento | Mira: **una sola** línea `[imagen]`, orden `cuenta → razón social → [imagen] → RENOVABLES`, **cuatro** tablas separadas (trampa 12), canceladas en tabla de una columna |
| 3 | **El CSV real, no uno sintético** *(etapa 2)* | Exporta tu lista de verdad y corre `lote.ps1 -Simular -Lista <ruta>` | Comprueba el número de clientes, las repetidas marcadas, **y que no falte ninguna en silencio**. Fuerza una fila rota a propósito y comprueba que se queja |
| 4 | **La simulación con tu CSV y las bases** | `lote.ps1 -Simular` y `lote.ps1 -VerBases` | La simulación dice en qué base caería cada cliente y dónde está el corte de los 10. `-VerBases` dice qué base está abierta y quién va dentro |

> **Las cuatro comprobaciones del documento activo (etapa 4) ya no existen**, porque ya no hay
> documento activo que comprobar: el lote crea y abre la base él mismo. Solo siguen vivas en
> `agregar_a_word.ps1` cuando se llama **a mano y sin decirle el documento**, que es como se
> usaba antes. La única que llegó a probarse fue la primera, el 06/09/2026.

---

## Bases automáticas (09/09/2026) — qué está probado y qué no

**Probado, en esta máquina, sin Azul:**

- **La memoria del correlativo, entera.** `diagnostico\probar_bases.ps1`, 24 comprobaciones,
  `TODO OK`. Cubre las cinco reglas: el número se aparta antes de que exista el documento; el
  contador no retrocede aunque se borren todos los documentos de la carpeta; si la carpeta va
  por delante gana la carpeta; un documento existente nunca se sobrescribe; un registro roto o
  vacío **aborta en vez de reiniciarse**; y una base abierta cuyo documento desapareció se
  detecta y se avisa. También que una razón social con acento sobrevive el viaje al registro.
- **El ciclo completo de una base en Word**, con tres corridas simuladas y el tope puesto en 2:
  crear la base, escribir un cliente, cerrarla; volver a abrirla en la corrida siguiente,
  añadir el segundo, cerrarla al llegar al tope; y abrir la 035 con el tercero. Comprobado que
  los dos clientes quedan en páginas distintas, que las tablas no se funden, y que una base
  recién creada **no arranca con una página en blanco**.
- **El escritor apuntando a un documento dado** (`-DocObj`), en `probar_word.ps1`, con Word
  abierto y una base real de Dorian dentro. No la tocó.

**Sin probar, y hay que verlo:**

- **Una corrida de verdad**, con Azul: que el lote cree la base, escriba clientes reales y la
  deje cerrada. Nada de la parte de Azul cambió, pero la de Word sí.
- **El salto de un cliente ya escrito** al relanzar la misma lista. La memoria lo reconoce en
  las pruebas; falta verlo dentro de una corrida.
- **El cierre a los 10 de verdad**, con el tope real y no con uno de 2.
- **`lote.ps1` mismo no tiene banco de pruebas sin Azul.** Sus funciones de Word se probaron
  replicadas; el archivo entero solo se ejercita corriéndolo. Es la única pieza de este cambio
  que no tiene gate propio.

### Con Azul — carpeta y documento temporales, nunca `BASES DE DATOS PS1`

| # | Qué | Cómo | Si falla |
|---|---|---|---|
| 5 | **Un cliente suelto** *(etapa 3)* | `buscar_cuenta.ps1 -Cuenta ... -Razon ...` y después `azul_fast.ps1 -SinWord -Cuenta ... -Razon ...` | Mira el CSV: cabeceras `# CUENTA:` y `# CLIENTE:`, el cuadre, los bloques. Y que **Azul quede en Suscripciones** |
| 6 | **`-Fila N` contra el CSV** *(etapa 2)* | `buscar_cuenta.ps1 -Fila 3` | Que saque la cuenta y la razón social correctas del CSV |
| 7 | **LA GUARDA DEL PANEL DEL CLIENTE** *(etapa 3)* | Abre el panel del cliente **tú**, saca tu captura, **déjalo abierto a propósito**, y lanza el cliente siguiente | Debe cerrarlo o avisar claro. **Nunca leer la tabla equivocada.** Antes era el caso que solo tú podías provocar; ahora también lo puede provocar un tope de tiempo que mate la corrida a media captura |
| 8 | **Un bloque de 1, luego de 3** | `lote.ps1 -Limite 1` y `-Limite 3` contra documento temporal | **Cronometra**, y después lee `progreso.txt` |
| 9 | **Los tres topes de la etapa 5** | No se fuerzan: solo se ven con Azul lento. **Míralos en `progreso.txt` después de una corrida mala** | Ver la tabla de abajo |
| 10 | **El contador JAB — cierra el punto 5.4 del plan** | `Select-String 'JAB'` sobre un `progreso.txt` viejo y uno nuevo, **de la misma cuenta** | Es lo que dice qué ganaron los hallazgos 8, 9, 15 y 16, que llevan aplicados y **sin medir**. No hace falta cronómetro: el reloj depende de lo cargado que esté Azul, el número de llamadas no |
| 11 | **El tope de TRAMO de 10 s (etapa 5.4)** | `Select-String 'EXCEDIDO' azul\progreso.txt` tras una corrida | Ver la nota bajo la tabla de los tres topes, más abajo. Un cliente con la tabla de Suscripciones genuinamente vacía TAMBIÉN puede disparar este `ERROR` — es a propósito, no un bug |
| 12 | **La captura automática restaurada** *(revierte parte de 1 y 3)* | Un cliente normal, **sin** `-SinWord`, para que llegue a Word | En `progreso.txt`: `captura: clic en el enlace Cliente en (500,120)`, `captura: panel abierto - Cuenta: ...`, `captura: espera de pantalla estable = N s`, `captura: guardada en ...`, `cierre del panel del cliente = True`. En el CSV: cabecera `# IMAGEN:`. **En el Word: la imagen real en el hueco, no el marcador `[imagen]`** |

> ✅ **Items 5 y 6 probados contra Azul real, 07/09/2026 — antes de que volviera la captura.**
> Esa corrida usó el código sin captura (etapa 3 tal cual, sin la restauración de más abajo),
> así que confirma la lectura y el cuadre, pero **no** dice nada sobre la foto: eso sigue sin
> probarse (item 12). `buscar_cuenta.ps1 -Fila 2` cargó
> la cuenta 507000004 (rejilla de 2 coincidencias, razón social comprobada contra la lista,
> `Seleccionar`, Suscripciones abierta). `azul_fast.ps1 -Cuenta -Razon` (sin `-SinWord`,
> escribiendo Word de verdad) leyó 7 filas: 6 activas + 1 cancelada, cuadre correcto, 0 a
> revisar. Escrito en `BASE 035 PS1 VIRLAN.docx`, sin guardar (falta la foto del cliente y la
> revisión de Dorian).
>
> De paso, primera corrida real de las pausas del mismo día (`PausaLineas`=4 s,
> `PausaLectura`=7 s, ver `docs/decisiones.md` si hace falta el porqué): **82,1 s** de lectura
> interna (`Run()`) para 6 líneas activas, **87,1 s** el script completo. Sin las pausas
> hubiera rondado los 20 s — la diferencia (~62 s) es justo lo que las pausas suman a propósito.
> Cuadra con el cálculo hecho antes de correrlo.

### Los tres topes de la etapa 5, y qué buscar en `progreso.txt`

Ninguno se puede forzar: son techos que solo se tocan cuando Azul va lento. **Con Azul sano
no cambian nada** — por eso no se van a "ver funcionar", solo se va a notar su ausencia de
fallos. Lo que hay que vigilar es que **no aparezcan** estos mensajes:

| Tope | De 800/950/400 a | Condición que espera | Señal de que hizo falta |
|---|---|---|---|
| `Init()` | **5 s** | El puente devuelve el contexto raíz | Que **desaparezca** `ERROR: sin contexto raiz. El puente no ve a Azul` |
| caché de panes | **2 s propios** | Que el árbol exponga los `desktop pane` | `AVISO: cache de panes VACIO` en `progreso.txt` — no falla, va lenta en silencio |
| `topeSel` | **2,5 s** | Fila `checked` **y** botón `enabled` | Que **desaparezcan** `no se marco la fila esperada` y `boton deshabilitado` |
| `CommitCF` | foco + los 400 ms de siempre | Que el **otro** campo gane el foco | `commit: el foco no llego al otro campo en 3 s` |

### Etapa 5.4 — tope de TRAMO de 10 s, búsqueda→Suscripciones

Capa **nueva por fuera** de los tres topes de arriba, no un reemplazo. Vive solo en
`buscar_cuenta.ps1`, sección 6: un reloj de pared de **10 s** que envuelve todo el tramo
desde que `$cargado` queda confirmado (una sola cuenta, ya sea directo o tras rejilla +
`Seleccionar`) hasta que la tabla de Suscripciones deja de crecer. Si se pasa de 10 s, el
script **no asume que la tabla está lista**: `ERROR`, captura (`tramo_10s_tabla` o
`pestana_sin_efecto`), `exit 1` — mismo idioma que cualquier otra guarda del archivo, y
`lote.ps1` ya lo cuenta como `FALLO BUSQUEDA` sin que haga falta tocarlo.

**Cómo se ve en `progreso.txt`:** `Select-String 'EXCEDIDO' azul\progreso.txt` — un mensaje
si venció esperando que la pestaña quedara `selected`, otro si venció esperando que la tabla
estabilizara.

> ⚠ **Riesgo asumido a propósito.** Antes de esta etapa, si la tabla de Suscripciones tardaba
> en pintarse (Azul cargado) o un cliente tenía la tabla genuinamente vacía (0 líneas), el
> script esperaba hasta ~24 s y seguía con un `AVISO` benigno y salida exitosa — no hay forma
> de distinguir "todavía está pintando" de "está vacía de verdad" solo con lo que expone
> `ResumenSuscripciones()`. Con este tope, los dos casos van a chocar casi siempre contra los
> 10 s y salir por `ERROR`/revisar-a-mano en vez del aviso suave de antes. No se intentó
> adivinar cuál de los dos es — el proyecto lo prohíbe expresamente («Nunca adivinar. Ante la
> duda, `REVISAR` con el motivo», `docs/trampas.md`) — así que el tope se aplica parejo, tal
> como lo pidió Dorian. Ninguno de los dos casos (tabla vacía de verdad, o tabla lenta con
> Azul cargado) se ha visto contra Azul real todavía.

### Casos de negocio que nunca se han ejercitado

Vienen de antes de la simplificación y **siguen sin verse**:

- **Un cliente sin renovables** — no debe entrar al documento ni contar para los 10.
- **Un cliente con bloque `A REVISAR`** — sí debe entrar.
- **La guarda de corrida incompleta** (`revisar_csv.ps1` bloqueando el paso a Word) nunca se
  ha disparado de verdad.
- **Un cliente con varias coincidencias** (rejilla), ahora con la razón social viniendo del
  CSV y no del Excel. Es la guarda que impide entrar al cliente equivocado.
- **El aviso de razón social vacía** *(etapa 2)*, que avisa de que esa guarda queda apagada.

---

## Probado contra Azul real — sigue valiendo

Nada de esto lo tocó la simplificación.

**Buscar la cuenta — `buscar_cuenta.ps1`** · 03/09/2026, cuatro cuentas, corridas
independientes con el mismo resultado. Ícono → campo `ID de CF` → `Buscar Ahora` → rejilla →
`Seleccionar` → cliente cargado → pestaña Suscripciones. También los **tres desenlaces** y las
guardas de número truncado, criterios viejos y razón social que no coincide.

**Leer las líneas — `azul_fast.ps1`** · lectura, clasificación y cuadre correctos.
**4,6 s/línea** (18 líneas en 82,6 s, 02/09/2026).

**Corrida completa de 2 clientes encadenados** · 04/09/2026 13:07–13:13, `2 escritos, 0
fallos`, 6,2 min. Incluye un cliente con **12 líneas canceladas** que cuadró.

**El tope de tiempo sobre los hijos** · se disparó contra Azul el 04/09 a las 13:24
(`azul_fast.ps1 paso de 300 s`), mató el PID y el lote siguió. **Sin él la corrida se habría
quedado colgada indefinidamente.**

**El bloque `A REVISAR` se ejercitó una vez**, en la corrida degradada del 04/09: 13 líneas
con su motivo al lado, y **no se escribió nada en Word**, que es lo correcto.

---

## Lo que cambió — las seis etapas

Rama `simplificacion`. Cada etapa deja el sistema funcionando; si paras a la mitad, lo
anterior sigue en pie.

| Etapa | Qué | Commit |
|---|---|---|
| 0 | El plan: `docs/plan-simplificacion.md` | `40af942` |
| 1 | Hueco `[imagen]` en el bloque de Word | `7089f8e` |
| 2 | La lista de órdenes sale de un CSV, no del Excel abierto | `5ddc851` |
| 3 | Fuera la captura del cliente; **se queda el cierre del panel** | `5fb8607` |
| 4 | Fuera las bases automáticas; entra el documento activo | `4e895de` |
| 5.1 | `Init()` espera al puente 5 s | `f20dac0` |
| 5.2 | El tope de selección de fila sube a 2,5 s | `adf1d83` |
| 5.3 | `CommitCF` espera la condición correcta | `579bfa8` |
| 6 | Fuera `Front()`, y la documentación | `881d3b4` |
| 5.4 | Tope de TRAMO de 10 s, búsqueda→Suscripciones | `1784d01` |
| — | Captura automática del cliente de vuelta (**revierte parte de 1 y 3**) | `c4369d9` |
| — | Bases automáticas de vuelta, con memoria real (**revierte la 4**) | *este* |

**Tres decisiones que corrigieron el plan sobre la marcha:**

1. **`MarcoCuentaFull` no se borró: se borró junto con `MarcoCuentaAc`, y todo pasó a la vía
   de panes.** Con la comprobación del panel en cada cliente, la red de seguridad por el árbol
   entero se pagaría en **todas** las corridas para contestar "no hay panel" — lo que
   `docs/trampas.md` prohíbe. El precio: si un panel no colgara de un pane cacheado, la guarda
   no lo vería y la corrida fallaría como fallaba antes. Mal mensaje, nunca datos equivocados.
2. **`CommitCF` no sondea el valor del campo**, como decía el plan. Releer el campo **no
   prueba el commit** —el texto no cambia cuando Siebel lo recoge— así que ese sondeo habría
   salido a los 0 ms y borrado los 400 ms que existen justo para evitar ese fallo. Se sondea
   el **foco**, que es el mecanismo, y los 400 ms se conservan después.
3. **`Init()` devolvía `true` con el caché de panes vacío.** No fallaba: iba lenta en
   silencio, cayendo al recorrido completo del árbol en cada búsqueda de marco. Ahora avisa.

---

## Tiempo — estimado, sin medir

| | Antes (con foto) | Con la foto quitada (06-07/09) | Con la foto de vuelta (08/09) |
|---|---|---|---|
| Cliente de 16 líneas activas | ~3,1 min | ~2,7–2,9 min *(inferencia)* | **de vuelta a ~3,1 min** *(inferencia)* |
| Base de 10 clientes | ~35 min | ~28 min *(inferencia)* | **de vuelta a ~35 min** *(inferencia)* |

> **La captura no tiene un costo fijo: es un sondeo con tres topes (40 s + 30 s + 15 s).**
> En el caso normal sale mucho antes de esos topes, que es de donde sale el ~10-25 s por
> cliente estimado en `docs/plan-simplificacion.md` §5.2. Pero el **peor caso real es ~85 s**
> si los tres topes se agotan (Azul lento), no 10-25 s. Con el tope de `lote.ps1` en 300 s por
> cliente (`-TopeSeg`), sigue habiendo margen, pero es menos margen del que sugiere la
> estimación optimista. **Nada de esto está medido contra Azul real todavía** (ver item 12
> más arriba).

**No llega a 1,5 min y no por poco.** La lectura línea a línea es el 55-65% del tiempo y la
simplificación no la toca: a 4,6 s/línea, 16 líneas son 74 s solo de lectura. Bajar de ahí
exige tocar intocables — el botón `Exportar` está vetado por `REGLAS.md` §4. Detalle completo
en `docs/plan-simplificacion.md` §5.

**Dorian dio los ~28 min por buenos el 05/09/2026:** el objetivo dejó de ser el reloj, es no
hacer el trabajo pesado a mano.

---

## Abierto por el lado de Azul — nada de esto lo arregla la simplificación

**1. Qué mira `Interaccion()`, y qué NO mira. — BLOQUEANTE, sigue abierto.**
Mira el **primer** nodo del árbol cuyo `role_en_US` sea `internal frame` y cuyo `name` empiece
por `"Inicio de Interacción"`. **No mira** el panel `Abrir Ventanas`, ni la banda superior
`Cuenta:`/`Cliente:` (es Chromium, el puente no la ve), ni la `Jerarquía de cuenta`.

**"No hay marco de interacción" NO significa "no hay cliente cargado".** El 04/09 a las 12:54
había cliente cargado —se veía en tres sitios de la pantalla— e `Interaccion()` devolvió
vacío. Las dos guardas que dependen de ella **fallan hacia el lado seguro** (la corrida cae en
la rama de rejilla, que sí comprueba la razón social, o aborta), pero no valen lo que se creía.
**Falta saber por qué el marco deja de exponerse como `internal frame`.** Se cierra con un
volcado puntual de marcos teniendo un cliente cargado y la rejilla abierta.

**2. `Seleccionar` sin efecto.** Visto **una vez** el 04/09 a las 12:54 y no reproducible: la
corrida idéntica a las 13:09 salió bien, con el mismo estado previo. `buscar_cuenta.ps1`
abortó como debe y dejó captura. **Apunta a algo transitorio de Azul, no a estado heredado.**
Si reaparece, toca decidir si conviene reintentar el clic en vez de abortar.

**3. Azul se degrada con el uso.** Medido, no supuesto: misma cuenta, mismo código,
13 renovables a las 13:07 y 2 renovables con 13 a revisar a las 13:24, con Azul **vivo, el
mismo proceso**, `Responding = True`, y **consumiendo un núcleo al 100% de forma sostenida**.
No fue una caída y no fue el código. **Es la razón por la que el alcance cambió.**

> **Segunda medición, 07/09/2026, con las pausas nuevas ya puestas** (`PausaLineas`=4 s,
> `PausaLectura`=7 s, `PausaEscribir`=5 s, `PausaResultado`=10 s — ver `docs/decisiones.md`).
> Bloque de 5 clientes seguidos, mismo Azul sin reiniciar. Los primeros 4 fueron bien (17 a
> 273 s de lectura, según el tamaño). En el 5º, `buscar_cuenta.ps1` tardó **más de 3 minutos**
> en un paso que en los otros cuatro tardó segundos (cliente ya cargado, solo esperando a que
> la tabla de Suscripciones apareciera). En la lectura de ese mismo cliente aparecieron DOS
> fallos nuevos, nunca vistos antes: una línea con "se agoto el tiempo esperando el detalle"
> (agotó los 40 s por intento, dos veces) y otra con **"abortado: la cuenta cambio durante la
> corrida"** — la guarda de la línea 371 encontró un número distinto al esperado en esa fila de
> la tabla. La corrida quedó incompleta y **no se escribió en Word**, que es lo correcto: nada
> malo llegó al documento, solo faltan esas dos líneas por revisar a mano. El tiempo total del
> script (583 s) no cuadra con el tiempo interno reportado (185 s) — hay unos 400 s sin
> explicar fuera del bucle de lectura, todavía sin diagnosticar.
>
> **Lectura de esto:** las pausas nuevas no evitaron la degradación — solo la retrasaron o la
> hicieron más visible al quinto cliente en vez de esconderla. Sigue sin saberse si el límite
> es de TIEMPO de sesión, de NÚMERO de clientes, o de ambos. Encaja con que Dorian ya haya
> fijado el alcance en 3 a 5 clientes por tirada: el quinto es justo donde esta corrida
> empezó a fallar.

**4. Un `Kill()` por tope de tiempo puede dejar Azul sucio.** Si `azul_fast.ps1` muere a media
lectura puede quedar una subventana de detalle abierta, o el marco del cliente. **No se ha
visto ocurrir.** Si el cliente posterior a un `FALLO TIEMPO` arranca sucio, es lo primero que
mirar. `Kill()` solo mata ese proceso, no descendientes; hoy no los hay.

---

## Deuda pendiente

- **Releer completa** la cuenta con la línea `4770000001`, que quedó sin leer cuando Azul se
  cayó el 02/09/2026. Lo parcial está en `azul\salidas\incompletas\`.
- **La verificación contra ADS** (en Chrome) sigue siendo **manual**.
- **`Plan Forzoso`** está extraído pero **no verificado contra pantalla**: la etiqueta
  `Duración del Compromiso` viene del mapa de datos, no de una corrida.
- **La tabla de mensajes de error sigue sin mudarse a `docs/trampas.md`.** Vivía en
  `azul\LEEME.md`, que se retiró del proyecto por estar superado. La tabla no se perdió
  —`git show 4e895de:azul/LEEME.md`, sección *Si algo falla*— pero mientras no se mude,
  quien tenga un error delante no la encuentra donde la busca.

### Resueltas

- **Dónde queda `azul_renovables.csv`** — 06/09/2026, leyendo el código: **junto al script**,
  `azul\azul_renovables.csv`. `CONTEXTO-sesion-03-09-2026.md` §8 lo ubicaba en `salidas\`;
  no rige.
- **`azul_fast.ps1` ¿escribe solo en Word?** — Sí, salvo con `-SinWord`, que es como lo llama
  `lote.ps1`. `LEEME.md` estaba desactualizado.
- **La tolerancia de 10 días RIGE** y el script la aplica (visto en la corrida del 04/09).
- **La contradicción de `REGLAS.md` §5** — resuelta por Dorian el 06/09/2026: autorizó cambiar
  **solo** la frase de los parámetros `-ClienteX`/`-ClienteY`, que dejaron de existir. El resto
  del párrafo se queda: sigue siendo cierto y ahora le sirve a él para saber dónde hacer clic.

---

## Nota de trato

Dorian pide **menos ceremonia**. Cuando el cambio está hecho y `verificar.ps1` da `TODO OK`,
correrlo y enseñar el resultado — no encadenar verificaciones intermedias.

Y al informar, **distinguir siempre lo que está probado contra Azul real de lo que solo está
escrito.**
