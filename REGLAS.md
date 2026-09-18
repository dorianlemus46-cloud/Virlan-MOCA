# REGLAS DE SEGURIDAD — Proyecto Azul

> **Este archivo no se modifica sin autorización expresa de Dorian.**
> Ni para "aclarar", ni para "actualizar", ni para reflejar un cambio en los scripts.
> Si una regla estorba, se pregunta y se espera respuesta. No se reescribe.

---

## 1. La regla base

**Azul es de solo lectura en cuanto a datos del cliente.**

Está permitido navegar y consultar. Está **PROHIBIDO** guardar, aprobar, descartar,
eliminar, modificar o exportar cualquier campo o registro del cliente.

Azul es producción, sobre cuentas de clientes reales. Las decisiones de aprobar o
descartar una renovación las toma Dorian, después, verificando contra ADS.

---

## 2. Acciones prohibidas

- Guardar cualquier formulario.
- Aprobar, descartar o eliminar cualquier registro.
- Modificar cualquier campo: escribir, borrar o pegar en un campo que no sea
  `ID de CF` dentro del ciclo 3 de abajo.
- Usar los botones **`Exportar`** de Azul.
- Hacer clic en cualquier botón, ícono, enlace o pestaña que no esté en la lista
  blanca del punto 3.
- Volcados amplios del árbol de accesibilidad. El árbol contiene pantallas con
  **datos personales de clientes**: se apunta solo a la tabla de Suscripciones y a
  la de productos asignados. Los diagnósticos que recorren todo (`jab_cliente.ps1
  -Todo`) se usan puntualmente y no se dejan corriendo.

---

## 3. Ciclos autorizados — y solo estos

**Ciclo 1 — Leer una línea.**
Seleccionar la línea → **`Ver Productos Asignados`** → leer → **`Cerrar`**.

**Ciclo 2 — Vista general del cliente.** *(autorizado por Dorian el 02/09/2026)*
Clic real en el enlace `Cliente:` → esperar a que el marco cargue → leer o
fotografiar → cerrar el marco con **la X de su barra de título**.
Es uno de los dos puntos donde el sistema toca la máquina: sube Azul al frente y da
un clic real de mouse. El cursor se devuelve a donde estaba.
**Dentro de esa vista hay campos editables y un botón `Crear`: no se toca nada más
que la X de cerrar.**

**Ciclo 3 — Buscar una cuenta.** *(autorizado por Dorian el 03/09/2026; el cierre final
del marco de interacción, autorizado el 09/09/2026)*
Clic real en el ícono de teléfono con lupa → escribir el número de cuenta en el campo
**`ID de CF`** → **`Buscar Ahora`** → si hay rejilla, marcar la **primera fila**,
comprobar contra la razón social del Excel → **`Seleccionar`** → clic real en la
pestaña **`Suscripciones`** → una vez leída la línea (o líneas) y cerrado el panel del
cliente, clic real en la X de la barra de título del marco **`Inicio de Interacción`**,
hasta que las bandas `Cuenta:` y `Cliente:` de arriba vuelvan a quedar vacías, antes de
pasar al siguiente cliente.
Llena un formulario de consulta. No modifica ningún dato del cliente.

> **Por qué hace falta este cierre.** Sin él, "Inicio de Interacción" queda listado en
> el panel `Abrir Ventanas` de Azul y el cliente sigue cargado al pasar al siguiente,
> que es justo lo que el sistema debe evitar. Confirmado el 09/09/2026: cerrarlo con la
> X, sin haber escrito nada en sus campos (`Razón 1/2`, `Resultado`), no pidió guardar
> ni mostró ningún aviso — un solo clic bastó para volver a `No hay Cliente Actual`.
> `Búsqueda: Contacto y Suscripción` **no se cierra**: es la vista base a la que se
> quiere volver, no una ventana más del cliente.

### Lista blanca — lo único que se puede pulsar

`Ver Productos Asignados` · `Cerrar` · `Buscar Ahora` · `Seleccionar` ·
la pestaña `Suscripciones` · el enlace `Cliente:` · el ícono de teléfono con lupa ·
la X de la barra de título del marco del cliente · la X de la barra de título del
marco de interacción (`Inicio de Interacción...`) · **`Descartar`, y solo en el aviso
que sale al cerrar una de esas dos subventanas — condiciones exactas en el punto 4.1** ·
el botón contador **`20+ Registro/s`** de la tabla de Suscripciones, **`Guardar`** en la
ventana que ese botón abre, y el cierre de esa ventana — **solo en el caso del punto 4.2**.

---

## 4. Nombres de botón que NUNCA se pulsan

Ningún botón cuyo nombre sea o contenga:

| Nombre | Dónde aparece |
|---|---|
| **`Crear`** | En la vista general del cliente (ciclo 2) y en el formulario en blanco que aparece mientras esa vista todavía no ha cargado. |
| **`Exportar`** | En la interfaz de Azul. Decisión tomada: no se usan; se mantiene el ciclo 1. |
| **`Guardar`** | En cualquier pantalla. **Una sola excepción, en el punto 4.2.** |
| **`Aprobar`** | En cualquier pantalla. |
| **`Descartar`** | En cualquier pantalla. **Una sola excepción, detallada justo abajo.** |
| **`Eliminar`** | En cualquier pantalla. |

### 4.1 La única excepción: `Descartar` en el aviso al cerrar una subventana

*(autorizada por Dorian el 11/09/2026)*

Se permite pulsar **`Descartar`** en **un solo sitio y bajo todas estas condiciones a la
vez**. Si falta una, no se pulsa.

1. Es el **cuadro de aviso que Azul saca en el centro de la pantalla** al intentar cerrar
   una subventana del cliente: el panel `Cuenta: <nombre>` o el marco
   `Inicio de Interacción`.
2. Ese aviso apareció **como consecuencia directa de nuestro propio clic** en la X de la
   barra de título de una de esas dos subventanas, en el mismo momento. Un aviso que ya
   estuviera en pantalla antes no cuenta: ese se deja como está.
3. Lo que se descarta es **el borrador de interacción que creó nuestra propia navegación**,
   el que marca el asterisco en el título. **No** se descarta ningún registro del cliente,
   ninguna renovación, ningún dato guardado y ninguna decisión de negocio.
4. Antes de pulsar se **anota al progreso el texto del aviso**, para que quede por escrito
   qué se aceptó y cuándo.
5. Tras pulsarlo se **sigue cerrando las subventanas que queden**, hasta que las bandas
   `Cuenta:` y `Cliente:` de arriba vuelvan a quedar vacías.

Fuera de eso, `Descartar` sigue prohibido en **cualquier** pantalla, formulario, barra de
herramientas o cuadro de diálogo, incluso si el texto se parece. Si el aviso dice algo
distinto de lo esperado, ofrece otros botones, o aparece en cualquier otro momento: **no se
pulsa nada**, se deja la pantalla como está y se avisa a Dorian.

> **Por qué se autorizó.** Medido el 11/09/2026: sin pulsar ese aviso el cierre de la
> subventana falla en silencio, el cliente siguiente hereda la subventana del anterior y
> pasa a costar 120.777 consultas al puente y 38 s, contra 13.200 y 6 s de un cliente
> limpio. El 10/09/2026 el mismo problema costó un cliente entero. Hasta esa fecha Dorian
> pulsaba el aviso a mano, lo que obligaba a estar delante durante toda la corrida.
> Esto **corrige** la nota del ciclo 3, que daba por hecho que cerrar con la X nunca pedía
> confirmación: eso se comprobó el 09/09/2026 con los campos vacíos, y no vale cuando la
> navegación ya dejó el borrador marcado con asterisco.

### 4.2 La segunda excepción: `Guardar` en la ventana de `20+ Registro/s`

*(autorizada por Dorian el 18/09/2026: "agrega el botón guardar solo en este caso, como
excepción")*

La tabla de Suscripciones enseña **20 filas como mucho**. Cuando el cliente tiene más, las
que faltan no se leen, y el cliente puede acabar mal clasificado o fuera de la base. Arriba,
en la banda del cliente, Azul dice cuántas tiene de verdad: el recuadro del teléfono,
**`N Inalámbricos`**.

Se permite, **en un solo sitio y bajo todas estas condiciones a la vez**:

1. **Los números no empatan.** El número de `N Inalámbricos` es **mayor** que las filas que
   trae la tabla. Si empatan, no se pulsa nada.
2. **Son 50 o menos.** Con **más de 50 inalámbricas el cliente se salta entero**: no se pulsa
   el botón, no se lee y no entra a la base. Queda anotado con el motivo.
3. Se pulsa **`20+ Registro/s`**, el botón contador de la tabla de Suscripciones del cliente
   que se acaba de cargar, y nada más de esa fila de botones (`Ver Todo` sigue sin tocarse).
4. En la ventana que se abre **como consecuencia directa de ese clic** se escribe **solo la
   cantidad de líneas**: el número de `N Inalámbricos`. Ningún otro campo se toca.
5. Antes de pulsar `Guardar` se **anota al progreso** el texto de esa ventana y el número
   escrito.
6. Se pulsa **`Guardar`** en esa ventana y se **cierra**. Después se vuelve a leer la tabla.

Fuera de eso, `Guardar` sigue prohibido en **cualquier** pantalla. Si la ventana no es la
esperada, trae otros campos o botones, o aparece sin nuestro clic: **no se pulsa nada**, se
deja la pantalla como está y se avisa a Dorian.

> **Cómo es la ventana (vista y probada el 18/09/2026 con un cliente de 24 inalámbricas y 20
> filas).**
> Es un cuadro de diálogo **aparte**, una ventana suelta en el centro de la pantalla, no un
> marco dentro de Azul. Por eso no sale en la foto de la ventana principal. Se titula
> `Inicio de Interacción [N] - <id> > Cambiar los registros máximos`, dice *"El Límite de
> usuario debe ser menor al límite del sistema."*, trae **un solo campo**,
> `Límite del usuario:` (estaba en 20), el texto `Límite del sistema: 200` y tres botones:
> `Guardar y buscar`, **`Guardar`** y `Cancelar`. **Solo se pulsa `Guardar`**; los otros
> dos no. Al pulsarlo **la ventana se cierra sola** y la tabla se recarga en un par de
> segundos con las filas nuevas; el botón contador pasa a decir `24+ Registro/s`.
>
> **El `N Inalámbricos` cuenta también las canceladas.** Con el límite en 24 salieron 17
> activas y 7 canceladas, que suman 24. La comparación es, por tanto, contra **todas** las
> filas de la tabla, no solo contra las activas. Medido en un solo cliente.
>
> **Sin confirmar:** si el límite se queda guardado para los clientes siguientes. Es un
> límite "del usuario", así que lo esperable es que sí.

Tampoco se pulsan los otros tres íconos de la fila del teléfono con lupa —
persona+lupa, dos personas+lupa y lentes— ni el campo de texto suelto a su izquierda.
No forman parte de ningún ciclo autorizado.

---

## 5. Coordenadas clavadas — las dos únicas del proyecto

La banda de arriba (`Cuenta:` / `Cliente:` / los cuatro íconos) y el panel izquierdo
son un **navegador Chromium incrustado** (`Chrome_RenderWidgetHostHWND`, cubre
`(8,31)-(1058,631)`). El Java Access Bridge **no los ve**: no hay geometría que
preguntar, así que estas dos posiciones están medidas a mano.

| Qué | Coordenada | Cómo se ajusta |
|---|---|---|
| Enlace `Cliente:` | **(500, 120)** relativo a la esquina de la ventana | en `azul_fast.ps1` *(vuelve a usarla un script desde el 08/09/2026: la captura automática del cliente se restauró; cambio de esta celda autorizado por Dorian, 08/09/2026)* |
| Ícono teléfono+lupa | **(791, 121)** relativo a la esquina de la ventana, caja de ~30x20 px | en `buscar_cuenta.ps1` |

> **Cambio del 04/09/2026, autorizado por Dorian: el enlace `Cliente:` pasó de (529,120) a
> (500,120).**
>
> El enlace **arranca siempre en x≈491** y se extiende hacia la derecha **según el largo del
> nombre del cliente**. Medido sobre dos capturas con cliente cargado: con `CLIENTE EJE...` el
> subrayado llegaba a x≈568 y 529 caía dentro; con `ACME` termina en x≈520 y **529 caía 9 px
> fuera, sobre gris**. Ése era el motivo real de que la foto se perdiera: no el tamaño de la
> ventana, sino los **clientes de nombre corto**. 500 cae dentro del primer carácter del
> nombre sea cual sea su largo.

> ### ⚠ Advertencia
>
> **Las dos fueron medidas con Azul en 1366x768.** Si la ventana cambia de tamaño o
> de posición relativa, dejan de acertar y el clic cae en otro sitio.
>
> Un clic que falla **no es inofensivo**: cae sobre lo que haya en ese píxel.
>
> Por eso el sistema nunca da por bueno un clic clavado: comprueba el efecto.
> Si el marco `Cuenta:` no aparece, la corrida lo dice y **se queda sin imagen —
> nunca mete una foto equivocada**. Si el resultado no se confirma, se aborta.
>
> **Volver a medirlas antes de asumir que siguen bien** cada vez que Azul haya
> cambiado de tamaño. Se mide con `diagnostico\zoom.ps1` o capturando la ventana.
> Ya hubo un caso en que la ventana se des-maximizaba sola entre medir y pulsar
> (ver trampa 7 en `docs/trampas.md`).

**Todo lo demás NO se clava.** Cualquier cosa que el puente vea —botones Swing,
marcos, pestañas, filas— se le pregunta al puente **en cada corrida** y se pulsa
sobre la geometría que devuelve. Las coordenadas que aparecen en `docs/mapa-datos.md`
para `ID de CF` (606,587), `Buscar Ahora` (938,173) o `Seleccionar` son **medidas de
referencia para orientarse, no valores que el código deba usar**.

---

## 6. Ante cualquier duda: abortar y preguntar

**No se adivina y no se hace clic "a ver qué pasa".**

- Si no está claro que una acción cae dentro de un ciclo autorizado → **no se hace,
  se pregunta.**
- Si un dato no se puede leer bien, o hay más de un `MPE` o más de una
  `Fecha Final` en la misma línea → **no se elige**: va a `A REVISAR` con el motivo.
- Si la condición que confirma un paso no se cumple (el marco no apareció, la
  pestaña no quedó `selected`, el número tecleado no coincide) → **se aborta**, no
  se sigue de largo.
- Si una corrida quedó incompleta → **se pregunta antes de escribir en Word**. Meter
  una tabla incompleta en el documento del cliente es peor que esperar.
- ~~Si hubiera más de un documento de Word abierto → **preguntar cuál** antes de
  escribir.~~ **Superada el 09/09/2026, autorizada por Dorian el 10/09/2026.** El lote
  crea la base, la abre, escribe y la cierra él mismo: no hay documento elegido a mano
  que confundir, y **Word no se abre a mano**. Sigue vigente en un solo caso: llamar a
  `agregar_a_word.ps1` suelto y sin decirle el documento, que es como se usaba antes.
- "Sin datos" no siempre es duda: la línea activa sin bloque `Compromiso` es una
  SIM/eSIM sin contrato, un caso conocido y clasificado (ver `docs/decisiones.md`).

---

*Procedencia: `PROYECTO-AZUL.md` (§ REGLA DE SEGURIDAD, Decisiones tomadas, Vista
general del cliente), `azul\CONTEXTO-sesion-03-09-2026.md` (§2, §3), `azul\CONTEXTO-sesion-02-09-2026.md`,
`azul\LEEME.md`, y las memorias `azul-lectura-tecnica`, `azul-vista-cliente`,
`azul-regla-renovacion`, `azul-flujo-por-cliente`.*
