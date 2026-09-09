# Mapa de datos — dónde vive cada cosa

> Ábrelo **solo cuando toques la lectura**. Para operar no hace falta.
> Todo lo de aquí está verificado en pantalla, salvo lo que se marca como no verificado.

---

## 1. La lista de órdenes (CSV)

**Desde el 06/09/2026 es un CSV en disco, no el libro abierto de Excel.** Lo exporta
Dorian: hoja `resultData` → *Guardar como* → CSV. Ruta por defecto
`azul\lista_ordenes.csv`, o `-Lista <ruta>`. Vive en `azul\lista.ps1`.

- **Se busca por encabezado, nunca por posición.** Por eso el CSV se exporta **sin borrar
  columnas**: las otras siete (folio, consecutivo, tipo, estatus, fecha, creado por,
  tecnología) sobran y no estorban.
- Hace falta una columna **`Cuenta`** (9 dígitos) y una **`Razón Social`**.
  **Fila 1 = encabezados**, así que `-Fila 6` sigue siendo la sexta fila que se ve en la hoja.
- **Puede haber cuentas repetidas.** El 03/09 la `636000001` salía en las filas 11 y 31.
  La repetida se salta: es la misma base de datos, y duplicarla rompería el conteo de 10.
- **Toda fila descartada se dice en voz alta.** Antes el descarte era silencioso, y con el
  libro vivo eso solo se comía el encabezado. Con un CSV exportado a mano, una cuenta en
  notación científica (`5,95E+08`) o sin el cero de la izquierda desaparecía del lote sin que
  nada lo dijera, y no se notaba hasta que faltaba un cliente en la base.
- **Sin columna `Razón Social` se avisa**: esa columna sostiene la guarda que impide entrar
  al cliente equivocado, y esa guarda es condicional. Sin el dato no se aplica, y callarlo
  sería dejarla apagada en silencio.

**El separador y la codificación se deducen, no se suponen.** Excel escribe el separador de
lista de Windows (`,` o `;`), y la codificación depende de qué opción de *Guardar como* se
use. Ver **trampa 15** en `docs/trampas.md`: UTF-8 sin BOM corrompía la razón social en
silencio, y esa cadena acaba en el encabezado del documento que ve el equipo comercial.

`diagnostico\probar_lista.ps1` prueba todo lo anterior sin tocar Azul, Excel ni Word.
`diagnostico\excel_ordenes.ps1` sigue sirviendo: cuando la exportación salga rara, es lo que
dice qué hay **de verdad** en el libro del que salió el CSV.

> **Por qué se leía el libro abierto, y por qué ya no.** El nombre del archivo trae el número
> de descarga y cambia en cada bajada (`ListaOrdenes (59).xls PNDIENE PS1.xls`), así que una
> ruta fija quedaba obsoleta al día siguiente — de ahí la localización por hoja `resultData`.
> Con el CSV el problema desaparece de otra forma: la ruta es fija porque **Dorian sobrescribe
> el mismo archivo** cada mañana. Lo que se pierde es que la lista deje de estar viva: si
> vuelve a bajar el archivo de órdenes, tiene que volver a exportar.

---

## 2. El formulario de búsqueda (`Encontrar Comunicante Búsqueda CIM`)

El ícono de teléfono con lupa abre **dos marcos Swing**:

| Marco | Geometría de referencia | Papel |
|---|---|---|
| `Encontrar Comunicante Búsqueda CIM` | (282,135) 1075x604 | el formulario de búsqueda, encima |
| `Identificar Cliente` | (282,135) 537x302 | la rejilla de resultados, debajo |

El formulario son **17 campos `text`** de 311x19, `act=SI` y `editable`, en cuatro
bloques: *Suscripción* (Móvil, Número de Vehículo, Código Postal, Colonia),
*Contacto* (Nombre, Segundo Nombre, Primer/Segundo Apellido, Teléfono, Correo
Electrónico), *Cuenta* (ID de la Cuenta, Nombre de la Cuenta) y *Entidades de
Facturación* (ID del Cliente, Nombre Fiscal, RFC, ID del Acuerdo de Facturación,
**ID de CF**).

**`ID de CF` es el último campo de todos**, en (606, 587) como referencia. Se
identifica **por `name`, no por coordenada**.

> **Trampa: dos nodos tienen ese mismo `name`.** La etiqueta (`label`, `act=no`,
> y=570) y el campo (`text`, `act=SI`, `editable`, y=587). Hay que exigir rol `text`
> **y** estado `editable`, o se escribe sobre la etiqueta y no pasa nada.

**Solo la mitad de los campos traen su etiqueta en `name`.** Sí la traen: Móvil,
Nombre, Primer Apellido, Teléfono, Correo, ID de la Cuenta, Nombre de la Cuenta, ID
del Cliente, Nombre Fiscal, ID del Acuerdo e **ID de CF**. Vienen con `name` vacío y
solo se ubican por la etiqueta de encima: Número de Vehículo, Código Postal, Colonia,
Segundo Nombre, Segundo Apellido y RFC. El campo que importa está en el grupo bueno.

**Botones:** `Buscar Ahora` (938,173) 77x19 y `Cancelar` (938,194), los dos `act=SI`.
Aunque el puente los ve y los puede disparar, **se pulsan con clic real** — ver
`docs/decisiones.md`.

### Cómo se sabe que el resultado llegó

Tres señales, ninguna es una espera fija:

1. El marco `Encontrar Comunicante Búsqueda CIM` **desaparece** del árbol.
2. En `Identificar Cliente`, el botón `Seleccionar` (290,411) **gana el estado
   `enabled`**. Sin resultados sus estados son `[focusable,visible,showing]`.
3. El contador de la barra —un `push button` cuyo `name` es el propio número
   (`1+ Registro/s`, `-- Registro/s`…)— **cambia de texto**.

### Los tres desenlaces, y hay que distinguirlos

| Qué pasó | Cómo se detecta |
|---|---|
| Coincidencia única → carga al cliente | aparece el marco **`Inicio de Interacción [N] - <id>`** |
| Varias → rejilla para elegir | un marco con un botón contador **`N Registro/s`** |
| Ninguna | ni marco de interacción ni filas |

> **Que el formulario se cierre NO significa que encontró**: se cierra igual sin
> coincidencias. Creerse esa señal hizo reportar éxito dos veces donde no lo había.

---

## 3. La tabla de Suscripciones

Pestaña **Suscripciones** del marco `Inicio de Interacción`. Tabla de **N × 20 columnas**:

| Columna | Contenido |
|---|---|
| 0 | radio de selección |
| **1** | **número telefónico** |
| 6 | **fecha de alta** — *no* de expiración |
| **7** | **estado** (`Activa` / `Cancelado`) |

**No confundir:** el `Fecha de ve...` del encabezado de cuenta es vencimiento de
**saldo**, no del contrato.

**Selección de una línea:** `addAccessibleSelectionFromContext(tabla, fila * 20)`.
El índice es de **celda, no de fila**. Marcar el radio con `doAccessibleActions`
cambia el widget a `checked` pero **no habilita el botón**; esta vía sí.

**El botón se llama `Ver Productos Asignados`** (no "Productos asignados") y está **al
pie del panel** de Suscripciones, no en la barra de la tabla. El enlace `Ver Productos`
del encabezado es otra cosa.

En la pestaña hay un selector **`Sólo contacto` / `Cuenta` / `Cliente`**; se usa
*Sólo contacto*, que es como Azul lo deja. Dorian confirmó que así está bien. Si
algún día faltan líneas, **es el primer sitio donde mirar**.

---

## 4. El árbol de atributos del detalle

`Ver Productos Asignados` abre una subventana dentro de la misma app, con una tabla
chica tipo Excel: **N × 4 columnas**.

> **El nombre del atributo NO está en `name`, sino en `description` y envuelto en
> HTML**: `<html>Fecha Final del Compromiso</html>`. Hay que quitar las etiquetas.
> El **valor** sí está en `name`. Sin esto, todo sale `sin nodo Plan Móvil`.

| Dato | Fila (etiqueta) | Columna |
|---|---|---|
| Plan | `Plan Móvil` | 3 · ej. `ATT Ármalo Negocios $399` |
| Plan Forzoso | `Duración del Compromiso` | 2 · ej. `24 Meses` |
| Fecha de expiración | `Fecha Final del Compromiso` | 2 |
| MPE | `MPE` | 2 |
| Dispositivo | `Marca (Fabricante)` + `Modelo` | 2 |

Jerarquía: `Plan Móvil → Compromiso → …` — **`MPE` es el último hijo de `Compromiso`**.
`MPE` y `Fecha Final` tienen que caer **dentro del bloque `Compromiso`**; `Plan Forzoso`
lleva la misma contención.

**El orden de los nodos cambia entre líneas. Buscar por nombre, nunca por posición.**

### Verificación obligatoria antes de leer

El título de la subventana identifica la línea:
`Detalles del Producto Asignado: Producto Móvil (8910000001)`.
**Comprobar que contenga el número pedido**, o se corre el riesgo de atribuir los
datos de una línea a otra.

### Casos borde

- **Sin nodo `Compromiso`** → SIM/eSIM sin contrato. Es **RENOVABLE**, se marca
  `(SIM)` junto al número (ver `docs/decisiones.md`).
- **`Marca`/`Modelo` = `N/A`** → equipo del cliente, sin registrar. `SIM y Equipos`
  dice `desconocido`.
- **Líneas `Cancelado`** suelen venir sin plan. No se les abre el detalle.
- **No confundir con el MPE:** `Monto total del RC` (bajo `Controlar`) y `Monto total`
  (bajo `Next Equipo`, financiamiento del equipo) son otra cosa.
- Si aparece **más de un `MPE`** o **más de una `Fecha Final`** en la misma línea →
  **no se elige**: `A REVISAR`.

> **No verificado contra pantalla:** la etiqueta `Duración del Compromiso` para
> `Plan Forzoso` sale del mapa de datos, no de una corrida. Si el resumen reporta
> líneas sin Plan Forzoso legible, correr `diagnostico\jab_attrs.ps1` y ajustar la
> constante `L_DUR` en `azul_fast.ps1`.

---

## 5. El marco de la vista general del cliente

**Desde el 08/09/2026 vuelve a abrirlo el sistema.** Del 06/09 al 07/09/2026 lo abría Dorian
a mano, para su captura; se restauró la captura automática (`Get-VistaCliente` en
`azul_fast.ps1`) porque el mecanismo ya estaba probado contra Azul real y se había quitado
solo por tiempo, no porque fallara. El enlace `Cliente:` sigue estando donde dice
`REGLAS.md` §5, y el ciclo 2 sigue autorizado igual — quien lo ejecuta vuelve a ser el
sistema.

**El marco sí es Swing** y el puente lo ve entero. **Su título lo gobierna todo:**

- mientras carga se llama **`Formulario`** — y se ve como un formulario en blanco con un
  botón `Crear`, que es el que **no se toca**
- cuando termina pasa a **`Cuenta: <nombre del cliente>`**

### El ciclo completo: abrir, esperar, fotografiar, cerrar

**El marco del cliente reemplaza al de la interacción y deja la tabla de Suscripciones fuera
del alcance del puente.** Con él abierto, el lector no la encuentra y la corrida muere con
`no encuentro la tabla de Suscripciones`, que no dice nada del motivo real.

`Get-VistaCliente` corre **después** de leer las líneas de Suscripciones, a propósito: si la
foto fallara, la lectura ya está en el CSV. Hace clic real en el enlace `Cliente:`, espera
hasta 40 s a que el título pase a `Cuenta: ...` (`[Azul]::EsperarMarcoCuenta`), espera además
a que la pantalla deje de cambiar (`Wait-PantallaEstable`, hasta 30 s), guarda el PNG y
comprueba que no sea una captura en blanco o negro (`Test-CapturaValida`) antes de cerrar.

**La guarda de arranque de cada cliente se conserva sin cambios**, porque sigue haciendo
falta: si un `Kill()` por tope de tiempo interrumpe la captura a medio camino, el panel puede
quedar abierto y taparía la tabla del cliente siguiente. Por eso `azul_fast.ps1` sigue
comprobando al arrancar cada cliente si hay uno abierto y lo cierra
(`Close-PanelClienteAbierto`), antes incluso de leer las líneas.

Se cierra con la X de su barra de título, cuya posición se **calcula de la geometría que da el
puente** (`x + ancho - 21`, `y + 9`) — el puente no expone los botones de la barra, pero sí el
marco. `[Azul]::GeoMarcoCuenta()` devuelve esa geometría. Es el mismo cierre
(`Close-MarcoCuenta`) que usan tanto la guarda de arranque como el final de
`Get-VistaCliente` — no hay dos mecanismos.

> **Se busca solo entre los `desktop pane` cacheados, sin recorrer el árbol entero.** El caso
> normal es que **no** haya panel abierto, y el recorrido completo es caro justo en ese caso:
> hay que visitar todo el árbol para poder decir "no está". Con la comprobación en cada
> cliente, esa red se pagaría en todas las corridas. Si algún día un panel no colgara de un
> pane cacheado, la guarda no lo vería y la corrida fallaría como fallaba antes: mal mensaje,
> nunca datos equivocados.

---

## 6. La salida — el CSV

`azul_renovables.csv`, **6 columnas**:
`Numero, Plan, Dispositivo, MPE, PlanForzoso, FechaExpiracion`

Empieza con dos líneas de cabecera:

```
# CUENTA: 595000001
# CLIENTE: DISTRIBUIDORA RIO
```

**Salen de la lista de órdenes, por parámetro** (`azul_fast.ps1 -Cuenta -Razon`, que le pasa
`lote.ps1`). Hasta el 06/09/2026 el nombre lo daba el título del panel del cliente, que se
abría para la foto, y había una tercera línea `# IMAGEN:` con la ruta del PNG.

No desaparecieron con la foto porque **son lo único que dice, mirando el archivo en disco, a
qué cliente pertenece**: sin ellas, reintentar el volcado a Word a mano sería a ciegas. La
cuenta es nueva; antes no estaba.

Son **inertes** para el parser de `agregar_a_word.ps1`, que solo abre bloque con
`# NOMBRE (n)` y solo toma como fila lo que empieza con comilla. Un CSV viejo con `# IMAGEN:`
sigue funcionando: esa rama no se tocó, y si llega una ruta la imagen se sigue insertando.

Y viene en bloques separados por líneas que empiezan con `#`:
`# RENOVABLES (n)`, `# NO RENOVABLES`, `# A REVISAR` (solo si hay),
`# CANCELADAS EXCLUIDAS`, y el cuadre final.

**No hay columna Estado ni Nota:** el bloque en el que cae la línea *es* el estado, y
el matiz va entre paréntesis junto al número — `8110000001 (SIM)`,
`8120000001 (múltiples MPE)`.

`agregar_a_word.ps1` cuenta las renovables leyendo el separador `# RENOVABLES (n)`,
porque ya no hay columna Estado que consultar.

> **Resuelto el 06/09/2026 leyendo el código.** El CSV queda **junto al script**,
> `azul\azul_renovables.csv`. Lo fija `azul_fast.ps1`
> (`Join-Path $PSScriptRoot "azul_renovables.csv"`) y es lo que leen `lote.ps1` y
> `agregar_a_word.ps1`. `CONTEXTO-sesion-03-09-2026.md` §8 lo ubicaba en `salidas\`: **no
> rige**, y ese archivo está entre los superados.

---

## 7. La regla de clasificación

Es regla de **negocio**, no de seguridad: por eso vive aquí y no en `REGLAS.md`.

**Filtro previo:** las **canceladas no se consideran** (columna 7). No entran a ninguna
tabla; se **descartan antes de abrir su detalle** —no se les toca la subventana— y solo
se listan sus números al final para poder cuadrar.

Siendo `fin` = `Fecha Final del Compromiso`:

| Condición | Resultado |
|---|---|
| `fin` ya pasó, sin importar cuánto hace | **RENOVABLE** |
| `fin` **antes de** `hoy + 3 meses + 10 días` | **RENOVABLE** |
| `fin` **en o después de** `hoy + 3 meses + 10 días` | **NO RENOVABLE** (tabla aparte, no se omite) |
| Sin nodo `Compromiso` | **RENOVABLE** — SIM/eSIM sin contrato, se marca `(SIM)` |

**`Suspendida` se trata igual que `Activa`** (Dorian, 09/09/2026): sigue exactamente la
misma tabla de arriba (fecha, tolerancia, sin nodo `Compromiso`), y se marca `(Suspendida)`
junto al número para que quede trazable — igual que `(SIM)`, y combinable con él si una
línea es las dos cosas a la vez. **Escrito y verificado sin Azul delante; no probado
contra Azul real todavía** (Dorian pidió explícitamente no probarlo antes de este push).

**Hacia atrás no hay límite**: una línea activa cuyo contrato venció hace dos años sí
es renovable. Hacia adelante la ventana es de 3 meses **más 10 días de tolerancia**.

**La tolerancia de 10 días** (02/09/2026, confirmada por Dorian el 04/09/2026). La
ventana no es un muro: si `fin` se pasa de los 3 meses por **menos de 10 días, sigue
siendo renovable**; a los 10 días o más, ya no. Dorian lo pidió textual —*"se un poco
flexible"*— al toparse con líneas que vencían 3 días después del corte.

> `PROYECTO-AZUL.md` (superado) decía *"la frontera es una sola: hoy + 3 meses"*, sin
> tolerancia. Es la versión del 01/09, anterior a este cambio. **No rige.**

**A REVISAR** — solo ambigüedad real o fallo de lectura, y ahí **no se adivina**:
múltiples `MPE`, múltiples `Fecha Final`, fecha ilegible, atributos que no cargaron,
o un estado de línea que no sea `Activa`, `Suspendida` ni `Cancelado` (que **no se
excluye**: excluir por un valor desconocido rompería el cuadre; entra como `REVISAR`
con el valor crudo).

`MPE = 0` y `MPE = -1` **se reportan tal cual**: sí tienen sentido en la operación.

**Matiz informativo** que no cambia el bucket, para las no renovables:
*un año o más hacia adelante* → renovada recientemente, verdaderamente no renovable
(la fecha basta, no hace falta leer la de inicio); *menos de un año* → fuera de la
ventana, pero se acerca.

---

*Procedencia: `PROYECTO-AZUL.md`, `azul\CONTEXTO-sesion-03-09-2026.md` (§4),
`azul\CONTEXTO-sesion-02-09-2026.md`, `azul\LEEME.md`, y las memorias
`azul-donde-viven-los-datos`, `azul-lectura-tecnica`, `azul-regla-renovacion`,
`azul-vista-cliente`.*
