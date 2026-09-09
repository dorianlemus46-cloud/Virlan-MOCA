# Plan de simplificación — del lote largo al bloque corto

> ## ✅ APLICADO — seis etapas, 06/09/2026, rama `simplificacion`
>
> Este documento es **el plan, no el resultado**. Se conserva porque explica *por qué* el
> sistema quedó como quedó. Para saber qué está probado y qué no, **`ESTADO.md`**.
>
> **Tres puntos del plan se corrigieron al aplicarlos. No sigas estos tal como están escritos
> abajo:**
>
> | Dónde | El plan decía | Por qué está mal |
> |---|---|---|
> | §3.1, `CommitCF` | *"sondea el valor del campo hasta que cuadre"* | **Releer el campo NO prueba el commit** — el texto no cambia cuando Siebel lo recoge. Ese sondeo sale a los 0 ms y borra los 400 ms que existen justo para evitar ese fallo. Se sondea el **foco**, y los 400 ms se conservan después |
> | §2.1, `MarcoCuentaFull` | *"se borra"* — y `MarcoCuentaAc` se queda | Se borraron **las dos**: con la comprobación del panel en cada cliente, la red por el árbol entero se pagaría en todas las corridas para contestar "no hay" |
> | §3.2, `Init()` | subir el tope a 5 s | Hacía falta además **plazo propio para el caché de panes** (compartían presupuesto) y **avisar cuando queda vacío**: `Init` devolvía `true` igual y la corrida iba lenta en silencio |
>
> El detalle de las tres está en `ESTADO.md`, *Lo que cambió*.

> **Escrito el 05/09/2026, SIN Azul delante.** Todo lo que hay aquí sale de leer el código y
> los documentos del repo. Nada se ejecutó, nada se midió. Cada afirmación va marcada:
>
> - **[código]** — lo comprobé leyendo el archivo. Cito archivo y línea.
> - **[registrado]** — lo dice `ESTADO.md` / `docs/bitacora.md`, medido contra Azul en su día.
> - **[inferencia]** — lo deduzco. Puede estar mal. No lo trates como dato.
> - **[no se puede saber sin Azul]** — hace falta una corrida para cerrarlo.
>
> Este archivo es para ejecutarlo en la oficina, en otra sesión. Está escrito para que se
> pueda seguir sin tener delante la conversación en la que salió.

---

## 0. El resumen en cinco líneas

1. Quitar la foto del cliente **no borra `captura.ps1`**: ese archivo es el teclado, el clic
   real y la guarda de ventana minimizada de la búsqueda. Sin él la búsqueda muere.
2. Hay que **conservar la mitad del ciclo de la foto**: la que *cierra* el panel del cliente.
   Vas a abrirlo tú a mano, y si queda abierto tapa la tabla de Suscripciones.
3. El desalineo del correlativo se resuelve **quitándole al sistema la potestad de nombrar**:
   escribe en el documento que tú tengas activo. Tú lo bautizas, como siempre.
4. Excel → CSV toca **dos** lectores casi idénticos, no uno. El segundo es `-Fila` de
   `buscar_cuenta.ps1`, y si se rompe se pierde la razón social — y con ella la guarda que
   impide entrar al cliente equivocado.
5. **El tiempo no llega a 1,5 min con estos cambios.** Estimo ~2,7–2,9 min por cliente de 16
   líneas. Lo que falta para 1,5 está dentro de la lectura línea a línea, y ahí sí hay que
   tocar intocables. Detalle en la sección 5.

---

## 1. Los cambios, uno a uno

### CAMBIO A — fuera la captura del cliente

#### Qué se quita, en archivos y funciones

**`azul/azul_fast.ps1`** (todo lo de abajo es **[código]**, con línea de la versión actual):

| Qué | Dónde | Nota |
|---|---|---|
| `param -ClienteX` / `-ClienteY` | 16–17 | la coordenada clavada (500,120) |
| `param -SinCaptura` | 19 | |
| `param -SoloCaptura` | 22 | |
| C# `const CTA_PREFIX` | 218 | |
| C# `MarcoCuentaFull()` | 236–238 | |
| C# `EsperarMarcoCuenta(int)` | 255–266 | **la espera de 40 s** |
| función `Get-VistaCliente` | 642–699 | el ciclo abrir / esperar / fotografiar / cerrar |
| rama `if ($SoloCaptura)` | 706–713 | |
| llamada a la captura | 724–730 | |
| bloque de cabeceras `# CLIENTE:` / `# IMAGEN:` | 735–740 | ver *qué se pierde* |
| líneas `"Cliente: ..."` / `"Imagen: ..."` | 743–744 | |

**`azul/captura.ps1`**: `Wait-PantallaEstable` (335–381) y `Test-CapturaValida` (386–413)
quedan **sin llamadas**. → **No las borres.** Motivo en la sección 8, punto 6.

**`azul/agregar_a_word.ps1`**: el bloque de imagen (163–178) deja de recibir ruta.
**No se borra**: se le añade la rama del hueco. Ver más abajo.

#### Qué NO se quita, aunque parezca parte de lo mismo

Esto es lo que más importa de todo el cambio A.

| Qué | Dónde | Por qué se queda |
|---|---|---|
| `. captura.ps1` en `azul_fast.ps1` | 619 | `Get-AzulHwnd` (703) vive ahí, y con él la guarda de Azul minimizado |
| C# `MarcoCuentaPanes()` | 223–234 | detecta un panel de cliente abierto |
| C# `MarcoCuentaAc()` / `MarcoCuenta()` | 239–247 | idem |
| C# `GeoMarcoCuenta()` | 250–254 | de ahí sale la X de la barra de título |
| C# `EsperarSinMarcoCuenta(int)` | 267–274 | confirma que el panel cerró |
| función `Close-MarcoCuenta` | 624–637 | **la mitad que se queda** |
| `Invoke-AzulClick` / `Get-AzulRect` | captura.ps1 | los usa `Close-MarcoCuenta` |

**Y se añade una línea:** en `[Azul]::Run()`, justo al lado del `CerrarDetalle()` que ya hay
en la línea 318, comprobar si hay un marco `Cuenta: ...` abierto y cerrarlo (o abortar con
mensaje claro). Hoy eso lo hacía `Get-VistaCliente` al empezar (652–655) **[código]**; si se
borra sin sustituto, nadie lo hace.

> **Por qué esto es el punto crítico del plan.** `docs/mapa-datos.md` §5: *el marco del
> cliente reemplaza al de la interacción y tapa la tabla de Suscripciones*. Hoy el sistema
> es el único que abre ese panel, y lo cierra siempre (el `finally` de 689–693). A partir de
> este cambio **lo vas a abrir tú**, para tu captura. Si lo dejas abierto —o si Azul tarda en
> cerrarlo— el cliente siguiente arranca con la tabla tapada. El código que protege contra
> eso es justo el que parece que sobra. **[código] + [inferencia]** sobre tu flujo manual.

#### Qué se gana

- Desaparece la coordenada clavada (500,120) y con ella toda la clase de fallo "el clic cayó
  9 px fuera sobre gris" (`REGLAS.md` §5). **[registrado]**
- Desaparece `EsperarMarcoCuenta(40)`: el modo de fallo que costó la corrida de las 13:24
  (`el panel del cliente no abrio en 40 s`) ya no existe. **[registrado]**
- Desaparece el suelo fijo de `Wait-PantallaEstable`: 250 ms × 6 confirmaciones = **1,5 s
  obligatorios** aunque Azul ya estuviera quieto (captura.ps1:348, 351). **[código]**
- Tiempo: **10–25 s por cliente**. **[inferencia]** — el desglose y por qué no lo sé mejor,
  en la sección 5.

#### Qué se pierde

1. **La foto, que la pones tú.** Aceptado.
2. **El nombre del cliente tal como lo ve Azul.** Hoy sale del título del marco
   (`Cuenta: ACME SA DE CV`) y viaja en el CSV como `# CLIENTE:` (azul_fast.ps1:737).
   `agregar_a_word.ps1` lo usa como **título de repuesto** cuando no le pasan `-Cuenta`/`-Razon`
   (línea 156) **[código]**. Con el lote no se nota —siempre pasa los dos (lote.ps1:303)— pero
   corriendo `azul_fast.ps1` suelto el bloque saldría titulado *"Lectura de suscripciones -
   fecha"*.
   Y hay una pérdida más silenciosa: era el **único sitio del proyecto donde el nombre que
   dice Azul y la razón social del Excel acaban en el mismo archivo**. Nadie los compara hoy,
   pero al desaparecer se pierde la posibilidad.

   **Recomendado (3 líneas, opcional pero yo lo haría):** que `azul_fast.ps1` acepte
   `-Cuenta` y `-Razon` y escriba `# CLIENTE:` a partir de ellos, y que `lote.ps1` se los pase
   junto a `-SinWord` (lote.ps1:267). Así el CSV en disco sigue diciendo de quién es, que es
   lo que hace falta para reintentar `agregar_a_word.ps1` a mano y para auditar después.
3. **El reintento barato de la foto** (`-SoloCaptura`). Ya no aplica.

#### El hueco en Word — condición innegociable

El encabezado no se toca. Orden actual **[código]**, `agregar_a_word.ps1` 152–178:

```
595000001                 <- Add-Parrafo $Cuenta -Negrita        (152-153)
DISTRIBUIDORA RIO         <- Add-Parrafo $Razon  -Negrita        (154)
[imagen]                  <- InlineShapes.AddPicture            (163-171)
RENOVABLES - 13           <- titulos del bloque                  (191)
...tablas...
```

**El cambio es de una rama, no de estructura.** Donde hoy dice
`if ($Imagen -ne "" -and (Test-Path ...))` se deja igual, y en el `else` final (176–178) —el
que hoy solo imprime *"Sin imagen"*— se añade `[void](Add-Parrafo $d "")`. Un párrafo vacío,
sin negrita, exactamente donde iba la foto.

- Si algún día pasas `-Imagen ruta.png`, **la foto se sigue insertando**: el código no se
  borra. Esa es la reversibilidad del cambio A entero.
- Si prefieres un blanco *visible* al que apuntar, pon `"[imagen]"` en vez de `""`: al pegar,
  Word sustituye la selección. Es un cambio de una palabra. Decídelo tú, es tu documento.
- **A comprobar en la prueba 5** (sección 9): que salga **exactamente una** línea en blanco,
  ni cero ni dos. `Add-Parrafo` hace `MoveEnd(1,-1)` sobre un párrafo recién creado y ese
  caso concreto no lo he podido ejecutar. **[no se puede saber sin correrlo]** — pero se
  comprueba sin Azul, contra un .docx de prueba.

#### Reversible

Sí, y limpiamente. Es un bloque contiguo (`Get-VistaCliente` + su llamada) más dos parámetros.
**Antes de empezar: `git commit` del estado actual.** Con eso, revertir es `git revert` o
copiar `Get-VistaCliente` de vuelta. El bloque de imagen de Word ni siquiera se toca.

#### Orden

**Tercero.** Después de Word y de CSV. Va después porque toca el archivo más delicado de los
tres y conviene llegar a él con las dos etapas anteriores ya probadas.

---

### CAMBIO B — fuera Excel, entra CSV

#### Qué se quita

| Archivo | Qué | Línea |
|---|---|---|
| `lote.ps1` | `Get-Ordenes` entera (COM + hoja `resultData` + encabezados) | 103–143 |
| `buscar_cuenta.ps1` | `Get-OrdenDeExcel` entera | 644–675 |

Son **dos lectores casi idénticos** **[código]**: el bucle que busca los encabezados `Cuenta`
y `Raz*n Social` está duplicado línea por línea en los dos archivos. Es la duplicación que
este cambio permite matar de paso.

#### Qué entra

Un archivo nuevo, **`azul/lista.ps1`**, con el lector del CSV. Lo cargan por `dot-source`
`lote.ps1` y `buscar_cuenta.ps1`, igual que hacen hoy con `captura.ps1`.

Dos funciones:

- `Get-ListaOrdenes -Ruta <csv>` → la lista completa, ya sin duplicados. La usa `lote.ps1`.
- `Get-OrdenDeLista -Ruta <csv> -Fila N` → una sola. La usa `buscar_cuenta.ps1 -Fila N`.

**Contrato del CSV.** Lo generas con *Guardar como* sobre la hoja `resultData` **sin borrar
columnas**: el lector busca por encabezado, igual que hoy, así que las otras siete no
estorban. Solo necesita una columna `Cuenta` y una `Razón Social`.

**Cinco detalles del lector que no son opcionales:**

1. **Se busca por encabezado, no por posición.** Es lo que hace hoy (lote.ps1:115–119) y es lo
   que te ahorra editar el CSV a mano. **[código]**
2. **El nombre de la columna lleva acento y este archivo tiene que quedar en ASCII puro**
   (trampa 1). No se puede escribir `$fila.'Razón Social'`. Hay que recorrer
   `$fila.PSObject.Properties` y comparar el nombre con `-like 'Raz*n Social'`, que es
   exactamente el truco que ya usan los dos lectores actuales (lote.ps1:118). **[código]**
3. **El separador puede no ser la coma.** Excel escribe el separador de lista de Windows, y
   en configuración regional española/mexicana eso puede ser `;`. `Import-Csv` supone `,`. Si
   no coincide, el CSV se lee como **una sola columna** y el lote dice *"0 clientes"* sin más
   explicación. Detectarlo es barato: leer la primera línea, contar `,` contra `;`, y pasar
   `-Delimiter`. **[inferencia]** — no puedo saber tu configuración regional desde aquí.
4. **La codificación se deja en paz.** No pongas `-Encoding UTF8`. PowerShell 5.1 lee ANSI por
   defecto, que es lo que escribe *CSV (delimitado por comas)*; y si usas *CSV UTF-8* el BOM
   lo detecta solo. Forzar UTF-8 rompe el caso ANSI y te deja `DISTRIBUIDORA RÃO` en el
   encabezado del Word. **[inferencia]**, pero el riesgo es real y el coste de evitarlo es cero.
5. **La fila que no cuadra tiene que quejarse en voz alta.** Hoy la validación
   `if ($num -notmatch '^\d{4,15}$') { continue }` (lote.ps1:126) **descarta en silencio**
   **[código]**. Con Excel vivo eso solo se comía la fila de encabezado. Con un CSV, una cuenta
   que Excel exportó en notación científica (`5,95E+08`) o sin el cero de la izquierda
   **desaparece del lote sin decir nada**, y no lo notas hasta que falta un cliente en la base.
   Cambia el `continue` por un `Log` con la fila y el valor crudo antes de saltar.

#### Qué se conserva

- **La detección de cuentas repetidas.** El hashtable `$vistas` de lote.ps1 (123, 131–135)
  pasa tal cual a `Get-ListaOrdenes`, con su mensaje *"repetida (ya salió en la fila N), se
  salta"*. **[código]**
- **`-Fila N` de `buscar_cuenta.ps1`**, ahora contra el CSV. Es lo que mantiene con vida el
  camino de un cliente suelto (`ARRANQUE.md`, sección *Cómo se corre*), que con el alcance
  nuevo importa **más** que antes, no menos.

#### Qué se pierde

- **La lista deja de estar viva.** Si vuelves a bajar el archivo de órdenes, tienes que
  volver a exportar el CSV. Hoy basta con tenerlo abierto.
- El motivo registrado en `docs/decisiones.md` para leer el libro abierto —*el nombre trae el
  número de descarga y cambia cada vez*— **sigue siendo cierto**, y por eso el CSV va a una
  ruta fija (`azul\lista_ordenes.csv`) y no al nombre que le ponga Excel. Que tú lo
  sobrescribas cada mañana es más simple que localizar la hoja por COM, pero es un paso
  manual más que antes no había.

#### Qué se gana

- Fuera el COM de Excel, fuera la dependencia de que el libro esté abierto mientras corre,
  fuera la localización por hoja `resultData`.
- **Fuera un lector duplicado.** Dos copias de la misma lógica de encabezados acaban
  divergiendo; es el mismo argumento que llevó `revisar_csv.ps1` a su propio archivo
  (`revisar_csv.ps1`, cabecera). **[código]**
- Tiempo: poco. El escaneo del libro corre **una vez por corrida**, no por cliente.
  **[inferencia]** 2–5 s por corrida entera.

#### Aviso que hay que añadir aquí, y no es opcional

`buscar_cuenta.ps1:1008` **[código]**:

```powershell
if ($razon -ne "") {   # ... comprobación contra la razón social ...
```

**La guarda más fuerte del sistema —la que impide entrar al cliente equivocado— es
condicional, y cuando la razón social viene vacía no se aplica y nadie lo dice.** Hoy pasa
desapercibido porque el Excel siempre la trae. Con un CSV exportado a mano, una columna mal
seleccionada la deja vacía y la corrida sigue igual de silenciosa.

Añade una línea en el `else`: `"AVISO: sin razon social. La comprobacion contra el cliente
equivocado NO se aplica en esta corrida."`. No cambia ningún comportamiento; solo deja de
callarse.

#### Reversible

Sí. Las dos funciones viejas se pueden pegar de vuelta desde git. Mientras dure la
transición puedes incluso dejar las dos: si `-Lista` apunta a un CSV que no existe, caer al
Excel. **No lo recomiendo** —dos caminos es justo lo que este plan viene a quitar— pero
existe si la primera mañana sale mal.

#### Orden

**Segundo.** Antes que la captura, porque se puede probar entero con `-Simular`, sin Azul y
sin Word.

---

### CAMBIO C — fuera la maquinaria del lote largo

#### Qué se quita, en `lote.ps1`

| Qué | Línea | |
|---|---|---|
| `param -Reanudar` | 35 | |
| `Get-YaHechas` | 158–167 | |
| bloque `if ($Reanudar)` | 213–217 | |
| comprobación `$yaHechas.ContainsKey` | 223, 246 | |
| línea final *"vuelve a correr con -Reanudar"* | 324 | |
| `param -Carpeta` | 26 | |
| `param -Desde` | 28 | |
| `param -PorBase` | 29 | |
| `Get-Word` | 170–177 | |
| `Get-RutaBase` | 181–188 | **el correlativo automático** |
| `New-Base` | 192–199 | |
| `Close-Base` | 201–207 | |
| corte de base a los 10 | 310–313 | |
| `Test-Path $Carpeta` | 235 | |

#### Qué se quita a medias, y con cuidado

- **`-Simular` (219–233) se simplifica, NO se borra.** Hoy imprime el plan con numeración de
  bases; sin bases eso sobra. Pero es **la única prueba del lote que no toca Azul ni Word**
  **[código]** — y con Excel fuera, es también la prueba del lector de CSV. Déjalo listando
  las cuentas en orden, con las repetidas marcadas, y el total.
- **`lote_progreso.csv` se queda.** Se va `-Reanudar` y `Get-YaHechas`; **`Add-Progreso`
  (146–156) no se toca.** Motivo en la sección 8, punto 4.

#### Qué entra en su lugar — y así se resuelve el nombrado

**El sistema deja de nombrar y de crear documentos.** Escribe en el que tú tengas activo.

Eso es lo que resuelve el desalineo, y lo resuelve de raíz: `Get-RutaBase` decide el número
mirando qué archivos existen —es su única memoria (`ESTADO.md`, nota final)— así que en cuanto
tú renombras uno, el número queda libre y la corrida siguiente lo reutiliza. **No hay forma de
arreglar eso mientras el sistema siga nombrando.** Si no nombra, no hay correlativo que
colisionar, y tú sigues bautizando al final con los 10 dentro, como haces hoy.

Al arrancar, `lote.ps1` hace estas cuatro comprobaciones —**antes del primer cliente**, no al
ir a escribir, para no descubrir el problema después de tres minutos de lectura:

1. Word abierto y con al menos un documento. Si no → error claro. **No abre Word ni crea nada.**
2. Si hay **más de un documento** abierto → exige `-Documento <ruta>` y lo activa. Es la
   traducción literal de `REGLAS.md` §6: *"si hubiera más de un documento de Word abierto →
   preguntar cuál"*.
3. El documento destino **tiene ruta en disco** y **no es una copia de autorrecuperación**.
   Es la misma pareja de guardas de `agregar_a_word.ps1` (108–115) **[código]**, adelantada.
4. **Imprime el nombre del documento destino**, y lo repite en cada línea de progreso.

Se conserva `$doc.Save()` después de cada cliente (línea 304) **[código]**: es lo que impide
que una caída a mitad se lleve lo escrito, que es el motivo por el que existía "guardar al
crear" (`docs/decisiones.md`).

#### Qué se gana

- El correlativo desaparece como concepto. **El problema no se mitiga: deja de existir.**
- `lote.ps1` pierde ~70 de sus 324 líneas **[código]** y todo el trato con COM de Word salvo
  activar y guardar.
- Un modo de fallo menos: hoy, si una corrida de prueba entra por error a la carpeta real, deja
  `BASE 034` con 2 clientes y la siguiente se va a `035` para no pisarla, **partiendo la base
  en dos** (`ESTADO.md`) **[registrado]**. Sin creación automática, eso no puede pasar.

#### Qué se pierde

- **El corte automático a los 10.** Ahora cuentas tú. Sigue siendo de 10; lo que cambia es
  quién lleva la cuenta. Apúntalo en `ARRANQUE.md` para que no se pierda la regla junto con
  el código que la aplicaba.
- **Reanudar.** Con bloques de 3–5 clientes y tú delante, relanzar es más barato que la
  máquina de reanudar. Y el registro de qué se hizo lo sigue dando `lote_progreso.csv`.
- **Un riesgo nuevo: equivocarte de documento activo.** Se mitiga con las cuatro
  comprobaciones de arriba, y de todos modos es exactamente el modo en que ya funciona
  `agregar_a_word.ps1` a mano —el camino más probado del proyecto.

#### Reversible

Sí, pero es el cambio menos reversible de los tres: el bloque de Word de `lote.ps1` son ~40
líneas contiguas y se recuperan de git, pero si a mitad de camino decides volver al
correlativo, vuelve también el desalineo. **Aplícalo el último y con el commit anterior limpio.**

#### Orden

**Cuarto.** El último de los cambios de fondo.

---

### CAMBIO D — limpieza de código que ya estaba muerto

Esto no lo pediste; lo encontré al buscar los huérfanos. Es independiente de A, B y C.

| Qué | Dónde | Estado |
|---|---|---|
| `[AzulShot]::Front()` | `captura.ps1` 72–87 | **cero llamadas hoy** **[código]** |
| `Test-CapturasDistintas` | `captura.ps1` 287–319 | **cero llamadas hoy** **[código]** |

**`Front()` sí conviene borrarlo**, y no por espacio: es *literalmente* la trampa 7. Llama a
`ShowWindow(SW_RESTORE)` siempre y des-maximiza la ventana entre medir y pulsar
(`docs/bitacora.md`, 03/09). Está ahí, compilado, con nombre casi idéntico a `Frente()`, y lo
único que impide que alguien lo llame por error es un comentario. Bórralo y deja el comentario
de `Frente()` explicando por qué no existe.

**`Test-CapturasDistintas` déjalo o bórralo, da igual.** Son 33 líneas inertes. Si lo borras,
comprueba que `probar_captura.ps1` no lo llame —**no lo llama**, copia el muestreo inline
(línea 21–22) **[código]**— y quita la mención del comentario de `zoom.ps1:43`.

---

## 2. Qué queda huérfano — las tres categorías

### 2.1 Se borra: no lo usa nadie

| Qué | Archivo | Comprobado |
|---|---|---|
| `Get-VistaCliente` | `azul_fast.ps1` 642–699 | única llamada en 708 y 726, las dos se van |
| `EsperarMarcoCuenta` | `azul_fast.ps1` 255–266 | única llamada en 662 **[código]** |
| `MarcoCuentaFull` | `azul_fast.ps1` 236–238 | solo la llama `MarcoCuentaAc`, que se queda pero puede usar solo la vía de panes |
| rama `-SoloCaptura` | `azul_fast.ps1` 706–713 | |
| `CTA_PREFIX` como prefijo de *espera* | `azul_fast.ps1` 218 | el `const` se queda: lo usa `MarcoCuentaPanes` |
| `Get-Ordenes` | `lote.ps1` 103–143 | |
| `Get-OrdenDeExcel` | `buscar_cuenta.ps1` 644–675 | |
| `Get-RutaBase`, `New-Base`, `Close-Base`, `Get-Word` | `lote.ps1` 170–207 | |
| `Get-YaHechas` | `lote.ps1` 158–167 | |
| `[AzulShot]::Front()` | `captura.ps1` 72–87 | **ya huérfano hoy** |
| `Test-CapturasDistintas` | `captura.ps1` 287–319 | **ya huérfano hoy** |

### 2.2 Se queda aunque no se use: es barato y puede hacer falta

| Qué | Dónde | Por qué |
|---|---|---|
| `Wait-PantallaEstable` | `captura.ps1` 335–381 | `probar_captura.ps1` §7 hace `Get-Command` sobre ella y lee su código fuente para comprobar la garantía de 1400 ms **[código]**. Borrarla obliga a editar el diagnóstico. Coste de dejarla: cero, nadie la llama |
| `[AzulShot]::Diferencia` + `Punto` | `captura.ps1` 181–214 | igual: `probar_captura.ps1` la prueba en cinco casos. Es lo que caza la trampa 14 |
| `Test-CapturaValida` | `captura.ps1` 386–413 | 28 líneas inertes. Si algún día vuelve cualquier captura automática, la necesitas |
| `diagnostico\excel_ordenes.ps1` | | Sin Excel en producción sigue siendo **la herramienta para depurar una exportación mala**: el CSV sale de ese libro. Y borrarlo obliga a editar `verificar.ps1` (ver 2.3) |
| `[Azul]::Init()` separado de `Run()` | `azul_fast.ps1` 288 | se separó para servir a `-SoloCaptura` (comentario 286–287). `-SoloCaptura` se va, la separación no estorba |
| `-SinCaptura` | | opcional: si lo dejas, queda como un `switch` que no hace nada. Yo lo quitaría por higiene, pero no rompe nada dejarlo |

### 2.3 Parece huérfano pero algo depende de él — **léelo entero**

**1. `captura.ps1` NO es "la captura".** Es lo único que toca la máquina, y la búsqueda —que
es intocable— depende de él en ocho puntos **[código]**:

| Qué usa `buscar_cuenta.ps1` | Línea |
|---|---|
| `Get-AzulHwnd` (+ guarda de minimizado) | 702 |
| `[AzulShot]::IsIconic` | 696 |
| `[AzulShot]::Frente` | 698, 710, 724, 775, 811 |
| `[AzulShot]::Tecla` (FIN / RETROCESO) | 784, 785, 817, 818 |
| `[AzulShot]::Escribir` (el número de cuenta) | 819 |
| `Invoke-AzulClick` (icono, Buscar Ahora, celda, Seleccionar, pestaña) | 733, 754 |
| `Get-AzulRect` | 729 |
| `Save-AzulShot` (vía `Save-Prueba`) | 713 |

Y `azul_fast.ps1` lo sigue necesitando para `Get-AzulHwnd` (703). Y `zoom.ps1` lo carga (37)
para medir coordenadas al píxel — que es el procedimiento que la propia `REGLAS.md` §5 manda
usar cuando haya que volver a medir. **Borrar `captura.ps1` mata la búsqueda intocable en la
primera línea.**

**2. Cerrar el panel del cliente.** Ya explicado en el cambio A. Lo repito porque es el que
se cuela: el código que abre la foto y el que la cierra están en la misma función, y separarlos
es el trabajo real de este cambio.

**3. `Save-AzulShot` es el mecanismo de evidencia de la búsqueda.** `Save-Prueba`
(`buscar_cuenta.ps1` 708–715) produce las capturas `marco_preexistente`,
`seleccionar_sin_efecto`, `sinresultado`, `pestana_sin_efecto` y `llenado` **[código]**. Cuando
una guarda dispara, esa imagen es lo único que dice qué había en pantalla — es lo que permitió
diagnosticar el `Seleccionar` sin efecto del 04/09 **[registrado]**.

**4. `verificar.ps1` se rompe si borras archivos sin editarlo.** Sus dos listas (13–18) incluyen
`captura.ps1`, `excel_ordenes.ps1` y `probar_captura.ps1`. Un archivo que falta cuenta como
**`NO EXISTE` → `$fallos++` → nunca dice `TODO OK`** (24, 72) **[código]**. Y tu regla es que
sin `TODO OK` no se corre nada. **Borrar un archivo y editar `verificar.ps1` van en el mismo
commit, siempre.** Lo mismo si añades `lista.ps1`: hay que meterlo en `$archivos` o su ASCII y
su sintaxis no se comprueban nunca.

**5. `probar_captura.ps1` depende del *código fuente* de `captura.ps1`, no solo de sus
funciones.** Lee el archivo con `Get-Content -Raw` y busca los valores por regex
(`\[int\]\$Intervalo\s*=\s*(\d+)`, líneas 141–143) **[código]**. Si renombras esos parámetros
o los mueves, el diagnóstico se cae al valor por defecto y **deja de comprobar lo que cree que
comprueba**, sin fallar.

**6. La razón social sostiene la guarda del cliente equivocado.** Ya explicado en el cambio B.
Si el CSV la pierde, la guarda no falla: **se desactiva en silencio**.

**7. Un acento roto en la razón social falla hacia el lado seguro — pero se ve en el Word.**
Si la codificación del CSV se tuerce, la guarda de la línea 1011 compara `MÃA` contra lo que
Azul devuelve y **aborta** (no entra al cliente equivocado, que es lo correcto). Pero la misma
cadena va al encabezado del documento (`agregar_a_word.ps1` 154), así que el síntoma que verás
primero es un nombre mal escrito en la base, no un error. **[código] + [inferencia]**

**8. `revisar_csv.ps1` tiene dos llamadores.** `lote.ps1:282` y `azul_fast.ps1:760`
**[código]**. Si tocas su valor de retorno, se rompen los dos. No hay motivo para tocarlo en
este plan; lo apunto porque no se ve desde ninguno de los dos.

---

## 3. Esperas por reloj contra esperas por condición

### 3.1 Qué esperas fijas quedan después de los cambios

**Todas las que quedan están en el camino del clic real**, que es justo el que decidiste
conservar (trampa 7). **[código]**

| Espera | Dónde | Decisión |
|---|---|---|
| `Start-Sleep 800` tras `-Restaurar` | `buscar_cuenta.ps1` 699 | se queda (tu lista) |
| `Start-Sleep 400` en `Save-Prueba` | `buscar_cuenta.ps1` 711 | se queda (tu lista) |
| `Start-Sleep 250` en `Invoke-ClicGeo` | `buscar_cuenta.ps1` 725 | se queda (tu lista) |
| `Start-Sleep 700` tras el clic en la celda | `buscar_cuenta.ps1` 1001 | se queda (tu lista) |
| `Thread.Sleep(350)` × ≤6 en `Frente()` | `captura.ps1` 140 | **no estaba en tu lista, pero es la misma familia.** Se queda |
| `Thread.Sleep(120)` + `(60)` en `Click()` | `captura.ps1` 146, 148 | misma familia. Se queda |
| `Thread.Sleep(40)` por carácter en `Escribir()` | `captura.ps1` 108 | medido el 03/09: con 25 el campo se quedaba a medias. **No tocar** |
| `Thread.Sleep(25)` en `Tecla()` | `captura.ps1` 117 | misma familia. Se queda |
| `Pump(80)` tras `clearAccessibleSelection` | `azul_fast.ps1` 392 | asentamiento de 80 ms. Se queda |
| `Pump(120)` idem en la rejilla | `buscar_cuenta.ps1` 506 | se queda |
| `Pump(1200)` entre reintentos de línea | `azul_fast.ps1` 559 | deliberada. Se queda |
| `Pump(250)` × ≤60 del árbol de atributos | `azul_fast.ps1` 467 | **prohibido tocar** (`docs/decisiones.md`, `docs/trampas.md`) |

**Y desaparece una que no era del clic**: el suelo de `Wait-PantallaEstable`, 250 ms × 6 = **1,5 s
obligatorios** aunque la pantalla ya estuviera quieta (`captura.ps1` 348, 351) **[código]**. Es
la única espera con suelo fijo que este plan elimina.

**Queda una fija que sí conviene convertir** — la única de la lista que no es del clic real:

> **`CommitCF`, `buscar_cuenta.ps1` línea 364** **[código]**
> ```csharp
> if(otro!=0){ requestFocus(vm,otro); Pump(400); }
> string tras=Texto(c);
> if(tras==null || tras.Trim()!=cuenta) return "ERROR: el campo perdio el valor...";
> ```
> Espera **400 ms fijos** a que Siebel recoja el valor al perder el foco, comprueba **una vez**,
> y si no cuadra **aborta la corrida**. Con Azul cargado, 400 ms se quedan cortos y el cliente
> se pierde por un falso negativo — el mismo patrón que ya costó corridas con el tecleo
> (trampa 8), donde la solución fue releer en bucle.
> **Cámbialo por un sondeo:** hasta ~3 s, comprobando cada 200 ms, y salir en cuanto el valor
> cuadre. Con Azul sano no cuesta nada (sale a la primera vuelta); con Azul cargado convierte
> un aborto en una corrida buena. **Es exactamente el caso que describiste: 15 s en vez de 10,
> pero por condición.**
> **[código]** el hecho; **[inferencia]** que sea la causa de algún fallo — no lo he visto ocurrir.

### 3.2 Topes de condición que conviene subir

Los tres son sondeos: **suben el techo del peor caso y no cuestan nada en el caso bueno**,
porque salen en cuanto la condición se cumple.

**1. `Init()`: 800 ms → 5000 ms.** `azul_fast.ps1:294` y `buscar_cuenta.ps1:236` **[código]**.

Es el tope para que el puente devuelva el contexto raíz **y** para cachear los desktop panes.
Dos motivos para subirlo:

- Cuando vence, la corrida muere con `ERROR: sin contexto raiz. El puente no ve a Azul`. Es
  **literalmente el mensaje con el que murió el cliente 2 de la corrida de las 13:24**, con
  Azul vivo y quemando un núcleo al 100% **[registrado]**. Es el modo de fallo del *Azul
  degradado*, que es el que más te ha costado.
- El bucle de caché de panes **comparte el mismo tope** (líneas 304–307 / 245–248). Si vence,
  `panes` queda vacío, y entonces **cada sondeo posterior cae al recorrido completo del árbol**
  — que es justo lo que el caché existe para evitar. O sea que un `Init` apurado no solo
  arriesga el aborto: deja la corrida entera lenta, en silencio. **[código]**

**2. `topeSel`: 950 ms → ~2500 ms.** `azul_fast.ps1:406` **[código]**.

Tope para que la fila quede `checked` y `Ver Productos Asignados` se habilite. Si vence,
`marcada=false` → `"no se marco la fila esperada"` → se marca transitorio y se **reintenta la
línea entera**. Con Azul cargado, subir el techo cuesta menos que el reintento. El comentario
del código dice que lo normal es salir a los ~160 ms **[código]**: en el caso bueno no cambia
nada.

**3. Los que ya están bien, para que no los toques.** `EsperarCIM(40)`, `EsperarResultado(90)`,
`EsperarPestana(40)`, el tope de 40 s por línea de `azul_fast`, y la estabilización de la tabla
de Suscripciones (60 × 400 ms, 3 lecturas iguales). Todos son sondeos con techo generoso.
**Ninguno necesita subir.**

Y una regla que el propio código documenta y conviene no romper
(`buscar_cuenta.ps1` 889–907) **[código]**: la doble llamada a `EsperarResultado` —15 s y luego
90 s— **no es redundante**. La de 15 s es una sonda para decidir si repetir el clic; la de 90 s
es la espera de verdad. Ya se intentó condicionar la segunda el 04/09 y **rompió la búsqueda**.
No lo vuelvas a intentar.

---

## 4. Documentación — qué encoge

**`REGLAS.md`: no se toca. Y no hace falta tocarlo.**

Merece explicación, porque parece que sí haría falta:

- **§3, ciclo 2 (vista general del cliente): sigue vigente.** No desaparece: **pasa a hacerlo
  tú a mano**, para tu captura. Las mismas reglas se aplican igual — clic en `Cliente:`,
  esperar, cerrar con la X, y no tocar el botón `Crear` que hay dentro. Que ahora lo ejecute
  una persona en vez de un script no cambia una coma del texto.
- **§5, la coordenada (500,120): pasa de ser un valor del código a una nota para un humano.**
  El párrafo sigue siendo verdad y sigue siendo útil (te dice dónde está el enlace y por qué
  con nombres cortos hay que apuntar a la izquierda). Lo único que deja de valer es la frase
  *"parámetros `-ClienteX`/`-ClienteY` de `azul_fast.ps1`"*, porque esos parámetros ya no
  existirán.

  **Eso es una contradicción entre `REGLAS.md` y el código, y es tuya de resolver, no mía.**
  Yo no lo toco. Lo dejo anotado aquí y en `docs/decisiones.md` para que quien lea el archivo
  después sepa que el desajuste es conocido y deliberado.

| Archivo | Qué encoge |
|---|---|
| `ARRANQUE.md` | *Cómo se corre*: parámetros nuevos de `lote.ps1`. *Condiciones*: fuera "Excel con la lista abierta", entra "CSV exportado" y "el documento de la base abierto y activo". **Y añade la regla de los 10 clientes**, que se queda sin código que la aplique |
| `ESTADO.md` | Mucho. Fuera: el objetivo de los 30, *Lo que sigue* 1/4/5, el corte de base sin ejercitar, las menciones a `-Reanudar`, y la nota final del nombrado (deja de aplicar). **Se queda entero**: el bloqueante de `Interaccion()`, el `Seleccionar` sin efecto, y las dos corridas del 04/09 — describen a Azul, no al alcance |
| `docs/mapa-datos.md` | §1 (Excel) → reescribir como el contrato del CSV. §5 (marco del cliente) → pasa a describir **lo que haces tú a mano**, más la guarda que lo cierra. §6 → nota de que `# IMAGEN:` ya no se escribe |
| `docs/decisiones.md` | **No borres nada. Añade una sección "Decisiones superadas"** y mueve ahí tres: *por qué se lee el Excel abierto*, *por qué 10 por base y por qué se guarda al crear*, *por qué la foto va al final*. Borrarlas invita a que alguien reproponga lo que ya se descartó — que es justo lo que ese archivo existe para impedir. **Añade una entrada nueva**: por qué el sistema dejó de nombrar las bases |
| `docs/trampas.md` | Casi nada. La 5 (*ninguna espera fija sirve*) usa la foto como ejemplo, pero **la regla sobrevive al ejemplo**. La 13 (autorrecuperación) pasa a ser **más** importante, no menos: ahora el lote escribe en `ActiveDocument` en vez de crear el documento él. La 14 se queda, como pediste |
| `docs/bitacora.md` | **Cero.** Es historia de incidentes. No encoge nunca. Añade una entrada solo si algo se rompe al aplicar esto |
| `azul/LEEME.md` | Ya está desactualizado hoy (`ESTADO.md` lista cinco puntos). Con estos cambios queda peor: su primer párrafo describe la foto del cliente. **Decide: actualizarlo o marcarlo superado como `PROYECTO-AZUL.md`.** Yo lo marcaría superado y movería la tabla de mensajes de error a `docs/trampas.md`, que es donde la mandan a buscar |

---

## 5. Tiempo estimado — y por qué no llega a 1,5 min

### 5.1 De dónde salen los números

**[registrado]**, corrida buena del 04/09, 13:07–13:13:

| | |
|---|---|
| 2 clientes, **6,2 min** = 372 s | **186 s = 3,1 min por cliente** |
| CLIENTE EJEMPLO UNO, 17 líneas (16 activas) | *"leída en 117,2 s"* |
| ACME, 20 líneas (8 activas, 12 canceladas) | *"leída en 87,2 s"* |
| Ritmo medido el 02/09 | **4,6 s/línea** (18 líneas en 82,6 s) |

Suma de las dos lecturas: 204,4 s de 372. **Quedan 167,6 s sin asignar = 83,8 s por cliente**
para búsqueda, foto, `revisar_csv`, Word y el arranque de dos procesos de PowerShell.

**Aquí está el hueco que no puedo cerrar desde aquí:** `ESTADO.md` dice *"leída en 117,2 s"*
sin decir si ese número incluye la foto o solo el bucle de líneas. Si la incluye, la foto está
dentro de los 204 s; si no, está dentro de los 84. **[no se puede saber sin Azul]**

**Cómo lo cierras en la oficina, sin cronómetro:** `progreso.txt` lleva `HH:mm:ss` en cada
línea (`azul_fast.ps1:99` **[código]**) y escribe estas marcas:

```
captura: clic en el enlace Cliente en (500,120)
captura: panel abierto - Cuenta: ACME SA DE CV
captura: espera de pantalla estable = N s
captura: guardada en ...
captura: cierre del panel del cliente = True
```

**La resta entre la primera y la última es el coste real de la foto.** Míralo en un
`progreso.txt` que ya tengas de una corrida vieja: no hace falta correr nada nuevo.

### 5.2 La estimación

Coste de la foto, sumando lo que el código garantiza **[código]** más lo que hay que suponer
**[inferencia]**:

| Paso | |
|---|---|
| `Frente()` | ≥350 ms |
| `Click()` | 180 ms |
| `EsperarMarcoCuenta` hasta el título `Cuenta: ` | **3–10 s [inferencia]** — es Azul cargando la vista entera del cliente |
| `Wait-PantallaEstable` | **≥1,5 s** garantizados + PrintWindow por vuelta |
| `Save-AzulShot` + `Test-CapturaValida` | ~0,5–1 s |
| `Close-MarcoCuenta` + `EsperarSinMarcoCuenta` | **1–3 s [inferencia]** |

**Total: 10–25 s por cliente. [inferencia]**

Excel: ~2–5 s **por corrida entera**, no por cliente **[inferencia]** — `Get-Ordenes` corre
una vez (lote.ps1:210).

**Resultado: 186 s − ~18 s ≈ 168 s ≈ 2,8 min por cliente.**

| Cliente | Hoy | Después | |
|---|---|---|---|
| 16 líneas activas | ~3,1 min | **~2,7–2,9 min** | **[inferencia]** |
| 8 líneas activas | ~2,4 min | **~2,0–2,2 min** | **[inferencia]** |

**No llega a 1,5 min, y no por poco.** Una base de 10 clientes pasaría de ~35 min a ~28.

### 5.3 Por qué, y qué habría que tocar

**La lectura línea a línea es el 55–65% del tiempo y este plan no la toca.** A 4,6 s/línea, un
cliente de 16 líneas son 74 s **solo de lectura** — 1,2 min. **1,5 min por cliente está por
debajo del suelo de la lectura para un cliente típico**, así que el objetivo no es alcanzable
sin entrar ahí.

Los tres sitios donde queda tiempo, **de mayor a menor**, y qué cuesta cada uno:

**1. Abrir el detalle de cada línea. — Toca `REGLAS.md` §4. No lo hagas.**
Cada línea abre `Ver Productos Asignados`, espera a que el árbol de atributos deje de crecer,
lee y cierra. Esa espera es fija a propósito y `docs/decisiones.md` la marca como intocable
(*"ahí no hay condición fiable que preguntar"*). La única forma de saltársela es el botón
**`Exportar`**, que está en la lista de los seis que nunca se pulsan. **Lo digo porque pediste
que lo dijera en vez de callarlo: la vía rápida existe y está vetada.**

**2. Dos procesos de PowerShell por cliente, cada uno compilando C# al arrancar.**
`lote.ps1` lanza `buscar_cuenta.ps1` y luego `azul_fast.ps1`, cada uno con su `Add-Type` **y**
el de `captura.ps1` (lote.ps1:251, 267 **[código]**). Fundirlos en un solo proceso por cliente
ahorraría un arranque y una compilación completos: **3–8 s por cliente [inferencia]**.
No toca la *lógica* de ninguna guarda, pero mueve el archivo donde viven todas.
**Recomendación: no ahora.** Mídelo primero.

**3. `Firma()` dentro de los bucles de espera de la búsqueda.**
`EsperarResultado` llama a `Firma()` en bucle (459–475 **[código]**), y cada `Firma()` hace
**cuatro operaciones sobre el árbol entero**: `Marco(F_CIM)` —que cuando la búsqueda acierta
**no encuentra el formulario y por eso recorre todo el árbol antes de poder decir que no
está**—, `MarcoResultados()` (que itera todos los marcos internos con un `First` dentro de cada
uno), el conteo de marcos, y `Interaccion()`. Es exactamente lo que `docs/trampas.md` prohíbe:
*"nunca recorrer el árbol dentro de un bucle de espera"*.
Arreglarlo probablemente sea la mayor ganancia por línea de código de todo el proyecto — y
**está en el corazón de la parte intocable**. **Recomendación: no ahora, y nunca sin medir
antes.**

### 5.4 La medida que ya tienes y no estás usando

Los hallazgos 8, 9, 15 y 16 llevan aplicados y sin medir. **No hace falta cronómetro**: el
contador de llamadas al puente ya está instrumentado y ya se imprime **[código]**.

- `buscar_cuenta.ps1` imprime `Llamadas JAB de arbol: N` al terminar (1095) y escribe dos
  marcas en `progreso.txt`: `JAB tras la busqueda` (922) y `JAB total` (1096).
- `azul_fast.ps1` mete `N llamadas JAB` en el pie del CSV (612) y en `progreso.txt` (613).

El propio comentario del código dice por qué sirve: *"el reloj depende de lo cargado que esté
Azul, el número de llamadas no"* **[código]**. **Dos corridas de la misma cuenta, antes y
después de un cambio, comparando ese número, te dicen si el cambio sirvió** aunque Azul esté
teniendo un mal día. Es la única medida de este proyecto inmune a la degradación que te tumbó
el objetivo original.

---

## 6. Dependencias ocultas — el resumen

Está todo en la sección **2.3**, que es la que hay que leer entera antes de borrar nada. Los
tres que de verdad muerden:

1. **`captura.ps1` sostiene la búsqueda intocable.** Ocho puntos de uso.
2. **El cierre del panel del cliente protege contra algo que tu flujo nuevo va a provocar.**
3. **Borrar un archivo sin editar `verificar.ps1` deja el proyecto sin `TODO OK`**, y por tu
   propia regla eso significa que no se puede correr nada.

Y dos que no muerden pero se callan:

4. La guarda de la razón social **se desactiva sola** si el campo viene vacío, sin decirlo.
5. `probar_captura.ps1` valida `captura.ps1` **leyendo su código fuente por regex**; si
   renombras esos parámetros, deja de comprobar y no falla.

---

## 7. (a) Orden exacto de aplicación

Pensado para que **si paras a la mitad, lo que quede siga funcionando entero**. Cada etapa es
un commit. Ninguna depende de la siguiente.

### Etapa 0 — antes de tocar nada (sin Azul, 10 min)

1. `git commit` de todo el estado actual. Es la única marcha atrás real.
2. `verificar.ps1` → tiene que decir **`TODO OK`**.
3. `probar_captura.ps1` → **`TODO OK`**.
4. Crea la carpeta temporal de pruebas y un `.docx` vacío dentro, **guardado en disco**.
5. Exporta el CSV desde Excel y **ábrelo con el Bloc de notas**: mira con tus ojos el
   separador (`,` o `;`), los acentos, y que las cuentas se vean con sus 9 dígitos y **no
   como `5,95E+08`**. Este minuto te ahorra la mitad de los problemas de la etapa 2.

### Etapa 1 — el hueco en Word (sin Azul)

Solo `agregar_a_word.ps1`: añadir `Add-Parrafo $d ""` en la rama sin imagen.

**Es aditivo: si el CSV todavía trae `# IMAGEN:`, la foto se sigue insertando igual.**
→ *Si paras aquí: el sistema funciona exactamente como hoy.*

### Etapa 2 — Excel → CSV

1. Crear `azul/lista.ps1` con las dos funciones.
2. **Añadir `lista.ps1` a `$archivos` en `verificar.ps1`.** En el mismo commit.
3. `lote.ps1`: fuera `Get-Ordenes`, dentro `Get-ListaOrdenes` + `param -Lista`.
4. `buscar_cuenta.ps1`: fuera `Get-OrdenDeExcel`, dentro `Get-OrdenDeLista`. `-Fila` sigue
   funcionando.
5. Añadir el aviso de razón social vacía (`buscar_cuenta.ps1` ~1008).

→ *Si paras aquí: el sistema funciona entero, leyendo del CSV, y con la foto puesta.*

### Etapa 3 — fuera la captura del cliente

Solo `azul_fast.ps1`. **En este orden dentro de la etapa**, que importa:

1. **Primero** mover el cierre del panel: añadir la comprobación de marco `Cuenta: ` abierto
   al principio de `Run()`, junto al `CerrarDetalle()` de la línea 318.
2. **Después** borrar `Get-VistaCliente`, `EsperarMarcoCuenta`, `MarcoCuentaFull`, la rama
   `-SoloCaptura`, los parámetros y el bloque de cabeceras.
3. (Opcional, recomendado) `-Cuenta`/`-Razon` en `azul_fast.ps1` para conservar `# CLIENTE:`.

Al revés —borrar primero y añadir después— dejas una ventana en la que el panel no lo cierra
nadie.

→ *Si paras aquí: el sistema funciona entero, sin foto, con bases automáticas.*

### Etapa 4 — fuera la maquinaria del lote largo

> ⚠ **REVERTIDA el 09/09/2026.** Las bases automáticas, el correlativo y el corte a los 10
> volvieron, y con ellos se fueron las cuatro comprobaciones del documento activo. Lo que **no**
> volvió es el fallo que motivó esta etapa: el correlativo ya no se deduce de la carpeta, sale
> de una memoria propia que solo avanza. El diagnóstico de abajo sigue siendo correcto y por eso
> se conserva entero — la cura es la que cambió. Ver `docs/decisiones.md`, *"Por qué las bases
> volvieron a ser automáticas"*.

Solo `lote.ps1`. Fuera reanudar, bases, correlativo y corte. Dentro las cuatro comprobaciones
del documento activo. `-Simular` simplificado, `Add-Progreso` intacto.

→ *Si paras aquí: ya está el alcance nuevo completo.*

### Etapa 5 — los topes de condición

**Uno por commit, con `verificar.ps1` entre medias:**

1. `Init()` 800 → 5000 ms (**en los dos archivos**).
2. `topeSel` 950 → 2500 ms.
3. `CommitCF`: `Pump(400)` fijo → sondeo hasta 3 s.

### Etapa 6 — limpieza y documentación

`[AzulShot]::Front()` fuera. Docs según la sección 4. `ESTADO.md` reescrito con lo que de
verdad quedó probado.

---

## 8. (b) Qué probar, en qué orden — de lo más barato a lo más caro

> **Nunca contra `BASES DE DATOS PS1`.** Siempre carpeta y documento temporales, hasta que la
> prueba 8 salga limpia **dos veces**.

| # | Prueba | Toca Azul | Cuándo |
|---|---|---|---|
| 1 | `verificar.ps1` → `TODO OK` | no | **después de cada edición, sin excepción** |
| 2 | `probar_captura.ps1` → `TODO OK` | no | siempre que toques `captura.ps1` |
| 3 | `lote.ps1 -Simular` con el CSV | no | etapa 2. Comprueba: nº de clientes, orden, repetidas marcadas, y **ninguna fila descartada en silencio** |
| 4 | `lote.ps1 -Simular` con el CSV **con una cuenta rota a propósito** (una letra, un vacío) | no | etapa 2. **Tiene que quejarse en voz alta.** Si la salta callando, el aviso no quedó puesto |
| 5 | `agregar_a_word.ps1 -Csv <un CSV viejo> ` contra el .docx temporal | no | etapa 1. Mira: cuenta y razón en negrita, **una sola línea en blanco**, `RENOVABLES - n`, dos tablas **separadas** (trampa 12), canceladas en su tabla de una columna |
| 6 | `buscar_cuenta.ps1 -Cuenta ... -Razon ... -Restaurar` suelto | **sí** | etapa 2. Que la búsqueda siga entera tras quitar Excel |
| 7 | `azul_fast.ps1 -SinWord` suelto, con el cliente ya cargado | **sí** | etapa 3. Mira el CSV (cuadre, bloques) y que **Azul quede en Suscripciones** |
| 8 | `lote.ps1 -Limite 1` contra el documento temporal | **sí** | etapas 3 y 4 |
| 9 | `lote.ps1 -Limite 3` contra el documento temporal | **sí** | final. **Cronometra**, y después lee `progreso.txt` |

### Casos que hay que forzar al menos una vez

Ninguno se ha ejercitado nunca o casi nunca (`ESTADO.md`) **[registrado]**:

- **Un cliente con varias coincidencias** (rejilla): que la comprobación contra la razón social
  del CSV siga imprimiendo *"La fila coincide..."*. Es la guarda que más te importa y ahora el
  dato viene de otra fuente.
- **Un cliente sin renovables**: no debe entrar al documento ni contar.
- **Un cliente con bloque `A REVISAR`**: sí debe entrar.
- **El panel del cliente dejado abierto a mano.** Ábrelo tú, saca tu captura, **déjalo abierto
  a propósito** y lanza el cliente siguiente. Tiene que cerrarlo o abortar con un mensaje
  claro. **Nunca leer la tabla equivocada.** Es la prueba del punto crítico del cambio A y no
  se puede hacer sin Azul.
- **`verificar.ps1` con un archivo renombrado a mano**, para ver con tus ojos que dice
  `NO EXISTE` y no dice `TODO OK`. Cuesta 20 segundos y te ahorra el fallo del punto 2.3.4.

---

## 9. (c) Lo que NO recomiendo quitar, aunque lo hayas pedido

**Sección obligatoria.** Ocho cosas.

**1. `captura.ps1`.** Pediste quitar la captura del cliente. Ese archivo no es la captura: es
el teclado real, el clic real y la guarda de ventana minimizada, y **la búsqueda intocable
depende de él en ocho puntos** (2.3.1). Se quita el *ciclo de la foto*, que vive en
`azul_fast.ps1`. El archivo se queda entero.

**2. El cierre del panel del cliente** (`MarcoCuenta`, `GeoMarcoCuenta`,
`EsperarSinMarcoCuenta`, `Close-MarcoCuenta`). Pediste quitar *"el ciclo completo de abrir el
panel / fotografiar / cerrarlo"*. **Quita las dos primeras partes, deja la tercera.** Con la
foto en tus manos, eres tú quien va a dejar ese panel abierto, y un panel abierto tapa la
tabla de Suscripciones del cliente siguiente. Son ~40 líneas y todo lo que necesitan
(`Invoke-AzulClick`, `Get-AzulRect`) se queda de todos modos.

**3. `Save-AzulShot` y `Save-Prueba`.** Caen dentro de *"fotografiar"*, pero no son la foto del
documento: son las capturas de las cinco guardas de la búsqueda. Cuando una dispara, esa
imagen es lo único que dice qué había en pantalla.

**4. `lote_progreso.csv` (`Add-Progreso`).** Pediste quitar *"reanudar con lote_progreso.csv"*.
Quita **`-Reanudar` y `Get-YaHechas`**; deja el registro. Con el nombrado en tus manos, ese
archivo pasa a ser **lo único que dice qué cuenta acabó en qué documento**, y es también donde
quedan los `SIN RENOVABLES` y los `FALLO TIEMPO` que no vas a recordar al día siguiente. Cuesta
una función de 10 líneas.

**5. `-Simular`.** Es maquinaria del lote, sí. También es **la única prueba del lote que no
toca Azul ni Word**, y tras el cambio B es la prueba del lector de CSV. Simplifícalo —fuera la
numeración de bases— pero no lo borres.

**6. `Wait-PantallaEstable` y `[AzulShot]::Diferencia`.** Quedan sin llamadas, pero
`probar_captura.ps1` las prueba en seis comprobaciones, una de ellas por `Get-Command`. Borrarlas
obliga a editar ese diagnóstico **para ganar exactamente cero segundos de corrida**, porque
nadie las llama. Déjalas quietas.

**7. `diagnostico/excel_ordenes.ps1`.** Parece el huérfano más obvio del cambio B. Pero el CSV
sale de ese libro de Excel, así que cuando la exportación salga rara —notación científica,
columna que no es, separador raro— **es la herramienta que te dice qué hay de verdad en la
hoja**. Y borrarlo obliga a editar `verificar.ps1` en el mismo commit o el proyecto se queda
sin `TODO OK`.

**8. La regla de los 10 clientes por base — como regla, no como código.** El corte automático
se va; la regla no. Si desaparece de `lote.ps1` **y** no aparece en `ARRANQUE.md`, en dos
semanas nadie recuerda por qué son 10 y aparecerá una base de 14. Muévela de código a
documento en el mismo commit.

---

## 10. Lo que este plan no puede cerrar sin Azul

Para que quien lo ejecute sepa qué queda abierto a propósito:

1. **El coste real de la foto.** Se cierra leyendo un `progreso.txt` viejo (sección 5.1).
2. **Si el párrafo vacío de Word sale como un solo salto de línea.** Se cierra con la prueba 5,
   sin Azul.
3. **El separador y la codificación de tu CSV.** Se cierra abriéndolo con el Bloc de notas
   (etapa 0, punto 5).
4. **Qué ganaron los hallazgos 8, 9, 15 y 16.** Se cierra comparando el contador JAB de dos
   corridas de la misma cuenta (sección 5.4).
5. **Si el aborto de `CommitCF` a los 400 ms llega a ocurrir de verdad.** No está en la
   bitácora. Lo propongo por analogía con la trampa 8, no porque lo haya visto.
6. **El bloqueante de `Interaccion()`** (`ESTADO.md`): por qué el marco deja de exponerse como
   `internal frame`. **Este plan no lo toca ni lo resuelve.** Sigue abierto y sigue siendo
   independiente del alcance.
