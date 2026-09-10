# Trampas — ninguna es opcional

> Ábrelo **solo cuando algo falle**. Si todo va bien, no hace falta.
> Cada una de estas costó una corrida perdida o peor. Ninguna se deduce leyendo el código.

**Antes de correr nada, siempre:**

```powershell
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\verificar.ps1"
```

Comprueba ASCII, parseo y **compila el C# por separado** —el de `azul_fast.ps1`,
`buscar_cuenta.ps1` **y `captura.ps1`**—, sin tocar Azul.
Debe decir `TODO OK`. Ha atrapado errores reales antes de llegar a la aplicación.

> El C# de `captura.ps1` estuvo **sin comprobar hasta el 04/09/2026**: su `Add-Type` usa la
> forma con `-ReferencedAssemblies` y aquí se buscaba el literal `Add-Type -TypeDefinition @'`.
> Ahora el marcador es solo `-TypeDefinition @'`, que sirve para las dos formas.

---

## Las 16

**1. ASCII puro obligatorio en los `.ps1`.**
PowerShell 5.1 lee los `.ps1` sin BOM como ANSI y **rompe los acentos en silencio**:
`"Plan Móvil"` se convierte en `"Plan MÃ³vil"`, la comparación deja de coincidir y
manda todas las líneas a REVISAR sin decir por qué. Los literales con acento van como
escapes: `"Plan M\u00F3vil"`, `"m\u00FAltiples "`. Los `.md` sí llevan acentos.

Verificación rápida — debe dar `0`:

```powershell
powershell.exe -Command "$b=[IO.File]::ReadAllBytes((Resolve-Path '.\azul\azul_fast.ps1')); ($b | Where-Object {$_ -gt 127}).Count"
```

**2. PowerShell de 32 bits** para todo lo que toque el puente:
`C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`. El JRE de Azul es de 32
bits y la DLL del puente también; desde el de 64 bits no carga. El COM hacia Word y
Excel de 64 bits **sí funciona** desde el de 32 (comprobado). `lote.ps1` corre en el
de 64 y lanza a los hijos en el de 32.

**3. El C# va en here-strings.** Un error de sintaxis **no aparece hasta la
ejecución**. Por eso `verificar.ps1` lo compila aparte.

**4. Bash está bloqueado en esta máquina** por Control de aplicaciones. Todo por
PowerShell.

**5. Ninguna espera fija sirve.** Siempre se sondea una condición: que el marco
aparezca, que el botón se habilite, que la pestaña quede `selected`, que el árbol deje
de crecer. Una espera de 3 s llegó a fotografiar un **formulario en blanco con botón
`Crear`**: parecía una imagen válida y no lo era.

**6. Azul numera las instancias de sus marcos.** La segunda vez el formulario se llama
`Encontrar Comunicante Búsqueda CIM [2]`. Hay que buscar por **prefijo**, no por
igualdad.

**7. `Front()` des-maximizaba la ventana.** Llamaba a `ShowWindow(SW_RESTORE)`
siempre, y sobre una ventana maximizada eso la mueve y la redimensiona — justo entre
medir un botón y pulsarlo, así que el clic caía fuera. Ahora existe **`Frente()`**, que
solo restaura si la ventana está **de verdad minimizada**. **Usar siempre `Frente()`.**

**8. El tecleo se trunca.** De `507000002` llegaron a entrar solo `507`. Un número
truncado **no falla: devuelve otro cliente.** Por eso se relee el campo, se espera
activamente y se reintenta hasta 3 veces; si no cuadra, **no se busca**.

**9. Criterios viejos en el formulario.** Se combinan con la cuenta y devuelven otro
cliente. Se limpian y se reporta qué se borró. Ojo: **un campo con solo espacios
cuenta** — mirar `v.Length`, no `v.Trim().Length`.

**10. Con una rejilla de resultados abierta, Azul ignora el clic en el ícono del
teléfono.**

**11. Azul reutiliza el marco de interacción.** El nombre puede ser idéntico para dos
clientes distintos. **"Existe el marco" es la condición**; que el nombre no cambie se
avisa aparte.

**12. `Range.Text` sobre un párrafo de Word borra la marca de fin de párrafo** y funde
las tablas en una sola. Se usa `Add-Parrafo`, que hace `MoveEnd` para dejar la marca
fuera del rango.

**13. No escribir a ciegas en `ActiveDocument`.** Word puede tener abierta una copia de
autorrecuperación; ya pasó. `agregar_a_word.ps1` se niega si el documento activo es
autorrecuperado o si nunca se guardó en disco (`-Forzar` lo salta).

> Desde el 09/09/2026 el lote **no usa `ActiveDocument`**: le pasa al escritor el documento
> por objeto (`-DocObj`), así que ya no hay nada que confundir aunque alguien pinche otra
> ventana de Word a media corrida. La guarda sigue viva para el uso a mano, que es donde el
> destino todavía se adivina.

**14. PowerShell REDONDEA al castear a `[int]`; C# TRUNCA.**

```
PowerShell:  [int]9.6        ->  10      (redondeo, y al PAR en los empates)
C#:          (int)9.6        ->   9      (trunca hacia cero)
```

Salió al portar el muestreo de pantalla de PowerShell a C# el 04/09/2026. La rejilla de
`Test-CapturasDistintas` calculaba `[int](($j + 0.5) * $alto / $rejilla)`; traducido a C# con
un `(int)` **caía en otros píxeles**. Con la ventana en 1366x768 el primer punto pasaba de
`(17,10)` a `(17,9)`.

Lo venenoso es que **no falla: da resultados parecidos.** La comparación seguía diciendo
"cambió" o "no cambió" casi siempre bien, y sobre imágenes distintas los dos muestreos
diferían en 3 puntos de 1600. Nada que se note mirando una corrida; solo se ve comparando
los dos muestreos sobre el mismo par de imágenes.

Si hay que replicar el redondeo de PowerShell en C#:

```csharp
(int)Math.Round(valor, MidpointRounding.ToEven)   // no (int)valor
```

Vale para cualquier traducción entre los dos lenguajes, no solo para píxeles: cualquier
índice, fila o coordenada calculada con `[int]` cambia de valor al pasar a `(int)`.

Lo caza `diagnostico\probar_captura.ps1`, que corre los dos muestreos —el viejo por
`GetPixel` y el nuevo por `LockBits`— sobre el mismo par de imágenes y exige que den el
**mismo número**. No toca Azul: se puede correr siempre.

**15. La codificación de los archivos que NO son `.ps1`: el CSV de entrada, y el arnés
que se miente a sí mismo.**

No es la trampa 1. La 1 va de los `.ps1` del proyecto y se resuelve con ASCII puro. Ésta va
de dos sitios donde el ASCII puro **no** te salva: el CSV que entra, y los propios scripts de
prueba.

**a) El CSV de la lista de órdenes.** Medido el 05/09/2026 sobre las tres codificaciones,
escribiendo `DISTRIBUIDORA RÍO` y volviéndola a leer:

| Cómo llega el archivo | Resultado |
|---|---|
| ANSI / Windows-1252 — *CSV (delimitado por comas)* de Excel | correcto |
| UTF-8 **con** BOM — *CSV UTF-8* de Excel | correcto |
| UTF-8 **sin** BOM | `DISTRIBUIDORA RÃO`, **en silencio** |

Excel no exporta el tercer caso. **El Bloc de notas de Windows 11 sí**, y abrir el CSV para
comprobarlo con los ojos antes de una corrida es justo lo que uno hace.

Importa porque esa cadena tiene **dos destinos**: la guarda que compara la razón social
contra la fila que trae Azul —que fallaría hacia el lado seguro, abortando— y **el encabezado
del documento que ve el equipo comercial**, donde el nombre del cliente saldría mal escrito.

La regla: **no forzar `-Encoding`, detectar.** `azul\lista.ps1` mira el BOM; si no lo hay y
hay bytes altos, comprueba si cumplen la estructura de UTF-8 (un byte líder y sus
continuaciones `10xxxxxx`) y decide. El texto ANSI con acentos casi nunca la cumple —`MÍA`
son los bytes `CD 41`, y `CD` exige una continuación entre `80` y `BF`, que `A` no es—, así
que la distinción es real y no una adivinanza. Cuando toca leer UTF-8 sin BOM, **lo dice**.

**b) El arnés que se miente. Éste es el que va a volver.**

La primera prueba de esto dio **`OK` siendo falso**. El script de prueba estaba en UTF-8,
PowerShell 5.1 lo leyó como ANSI (trampa 1) y **rompió por igual el literal esperado y el que
la prueba escribía al archivo**. Los dos lados quedaron corrompidos de la misma forma, la
comparación cuadró, y la prueba certificó como correcto exactamente lo que estaba roto.

> **Una prueba de codificación escrita con literales acentuados no prueba nada.**
> Construye los caracteres desde su código en tiempo de ejecución:
> `$I = [char]0xCD`, y la cadena esperada se arma con él. Así la codificación del archivo de
> prueba no puede contaminar el resultado.

Al arreglar el lector, la misma prueba pasó de `OK` falso a **`FALLA` falso** —el código ya
devolvía la cadena correcta y el literal de la prueba seguía roto—. Ése es el síntoma: **una
prueba de codificación que cambia de veredicto cuando arreglas el código, sin que nadie haya
tocado la prueba, está midiendo su propio archivo y no el tuyo.**

Lo caza `diagnostico\probar_lista.ps1`, que construye los acentos por código de carácter y
prueba las tres codificaciones. No toca Azul, ni Excel, ni Word.

**16. Un solo elemento no es una lista de uno: `.Count` llega VACÍO, no `1`.**

`ConvertFrom-Csv` —y cualquier tubería— devuelve un **objeto suelto** cuando le llega una sola
línea. Ese objeto no tiene `.Count`, así que `$filas.Count + 1` da **1** en vez de 2, y con dos
o más elementos da lo correcto. Por eso no se ve: falla **solo** en el caso de uno.

```
una linea  -> PSCustomObject   .Count vacío   .Count + 1 = 1
dos lineas -> Object[]         .Count = 2     .Count + 1 = 3
```

Costó (09/09/2026) que los bloques de **una sola línea** salieran en el documento sin su fila
de encabezados. Se envuelve siempre en `@()`: `@($algo | ConvertFrom-Csv ...)`.

**La regla general: cualquier prueba que solo use dos o más elementos no prueba nada sobre
uno.** El CSV sintético de `probar_word.ps1` trae tres bloques de una sola línea a propósito.

---

## Otras condiciones que rompen la corrida

Documentadas aparte de las 16. Salieron de `LEEME.md` y `PROYECTO-AZUL.md`, ya retirados
del proyecto, y de las memorias.

**Del Java Access Bridge:**

- **Seleccionar la fila por el radio no basta.** `doAccessibleActions` sobre el radio
  deja el widget `checked` pero **no dispara la lógica que habilita `Ver Productos
  Asignados`**. Hay que usar `addAccessibleSelectionFromContext(vm, tabla, fila *
  columnas)` — índice de **celda, no de fila**. Un clic real también funcionaría, pero
  las celdas de tabla reportan bounds `-1`: no hay coordenadas que usar.
- **Los nombres de atributo no están en `name`,** sino en `description` envuelto en
  HTML. Sin esto, todo sale `sin nodo Plan Móvil`. Ver `docs/mapa-datos.md` §4.
- **Nunca recorrer el árbol dentro de un bucle de espera.** Cada llamada JAB es IPC
  síncrono que se bloquea mientras Azul carga. Recorrerlo 30-40 veces por línea daba
  **11 minutos con 20 s de CPU**. Cachear contextos; el `Find` **no debe descender
  dentro de nodos `table`** (ahí están las miles de celdas).
- **El árbol de atributos carga perezosamente.** Leerlo apenas abre da 20 filas en vez
  de 120. Esperar a que `rowCount` **se estabilice**, no a que supere un umbral.
  Esta espera **no se tocó al optimizar el lector y no se debe tocar**: ahí no hay
  condición fiable que preguntar.

**Del estado de la ventana:**

- **Azul minimizado** → Swing no crea la subventana y el puente no la ve. Puede estar
  detrás de otras ventanas, pero **no iconizado**.
- **Con el detalle abierto, la tabla de Suscripciones no es alcanzable** en el árbol.
  Hay que cerrarlo antes de buscarla.
- **Si se cambia de cuenta a media corrida**, los contextos quedan muertos. Hay una
  guarda que compara el número de la fila contra el capturado al inicio y aborta.

**De la lectura del CSV / Word:**

- `progreso.txt` se escribe en vivo, una línea por paso. **Ábrelo con UTF-8** o los
  acentos se ven rotos.
- La base de Word **se guarda al crearla**, no al llegar a 10 clientes:
  `agregar_a_word.ps1` se niega a escribir en un documento sin ruta, y así una caída a
  mitad no se lleva lo escrito.

---

## Guardas que no se quitan

Todas nacieron de fallos reales (ver `docs/bitacora.md`).

- Abortar si Azul está minimizado.
- Cerrar cualquier detalle abierto antes de buscar la tabla de Suscripciones.
- Abortar si la cuenta cambia a media corrida.
- Verificar que el título de la subventana traiga el número de la línea pedida.
- Más de un `MPE` o más de una `Fecha Final` → **no elegir**: `REVISAR`.
- Esperar a que el árbol de atributos deje de crecer antes de leerlo.
- `MPE = 0` y `MPE = -1` se reportan tal cual.
- **Nunca adivinar.** Ante la duda, `REVISAR` con el motivo.
- No escribir en Word si la corrida quedó incompleta.

**Añadidas el 06/09/2026 con la simplificación.** No nacieron de un fallo pasado sino de un
riesgo que el cambio de alcance *crea*, que es igual de válido y más fácil de olvidar:

- **Cerrar el panel del cliente si quedó abierto, antes de leer.** Vuelve a abrirlo el
  sistema (captura automática, desde el 08/09/2026), y ese panel tapa la tabla de
  Suscripciones. La guarda sigue haciendo falta: si un tope de tiempo mata la corrida a
  media captura, el panel puede quedar abierto para el cliente siguiente.
- ~~**Comprobar el documento destino ANTES del primer cliente**~~ — **ya no rige desde el
  09/09/2026, y no se abre Word a mano.** Nació cuando el documento lo elegía una persona y
  había que averiguar cuál de los abiertos era el bueno. Hoy el lote crea la base, la abre,
  escribe y la cierra él solo. No hay documento que comprobar. La comprobación sigue viva
  únicamente en `agregar_a_word.ps1` cuando se le llama suelto y sin decirle el documento,
  que es como se usaba antes.
- **Denunciar en voz alta toda fila descartada de la lista de órdenes.** Con el libro vivo el
  descarte silencioso solo se comía el encabezado; con un CSV se come clientes.
- **Avisar cuando la razón social viene vacía.** La guarda que impide entrar al cliente
  equivocado es condicional: sin ese dato no se aplica, y callarlo la deja apagada en secreto.

---

## Diagnóstico — qué correr cuando algo no cuadra

| Herramienta | Para qué |
|---|---|
| `diagnostico\verificar.ps1` | ASCII, parseo y compilación del C# de los tres archivos que lo llevan. **Antes de nada.** |
| `diagnostico\probar_captura.ps1` | Prueba el C# de `captura.ps1` sobre imágenes sintéticas: que el muestreo nuevo por `LockBits` dé el mismo número que el viejo por `GetPixel` (trampa 14), y que la garantía de pantalla estable no baje de 1.4 s. **No toca Azul.** |
| `diagnostico\probar_lista.ps1` | Prueba el lector de la lista de órdenes (`lista.ps1`): las tres codificaciones (trampa 15), el separador, las cuentas repetidas, las filas descartadas y las columnas que faltan. **No toca Azul, ni Excel, ni Word.** |
| `diagnostico\probar_word.ps1` | Prueba la forma del bloque que se escribe en el documento: orden, el hueco `[imagen]`, y que las tablas no se fundan (trampa 12). Levanta su propio Word y su propio documento. **Ya no se niega si Word está abierto** (cambio del 09/09/2026): avisa, y no toca nada de lo que tengas abierto. **No toca Azul.** |
| `diagnostico\jab_now.ps1` | Estado actual de Azul: marcos, tablas y si el botón está habilitado. Lo primero cuando algo no cuadra. |
| `diagnostico\jab_attrs.ps1` | Vuelca los nombres reales de los atributos. Cuando todo sale `sin nodo Plan Móvil`. |
| `diagnostico\jab_cliente.ps1` | Busca nodos en el árbol (`-Buscar <texto>`, `-Todo`). Prueba decisiva de si un widget existe para el puente. |
| `diagnostico\zoom.ps1` | Recorta y amplía un pedazo de la ventana con rejilla rotulada en coordenadas de origen. Para medir un ícono al píxel. Acepta `-Fuente` para medir sobre una captura vieja, sin Azul abierto. |
| `diagnostico\excel_ordenes.ps1` | Lista libros, hojas y columnas del Excel abierto. **Solo para depurar una exportación que salió rara**, y entonces sí hay que tener el libro abierto. **Excel no interviene en una corrida**: la lista sale de un CSV en disco desde el 05/09/2026. |

La tabla de **mensajes de error concretos** y qué hacer con cada uno nunca se mudó aquí, y
`azul\LEEME.md` ya no está en el proyecto. Sigue entera en el historial, en la sección
*Si algo falla*:

```powershell
git show 4e895de:azul/LEEME.md
```

---

*Procedencia: `azul\CONTEXTO-sesion-03-09-2026.md` (§5, las 13), `azul\CONTEXTO-sesion-02-09-2026.md`
(§3), `azul\LEEME.md`, `PROYECTO-AZUL.md`, `azul\PROMPT-actualizar-reglas.md`, y las
memorias `azul-lectura-tecnica`, `azul-vista-cliente`.*
