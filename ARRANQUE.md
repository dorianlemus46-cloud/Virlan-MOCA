# ARRANQUE — Proyecto Azul

Lo único que se lee siempre. Todo lo demás se abre solo cuando hace falta.

## Cómo hablarme

Explicaciones simples, sin nombres de función ni de archivo salvo que los pida. Un paso a
la vez, no la lista completa de lo que vas a hacer. Si el cambio ya está hecho y
`verificar.ps1` dice `TODO OK`, corre el bloque y enseña el resultado directo — no
encadenes comprobaciones intermedias antes de mostrar algo. Y al informar sobre el
sistema, distingue **siempre** lo que está probado contra Azul real de lo que solo está
escrito: es la distinción que más importa en este proyecto.

## Qué es y para qué sirve

Dorian trabaja en una empresa de telecomunicaciones afiliada a AT&T que **renueva
contratos de líneas para negocios**. El ciclo es: una lista de órdenes en CSV → buscar
cada cuenta en el CRM **Azul** → leer qué líneas están cerca de vencer → volcarlo a un
documento de **Word** para el equipo comercial. Ahí termina el sistema: **Dorian decide
después, a mano, verificando contra ADS**, si cada renovación se aprueba o se descarta.
Azul nunca decide nada y el sistema nunca aprueba nada por su cuenta.

**Bloques cortos, con Dorian delante.** No es un lote desatendido: desde el 06/09/2026 el
alcance es tres a cinco clientes por tirada, él presente. El sistema hace lo que a mano
cuesta caro: buscar la cuenta, leer línea por línea, clasificar, escribir las tablas, y
—desde el 09/09/2026— **crear, nombrar y cerrar las bases él solo**.

**Azul** es una aplicación **Java (Swing)** sobre Siebel (ventana *Ejecutivo de
interacción del cliente de AT&T*, proceso `jp2launcher`), no es web. El **Java Access
Bridge lee la interfaz perfectamente, pero no la conduce**: pulsar un botón o escribir un
campo por el puente puede devolver éxito sin que Siebel se entere. Por eso **leer va por
el puente** y **escribir y pulsar van con teclado y mouse reales** — y por eso las reglas
de abajo son tan estrictas sobre qué se puede pulsar. El resto de por qués de diseño está
en `docs/decisiones.md`.

> **El número de servidor cambia entre sesiones y no identifica nada.** Este archivo decía
> `CRMServer_6_2`; el 04/09/2026 la ventana viva decía `CRMServer_6_1`. **El código nunca lo
> usa**: Azul se localiza siempre por `Get-Process -Name jp2launcher` con
> `MainWindowHandle -ne 0`, jamás por el título. No introducir esa dependencia.

## Seguridad — resumen. El archivo completo es `REGLAS.md`

Azul es **solo lectura** en cuanto a datos del cliente. Hay **tres ciclos autorizados**
—leer una línea, ver la vista general del cliente, buscar una cuenta— y **nada más**.
Prohibido guardar, aprobar, descartar, eliminar, modificar y exportar.
Nunca se pulsa `Crear`, `Exportar`, `Guardar`, `Aprobar`, `Descartar` ni `Eliminar`.

**Ante cualquier duda: no hacer clic. Abortar y preguntar.**

> `REGLAS.md` **no se modifica sin autorización expresa de Dorian.**

## Cómo se corre

**Todos los comandos se lanzan desde la raíz del proyecto**, la carpeta que contiene
`azul\` y `docs\`. Las rutas son relativas a propósito: el proyecto no depende de en qué
carpeta ni bajo qué usuario esté clonado.

Antes de correr nada, **siempre**. Debe decir `TODO OK`:

```powershell
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\verificar.ps1"
```

```powershell
# el bloque completo (lo normal)
# admite -Simular  -Limite N  -Lista <ruta>  -VerBases  -CerrarBase  -Documento <ruta>
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\lote.ps1"

# un solo cliente, a mano — SIEMPRE con el PowerShell de 32 bits ($ps32)
$ps32 = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
& $ps32 -NoProfile -ExecutionPolicy Bypass -File ".\azul\buscar_cuenta.ps1" -Fila 6 -Restaurar
& $ps32 -NoProfile -ExecutionPolicy Bypass -File ".\azul\azul_fast.ps1"
```

**Condiciones:** Azul abierto y **no minimizado** (puede estar detrás de otras ventanas), y
`azul\lista_ordenes.csv` exportado desde Excel (*hoja `resultData` → Guardar como → CSV*, sin
borrar columnas). **Word ya no hace falta abrirlo**: el lote lo levanta si no está, escribe en
la base que toca y la cierra al terminar. Si ya tenías Word abierto, usa ese y no te lo cierra.

> **Ni Excel ni Word se abren a mano. La única aplicación que tiene que estar abierta es
> Azul.** Excel se usa una vez, antes y aparte, para exportar el CSV; después no interviene y
> no hace falta que siga abierto. Se dice aquí porque la documentación vieja pedía las dos
> cosas y todavía quedan rastros: si algo te manda abrir Excel o Word para correr el lote,
> está desactualizado.
**No tocar Azul mientras corre.** Todo lo que toque el puente va en el PowerShell de
**32 bits**; `lote.ps1` corre en el de 64 y lanza a los hijos en el de 32.
**Bash está bloqueado en esta máquina**: todo por PowerShell.

## Cómo se arma una base — la regla de los 10

**Una base son 10 clientes, y desde el 09/09/2026 el corte lo lleva el código otra vez.**

Las bases viven en **`azul\bases`**, se llaman `BASE 034 PS1 VIRLAN.docx`, `035`, `036`… y no
hay que abrir nada ni nombrar nada. Como una tirada son tres a cinco clientes y una base son
10, **una base cruza varias corridas**: el lote continúa la que quedó a medias y la cierra
sola al llegar a 10.

**Hay dos series, y cada una lleva su propia cuenta** (desde el 18/09/2026): la **PS1** de
siempre y la **PS2**, que empezó en `BASE 071 PS2 VIRLAN.docx`. Numerar en una no mueve a
la otra. Sin `-Serie` es la PS1; cuál toca en cada momento lo dice Dorian.

```powershell
lote.ps1 -Serie PS2    # numerar en la serie PS2; vale también con -VerBases, -CerrarBase...
lote.ps1 -VerBases     # qué base está abierta, quién va dentro, cuáles están cerradas
lote.ps1 -CerrarBase   # cerrarla ya, aunque no haya llegado a 10

# adoptar una base que empezaste a mano: sin -Confirmar solo mira y cuenta
lote.ps1 -Adoptar "C:\...\BASE 036 PS1.docx"
lote.ps1 -Adoptar "C:\...\BASE 036 PS1.docx" -Confirmar
```

**Adoptar copia, no mueve.** El original se queda donde está y pasa a ser una copia vieja:
archívalo para no seguir escribiendo en él por error. Y **guárdalo antes de confirmar**, porque
la copia sale del disco y se quedaría sin lo último que escribiste. El sistema se planta si lo
tienes abierto con cambios sin guardar.

**Un cliente sin nada que renovar no entra al documento y no cuenta para los 10.** Queda
anotado en `salidas\lote_progreso.csv` con el motivo, así que no se pierde.

**Si relanzas la misma lista, no se duplica nada:** el sistema sabe quién está ya dentro de la
base abierta y se lo salta antes de buscarlo en Azul.

> **El correlativo ya no se deduce, se recuerda.** Del 06/09 al 09/09/2026 el sistema no
> nombró bases, y por un motivo bueno: las nombraba mirando qué archivos había en la carpeta
> —su única memoria— así que en cuanto renombrabas un documento terminado, ese número quedaba
> libre y la corrida siguiente lo reutilizaba. Eso **no ha vuelto**. Ahora el número sale de
> un registro que solo avanza, nunca se deduce de la carpeta, y jamás se sobrescribe un
> documento que ya exista. El porqué completo está en `docs/decisiones.md`.

## En qué estado está

La búsqueda de cuenta, la lectura de líneas y una corrida completa de dos clientes están
**probadas contra Azul real** — fechas y detalle en `ESTADO.md`.

> ### ⚠ La simplificación del 06/09/2026 NO se ha ejercitado contra Azul.
>
> Seis etapas, rama `simplificacion`, escritas y verificadas **sin Azul delante**.
> **`ESTADO.md` abre con la lista de todo lo que falta probar, en orden y con cómo
> probarlo.** Léela antes de correr nada. La primera mitad no toca Azul.
>
> **Las bases automáticas del 09/09/2026 tampoco.** Lo que sí está probado, y a fondo, es
> todo lo que no necesita Azul: la memoria del correlativo y el ciclo entero de una base en
> Word —crearla, continuarla en otra corrida, cerrarla al llegar al tope—. Lo que falta es
> verlo con clientes de verdad saliendo de Azul.

## Índice — abre solo lo que necesites

| Archivo | Cuándo abrirlo |
|---|---|
| `INSTALAR.md` | **Solo la primera vez en cada máquina.** Habilitar el Java Access Bridge, correr los gates, preparar la lista de órdenes. |
| `ESTADO.md` | **Al empezar, siempre.** Qué falta probar contra Azul, en orden; después, qué está probado y qué solo escrito. Se sobrescribe cada sesión. |
| `REGLAS.md` | **Antes de cualquier acción que toque Azul.** Prohibiciones, ciclos autorizados, botones vetados y las dos coordenadas clavadas. |
| `docs/decisiones.md` | Por qué el sistema es así y no de otra forma: Java Access Bridge y no OCR, clics reales y no por el puente, y qué se rompió al intentar lo contrario. |
| `docs/plan-simplificacion.md` | Por qué el sistema es corto y no largo: qué se quitó en cada etapa, qué se ganó, qué se perdió. |
| `docs/mapa-datos.md` | Solo cuando toques la lectura: dónde vive cada dato, y la regla de renovación. |
| `docs/trampas.md` | Solo cuando algo falle: las 16 trampas y las guardas que no se quitan. |
| `docs/bitacora.md` | Cuando algo se rompa: qué se rompió antes y cómo se arregló. Una entrada por incidente. |

**Retirados del proyecto.** Cinco archivos describían el sistema anterior y ya no lo
describen: `PROYECTO-AZUL.md`, `azul/LEEME.md`, `azul/CONTEXTO-sesion-02-09-2026.md`,
`azul/CONTEXTO-sesion-03-09-2026.md` y `azul/PROMPT-actualizar-reglas.md`. Nada se perdió
—siguen enteros en el historial— pero ya no estorban a quien clona el repositorio hoy.
Para leer cualquiera de ellos:

```powershell
git log --oneline --diff-filter=D -- PROYECTO-AZUL.md   # el commit que lo retiro
git show <commit>^:PROYECTO-AZUL.md                      # el archivo, tal cual estaba
```
