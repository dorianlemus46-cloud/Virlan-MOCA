# INSTALAR — poner el proyecto a andar en una máquina nueva

Para quien clona el repositorio por primera vez. Una sola vez por máquina.
Cuando termines, el punto de entrada es [`ARRANQUE.md`](ARRANQUE.md).

> **Lee [`REGLAS.md`](REGLAS.md) antes de correr nada contra Azul.** Azul es producción sobre
> cuentas de clientes reales. Aquí se explica cómo instalar, no qué está permitido hacer.

---

## 1. Lo que tiene que haber ya en la máquina

| Qué | Por qué | Cómo compruebas que está |
|---|---|---|
| **Azul** instalado y con sesión iniciada | Es lo que se lee | La ventana *Ejecutivo de interacción del cliente de AT&T* abierta |
| **Microsoft Word** | Ahí se escriben las bases | Que esté instalado. **No hay que abrirlo** |
| **Windows PowerShell 5.1** | Es el intérprete de todo el proyecto | `$PSVersionTable.PSVersion` |
| **El JRE de 32 bits de Azul** | El puente de accesibilidad vive ahí | Existe `C:\Program Files (x86)\Java\` |

No hay nada que instalar del proyecto: no usa paquetes, ni módulos externos, ni
compilación previa. Se clona y se corre.

> **Solo Azul se abre a mano.** Word lo levanta el lote cuando le toca escribir, y lo cierra
> al terminar. Excel se usa una vez, aparte, para exportar la lista a CSV, y después no
> interviene. Si algo te pide abrir Excel o Word para correr una corrida, está
> desactualizado: son restos del sistema anterior.

## 2. Habilitar el Java Access Bridge

Es el único paso de configuración del sistema, y sin él no funciona nada. El puente es
lo que deja leer la interfaz de Azul como texto en vez de adivinarla con imágenes.

```powershell
& "C:\Program Files (x86)\Java\jre1.8.0_451\bin\jabswitch.exe" -enable
```

**La versión del JRE cambia de una máquina a otra.** Lo que no cambia es que sea el de 32
bits, el de `Program Files (x86)`. Si esa ruta no existe tal cual, búscala:

```powershell
Get-ChildItem "C:\Program Files (x86)\Java" -Recurse -Filter jabswitch.exe -ErrorAction SilentlyContinue
```

Crea `.accessibility.properties` en el perfil del usuario. **Hay que reiniciar Azul
después**; no hace falta cerrar sesión de Windows. Se revierte con `jabswitch -disable`.

Comprueba que el puente quedó puesto:

```powershell
Test-Path (Join-Path $env:USERPROFILE ".accessibility.properties")
```

## 3. Clonar y comprobar

Clona donde quieras: el proyecto no depende de la carpeta ni del nombre de usuario.
**Todos los comandos se lanzan desde la raíz del proyecto**, la carpeta que contiene
`azul\` y `docs\`.

Los cuatro gates. No tocan Azul, ni Word, ni ningún dato real, así que puedes correrlos
ahora mismo. Los cuatro tienen que decir `TODO OK`:

```powershell
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\verificar.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\probar_bases.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\probar_lista.ps1"
& "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\probar_captura.ps1"
```

Si alguno falla, **para aquí**. Ninguno de ellos necesita Azul, así que un fallo es del
entorno o del código, y correr contra Azul con un gate en rojo es exactamente lo que estas
comprobaciones existen para impedir.

Hay un quinto, `probar_word.ps1`, que levanta su propio Word con un documento propio y lo
cierra al terminar. **No hace falta que cierres el Word que tengas abierto**: avisa si lo
detecta y no toca nada tuyo. Córrelo cuando puedas dejar la pantalla un momento.

## 4. La lista de órdenes

Es la puerta por la que entran las cuentas. **No viaja en el repositorio**: lleva números
de cuenta y razones sociales de clientes reales, así que está excluida en
[`.gitignore`](.gitignore) y nunca debe subirse.

En su lugar va [`azul/lista_ordenes.ejemplo.csv`](azul/lista_ordenes.ejemplo.csv), con la
forma exacta y datos inventados. Para trabajar de verdad, exporta tu lista con **Guardar
como** sobre la hoja `resultData`, sin borrar columnas, y déjala en `azul\lista_ordenes.csv`.

El lector busca **por encabezado**, no por posición, así que las columnas que sobran no
estorban. Las dos que importan:

- **`Cuenta`** — sin ella el lector aborta.
- **`Razón Social`** — sin ella arranca igual, pero **avisa de que queda apagada la guarda
  que impide entrar al cliente equivocado**. No trabajes así.

Pruébala sin tocar Azul ni Word antes de la primera corrida:

```powershell
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\lote.ps1" -Simular
```

## 5. La carpeta de las bases

Se crea sola, en `azul\bases`, la primera vez que corre el lote. No hay nada que preparar.
Como la carpeta lleva documentos de clientes reales, está excluida entera del repositorio,
y con ella el registro que numera las bases.

> **Ojo con el correlativo si hay más de una máquina.** El número de base vive en el
> registro local, que no viaja en el repositorio. Una máquina recién clonada empieza a
> numerar desde la 034 otra vez, sin saber lo que haya hecho la otra. Dos máquinas
> trabajando en paralelo producirán cada una su propia BASE 034. **Pendiente de resolver**;
> mientras tanto, que numere una sola.

## 6. Antes de la primera corrida contra Azul

Dos cosas que dependen de esta máquina y no del código:

**El tamaño de la ventana.** El proyecto tiene exactamente dos coordenadas de clic
clavadas, y **están medidas con Azul en 1366x768** ([`REGLAS.md`](REGLAS.md) §5). Si esta
máquina usa otra resolución o la ventana va a otro tamaño, dejan de acertar y el clic cae
sobre lo que haya en ese píxel. Se remiden con `diagnostico\zoom.ps1`, que amplía la zona
con una rejilla rotulada. El sistema nunca da por bueno un clic clavado —comprueba el
efecto y aborta si no lo ve—, pero eso evita el daño, no el fallo.

**Empieza corto.** La primera vez, `-Limite 1` y con alguien delante.

---

## Lo que nunca se sube

Está resuelto en [`.gitignore`](.gitignore) y conviene saber por qué: las capturas, los CSV
por cliente, `progreso.txt`, la lista de órdenes y la carpeta de bases llevan nombres,
números de línea, números de cuenta y saldos de clientes reales. Subir eso **no se
deshace**: queda en el historial, en los forks y en las cachés aunque después se borre.

Los nombres y números que aparecen en el código y en la documentación son ficticios.
