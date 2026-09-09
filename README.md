# Proyecto Azul

Automatiza la lectura de suscripciones en **Azul** (Siebel sobre Java/Swing) para armar
bases de datos de renovación en Word. Lee por el **Java Access Bridge**; solo toca el ratón
donde el puente no llega.

## Por dónde se entra

Si acabas de clonar el repositorio en una máquina nueva: **[`INSTALAR.md`](INSTALAR.md)**,
una vez. Después, y siempre, **[`ARRANQUE.md`](ARRANQUE.md)**. Lo demás se lee cuando haga
falta.

| Archivo | Para qué |
|---|---|
| [`INSTALAR.md`](INSTALAR.md) | Solo la primera vez en cada máquina: el Java Access Bridge, los gates, la lista de órdenes |
| [`ARRANQUE.md`](ARRANQUE.md) | El punto de entrada |
| [`REGLAS.md`](REGLAS.md) | Qué se puede tocar en Azul y qué no. **Azul es producción** |
| [`docs/trampas.md`](docs/trampas.md) | Las 16 trampas. Cada una costó una corrida perdida o peor |
| [`docs/decisiones.md`](docs/decisiones.md) | Por qué el sistema es como es |
| [`docs/mapa-datos.md`](docs/mapa-datos.md) | Dónde vive cada dato en el árbol de accesibilidad |
| [`docs/bitacora.md`](docs/bitacora.md) | Fallos reales y cómo se arreglaron |

## Antes de correr nada

```powershell
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\azul\diagnostico\verificar.ps1"
```

Comprueba ASCII, parseo y compila el C# de los tres archivos que lo llevan, **sin tocar
Azul**. Tiene que decir `TODO OK`.

## Dos cosas que no son negociables

**Azul es de solo lectura en cuanto a datos del cliente.** Se navega y se consulta; no se
guarda, ni se aprueba, ni se descarta, ni se exporta nada. Los ciclos autorizados y la lista
blanca de lo que se puede pulsar están en [`REGLAS.md`](REGLAS.md).

**Los `.ps1` van en ASCII puro.** PowerShell 5.1 lee los `.ps1` sin BOM como ANSI y rompe los
acentos en silencio, lo que manda todas las líneas a REVISAR sin decir por qué. Los literales
con acento van como escapes `\uXXXX`. Es la trampa 1 y la comprueba `verificar.ps1`.

---

## Sobre los datos de este repositorio

**Los nombres y números que aparecen en el código y en la documentación son ficticios.** Los
ejemplos reales se sustituyeron por equivalentes inventados conservando su forma —el largo de
un nombre, el prefijo de una cuenta y la posición de un acento son parte de lo que explican
algunas trampas—, pero no corresponden a ninguna cuenta, línea ni persona real.

> **La sustitución quedó incompleta hasta el 09/09/2026.** Cuatro números de cuenta y una
> razón social se habían escapado y siguieron en el repositorio, y en el historial, desde el
> commit de la Etapa 0. Se sustituyeron al detectarlos.
>
> **Cómo se detectan, si vuelve a pasar.** No basta con leer los ejemplos y que parezcan
> inventados: los que se escaparon lo parecían. Lo que los delata es cruzarlos con
> `azul\salidas\`, que son corridas reales y llevan la cuenta en el nombre del archivo. Un
> número inventado no coincide con una corrida real por casualidad. Los ejemplos buenos se
> reconocen aparte por su forma: terminan en `0000001`, `0000002` y así.

Las salidas de una corrida —capturas de pantalla, CSV por cliente, `progreso.txt`— **sí**
llevan datos reales y por eso están excluidas en [`.gitignore`](.gitignore). No se suben, y
no deben subirse: quedan en el historial, en los forks y en las cachés aunque luego se
borren.
