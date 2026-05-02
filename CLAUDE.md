# TOS Alert Formatter — CLAUDE.md

> Guía de referencia para Claude Code. Documenta arquitectura, decisiones de diseño, convenciones y estado del proyecto.

---

## 1. Descripción general

Aplicación web de **una sola página** (`index.html`) desplegada en GitHub Pages.  
Convierte alertas de órdenes de ThinkorSwim (TOS) en texto listo para pegar en TOS, con tres variantes de precio (+0 / +1c / +5c), y genera automáticamente URLs para OptionStrat y strings para el TOS Chart.

**URL de producción:** `https://nfpesce.github.io/tos-alert-formatter/`  
**Repositorio:** `https://github.com/nfpesce/tos-alert-formatter`  
**Branch de producción:** `master` (GitHub Pages sirve directamente desde `master`)

---

## 2. Archivos del proyecto

| Archivo | Rol |
|---|---|
| `index.html` | Producción. Tiene todos los features activos. |
| `index_v3.html` | Referencia histórica de v3 (espejo de index.html durante desarrollo). |
| `index_v2.html` | Referencia histórica de v2. No se modifica. |

> **Convención de deploy:** Los cambios van siempre a `index.html`. `index_v3.html` se actualiza en paralelo como snapshot histórico. Al terminar una sesión de desarrollo se hace commit + push a `master`.

---

## 3. Arquitectura

Todo el código vive en un único `index.html`. No hay bundler, no hay dependencias npm, no hay módulos ES. La estructura interna del `<script>` está dividida en secciones numeradas:

```
Section 1  — formatAlert()          Limpieza y normalización del texto de alerta
Section 2  — adjustPrice()          Variantes de precio (+1c, +5c)
Section 3a — parseSingleLineStrategy()  Parser de estrategias TOS
Section 3b — Black-Scholes pricing  Calibración de IV y distribución de precios por leg
Section 3c — optionStratURLAsync()  Generación de URL OptionStrat (async, usa BS)
Section 3d — Expiration P/L chart   Canvas 2D para gráfico de P&L al vencimiento
Section 3e — parseOptionStratURL()  Parser URL OptionStrat → TOS string + parsed object
Section 4  — UI Logic               processInput(), selectCard(), renderChart(), etc.
```

### Flujo principal al pegar una alerta TOS

```
paste → processInput()
    ├── formatAlert()           → limpia texto
    ├── parseSingleLineStrategy()  → extrae legs, ticker, dateCode, strategy, isWeekly
    ├── adjustPrice() x2        → genera +1c y +5c
    ├── selectCard()            → copia al clipboard (PRIORIDAD MÁXIMA, síncrono)
    ├── renderPayoffForOrder()  → gráfico P&L canvas (async, no bloquea clipboard)
    └── setTimeout(renderChart) → widget TradingView (diferido, no bloquea clipboard)
```

### Flujo alternativo: URL de OptionStrat como input

```
paste URL → processInput()
    ├── isOptionStratURL()      → detecta formato https://optionstrat.com/build/…
    ├── parseOptionStratURL()   → builds tosFormatted (@0.00) + parsed object DIRECTAMENTE
    │       (bypasses formatAlert y parseSingleLineStrategy)
    ├── adjustPrice() x2        → genera +.01 y +.05
    ├── selectCard()            → copia al clipboard (PRIORIDAD MÁXIMA)
    ├── renderPayoffForOrder()  → gráfico P&L con @0.00 (muestra estructura sin costo)
    │       (calendars/diagonals: multi-date → hidePayoffChart silencioso)
    └── setTimeout(renderChart) → widget TradingView (diferido)
```

`currentOptionStratURL` guarda la URL original → botón "Open in OptionStrat" la reabre directamente (sin recomputar).

---

## 4. Estrategias soportadas

El parser (`parseSingleLineStrategy`) reconoce estos valores en `strategy`:

| Valor | Descripción | URL OptionStrat |
|---|---|---|
| `SIMPLE` | Una sola pata | `custom/` con precio B-S |
| `VERTICAL` | Spread de 2 patas | `custom/` con precio B-S |
| `BUTTERFLY` | Butterfly simétrico o broken-wing | `long/short-put/call-butterfly` o `put/call-broken-wing` |
| `~BUTTERFLY` | Broken-wing/ratio explícito (TOS usa `~`) | `put/call-broken-wing` |
| `BACKRATIO` | Back ratio spread | `custom/` con precio B-S |
| `CONDOR` | Condor single-type | `long/short-call/put-condor` (sin precios) |
| `IRON_CONDOR` | Iron Condor PUT+CALL | `iron-condor/` (sin precios) |
| `DIAGONAL` / `CALENDAR` | Diagonal/calendar | `custom/` con precio B-S |

### Detección de Butterfly simétrico vs broken-wing (BUTTERFLY)

```javascript
const isSymmetric = Math.abs((strikes[1]-strikes[0]) - (strikes[2]-strikes[1])) < 0.01;
// simétrico  → "long-put-butterfly" / "short-put-butterfly"
// asimétrico → "put-broken-wing"
```

`~BUTTERFLY` siempre mapea a `broken-wing` sin calcular anchos.

---

## 5. Decisiones de diseño clave

### 5.1 Clipboard es máxima prioridad
El widget de TradingView se carga con `setTimeout(..., 0)` para encolarlo **después** de que `selectCard()` copie al clipboard. Esto garantiza que al pegar rápidamente en TOS la orden siempre esté lista.

### 5.2 Black-Scholes para distribución de precios por leg (v2)
Para estrategias multi-pata (VERTICAL, BUTTERFLY, etc.) el precio total de la alerta se distribuye entre legs usando calibración de IV implícita con B-S, en lugar de asignar todo al primer leg (comportamiento v1). Requiere fetch de precio del subyacente vía Finnhub.

### 5.3 Índices no usan Finnhub directamente
`fetchStockPrice` mapea índices a ETFs proxy con factor de escala:
```javascript
const INDEX_TO_ETF = {
    'SPX':  { etf: 'SPY',  factor: 10  },
    'SPXW': { etf: 'SPY',  factor: 10  },
    'NDX':  { etf: 'QQQ',  factor: 47  },
    'RUT':  { etf: 'IWM',  factor: 10  },
    'DJX':  { etf: 'DIA',  factor: 100 },
};
```
El gráfico P&L de índices no muestra el marcador de precio subyacente (Finnhub free plan no cotiza índices).

### 5.4 TradingView widget: unique container ID por render
Cada llamada a `renderChart()` crea un `<div id="tv_TIMESTAMP">` nuevo y llama `new TradingView.widget(...)`. Esto fuerza un widget completamente fresco y evita que TradingView lea estado cacheado de localStorage.

### 5.5 SPX Weeklys → SPXW en leg identifiers
OptionStrat requiere `SPXW` (no `SPX`) para opciones semanales del SPX. El parser detecta el token `(Weeklys)` y setea `isWeekly: true`. `buildButterflyLegsString` y `buildCondorLegsString` usan `legTicker = SPXW` cuando `ticker === 'SPX' && isWeekly`.

### 5.6 isFutures regex: requiere whitespace prefix
```javascript
const isFutures = /(?:^|\s)\/[A-Z]/.test(input);
```
La versión sin `(?:^|\s)` daba falso positivo con `CALL/PUT` en Iron Condor.

### 5.8 Parsing de URL OptionStrat → TOS (Section 3e)

`parseOptionStratURL(url)` convierte directamente una URL de OptionStrat en:
- `tosFormatted`: string TOS con precio `@0.00 LMT` (placeholder)
- `parsed`: objeto compatible con `parseSingleLineStrategy` output

**Convenciones de BUY/SELL por slug** (contra-intuitivo en IRON_CONDOR):
- `iron-condor` / `iron-butterfly` → `BUY 1` aunque sea la estructura de crédito (long outer, short inner). Esto coincide con cómo `parseSingleLineStrategy` reconstruye los legs: `BUY IRON CONDOR` produce long outer / short inner = crédito.
- `bull-put-spread` → `BUY 1 VERTICAL` (primera pata long); `bear-call-spread` → `SELL -1 VERTICAL` (primera pata short).

**Orden de strikes en IRON_CONDOR**: `callSell/callBuy/putSell/putBuy` (reconstruido desde los legs de la URL, no desde el orden literal).

**BACKRATIO**: `BUY` → `parts[0]=short, parts[1]=long`, ratioStr=`|short|/|long|`; `SELL` → `parts[0]=long, parts[1]=short`.

**DIAGONAL/CALENDAR**: back month (long, fecha posterior) = parts[0], front month (short, fecha anterior) = parts[1]. Formato fecha: `"DD1 MON1 YY1/DD2 MON2 YY2"` (back first, para que el parser la lea correctamente).

**isWeekly**: calculado con `isThirdFriday(dateCode)` — true si la fecha NO es el 3er viernes del mes.

### 5.9 Formato del string para TOS Chart
```
+.SPY260420P711 -.SPY260420P712 -.SPY260420C713 +.SPY260420C714
```
- Signo `+`/`-` según qty
- Prefijo `2*` / `3*` si `|qty| > 1`
- Usa `parsed.ticker` (no `legTicker`) porque TOS acepta SPX directamente

---

## 6. Convenciones de código

- **Sin framework**: JS vanilla, sin jQuery, sin React.
- **Async solo cuando es necesario**: `processInput()` es síncrona hasta el clipboard; solo el fetch de Finnhub y la URL de OptionStrat son async.
- **`dateCode` siempre en formato `YYMMDD`**: ej. `260422` para 22 ABR 2026.
- **Posiciones**: cada pata es `{ strike: Number, optType: 'CALL'|'PUT', qty: Number, dateCode?: String }`. `qty > 0` = long, `qty < 0` = short.
- **`parseSingleLineStrategy` recibe texto ya formateado** (output de `formatAlert`), no el raw de TOS.
- **`currentParsed`** se pasa como `preParsed` a `buildPayoffModel` y `renderPayoffForOrder` para evitar re-parsear.
- **Funciones de URL sin precios** (condor, butterfly): usan `buildCondorLegsString` / `buildButterflyLegsString`. Funciones con precios usan `buildLegsString`.

---

## 7. Gráfico P&L al vencimiento (Section 3d)

### Funciones principales

| Función | Propósito |
|---|---|
| `buildPayoffModel(formattedOrder, preParsed?)` | Calcula curva P&L, breakevens, min/max |
| `inferEntryValuePerSpread(netPrice, allPositions)` | Detecta si la alerta es crédito o débito |
| `findBreakevens(points)` | Interpolación de cruces por cero (máx 4) |
| `drawPayoffChart(model)` | Render canvas con DPR-awareness |
| `renderPayoffForOrder(formattedOrder, preParsed?)` | Orquesta build + draw + fetch underlying |
| `hidePayoffChart()` | Oculta panel y limpia `currentPayoffModel` |

### `inferEntryValuePerSpread` — lógica de crédito vs débito

Cuando `netPrice > 0` (ej. `@.84` en `BOT`): prueba con precios extremos del subyacente. Si el gross máximo es ≤ 0 → estructura netamente corta → debit. Si el gross max > 0 → crédito. Permite detectar correctamente Iron Condors BOT como crédito.

### Race condition prevention
`payoffRenderToken` es un contador global. Cada render increments el token y verifica que sigue siendo el render activo antes de actualizar el modelo con el precio del subyacente.

---

## 8. Normalización de tickers especiales

| Input (raw TOS) | Output (formatAlert) |
|---|---|
| `/GCM26:XCEC` | `GCM2026` |
| `/OG2J26:XCEC` | `OG2J2026` |
| `BOT ... (cr)` | strip `cr` |
| `BOT ... (db)` | strip `db` |
| `#1462 NEW BOT ...` | strip `#1462 NEW ` |

Regex futuros: `processed.replace(/\/([A-Z][A-Z0-9]*)(\d{2})(?::[A-Z]{1,5})?/g, '$120$2')`

---

## 9. Variables globales de estado (Section 4)

```javascript
let currentResults  = { original:'', plus1:'', plus5:'' };
let lastSelectedOption = 'original';   // card seleccionada
let currentTicker   = '';              // ticker del último trade parseado
let currentParsed   = null;            // resultado de parseSingleLineStrategy (o parseOptionStratURL)
let currentInterval = 'D';            // timeframe del widget TV
let tvWidget        = null;            // instancia TradingView.widget
let currentPayoffModel = null;         // modelo P&L actual
let payoffRenderToken  = 0;            // anti-race condition para renderPayoffForOrder
let currentOptionStratURL = '';        // URL original si el input fue una URL OptionStrat; '' si fue TOS alert
```

---

## 10. Estado actual

### ✅ Implementado y funcionando

- Parseo de alertas TOS: SIMPLE, VERTICAL, BUTTERFLY, ~BUTTERFLY, BACKRATIO, CONDOR, IRON_CONDOR, DIAGONAL/CALENDAR
- Normalización de futuros (`/GCM26:XCEC` → `GCM2026`)
- 3 variantes de precio (+0, +1c, +5c) con delta badge dinámico
- Copia al clipboard (prioridad máxima, no bloqueada por otros renders)
- URLs OptionStrat con precios distribuidos por Black-Scholes (v2)
- URLs OptionStrat sin precios para CONDOR, IRON_CONDOR, BUTTERFLY
- BUTTERFLY: detección simétrico vs broken-wing, path correcto
- SPX Weeklys → SPXW en leg identifiers de OptionStrat
- Botón "Open in TradingView" (morado): abre chart full en nueva pestaña
  - Oculto para futuros e índices (sin widget embebido, pero botón sí aparece)
- Widget TradingView embebido (D/30m/15m/5m) con EMA8 + EMA21
- Botón "Copy Strategy for TOS Chart" (marrón): formato `+.SPY...` con multiplicadores
- Gráfico P&L al vencimiento (canvas 2D):
  - Stats: Net Credit/Debit, Max Loss, Max Profit, Breakevens
  - Zonas verde/roja con gradiente, líneas de breakeven y strikes
  - Marcador de precio subyacente (async Finnhub, no bloquea)
  - No muestra marcador para índices (SPX/NDX/RUT/DJX)
- Auto-select en textarea al hacer focus
- **Input dual**: acepta tanto alertas TOS como URLs de OptionStrat (`https://optionstrat.com/build/…`)
  - Estrategias soportadas vía URL: SIMPLE, VERTICAL, BUTTERFLY, ~BUTTERFLY, CONDOR, IRON_CONDOR, BACKRATIO, DIAGONAL/CALENDAR y sus variantes (bull/bear, broken-wing, ratio-spread, inverse, etc.)
  - TOS output con precio `@0.00` como placeholder; variantes +1c/+5c funcionan normalmente
  - Gráfico P&L muestra la estructura al precio cero; calendars/diagonals ocultan el chart (multi-date)
  - Botón "Open in OptionStrat" reabre la URL original directamente

### 🔲 Pendiente / posibles mejoras

- **NDX/RUT butterfly**: mismo problema que SPX+Weeklys puede existir si OptionStrat requiere tickers alternativos para estos índices — no verificado.
- **`buildTOSChartString` para índices**: usa `parsed.ticker` (SPX), no `SPXW`. En TOS esto probablemente está bien, pero no fue verificado con cuentas reales.
- **Marcador subyacente para índices**: actualmente se omite. Se podría agregar usando el proxy ETF * factor (el mismo que usa BS pricing) ya que `fetchStockPrice` sí lo resuelve vía `INDEX_TO_ETF`.
- **Estrategias no soportadas**: RATIO_SPREAD genérico, opciones sobre futuros con vencimiento propio.
- **Mobile UX**: funciona pero no está optimizado para pantallas < 400px.
- **Test unitario formal**: actualmente la verificación es manual + snippets Node ad-hoc.
- **URL con precio en legs**: `parseOptionStratURL` ignora precios incrustados en legs (formato `@N.NN`). Se podría usar para pre-llenar el precio en lugar de `@0.00`.

---

## 11. Problemas conocidos

| Problema | Impacto | Estado |
|---|---|---|
| Finnhub free plan: límite de rate (60 req/min) | Si se pegan muchas alertas seguidas, el fetch puede fallar silenciosamente y el gráfico P&L no muestra marcador de subyacente | Mitigado: `AbortSignal.timeout(3000)` + fallback silencioso |
| TradingView widget en Firefox/Safari privado: puede bloquearse por restricciones de 3rd-party scripts | Widget TV no carga | No se puede evitar; el botón "Open in TradingView" sigue funcionando |
| `parseSingleLineStrategy` asume que el texto ya fue normalizado por `formatAlert` | Si se llama con texto raw puede fallar con futuros o tickers con `:XCEC` | Por diseño: siempre llamar con output de `formatAlert` |
| Iron Condor isFutures false-positive (histórico, resuelto) | Estaba oculto el widget TV para Iron Condor | Resuelto con `(?:^|\s)\/[A-Z]` |

---

## 12. Dependencias externas

| Dependencia | Uso | Condición |
|---|---|---|
| `https://s3.tradingview.com/tv.js` | Widget TradingView embebido | Cargado en `<head>`, falla silencioso si offline |
| `https://finnhub.io/api/v1/quote` | Precio del subyacente para B-S + marcador P&L | API key hardcoded (plan free); falla silencioso con `AbortSignal.timeout(3000)` |
| GitHub Pages | Hosting | `master` branch → producción automática |

---

## 13. Deployment

```bash
# Hacer cambios en index.html (y opcionalmente index_v3.html)
git add index.html index_v3.html
git commit -m "Descripción del cambio"
git push origin master
# GitHub Pages se actualiza en ~1 minuto
```

No hay build step. No hay CI. El push a master es el deploy.
