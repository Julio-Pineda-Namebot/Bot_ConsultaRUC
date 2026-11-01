# Documentación técnica — Bot TA4 Consulta RUC SUNAT

**Proyecto:** TA4_Bot_ConsultaRUC (UiPath, tipo Proceso) · **Fecha:** 01/11/2026 · **Grupo:** 1

Este documento describe el bot **tal como quedó implementado**. Todo el contenido se obtuvo de los archivos reales del proyecto (`Main.xaml`, `project.json`, Object Repository en `.objects/`), del registro de Windows y del log de ejecución de UiPath. Las rutas que contienen el nombre de usuario de Windows se muestran como `<usuario>`.

**Objetivo del bot:** leer los RUC de la hoja `Consultas` del archivo `Data\TA4_Resultados_RUC_Grupo1.xlsx`, consultar cada uno en la Consulta RUC pública de SUNAT (`https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp`) y escribir en la misma fila la razón social, el domicilio fiscal, la condición del contribuyente y la fecha de inicio de actividades, junto con el estado de la consulta, una observación, la fecha y hora y el nombre de la captura de evidencia.

---

## 1. Versiones

| Componente | Versión | Fuente |
|---|---|---|
| UiPath Studio | 26.0.204.0 (compilación instalada `26.0.204-cloud.25540`, edición Community) | `project.json` → `studioVersion`; carpeta `C:\Program Files\UiPathPlatform\Studio\26.0.204-cloud.25540`; `robotVersion` en el log |
| UiPath.Excel.Activities | 3.6.1 | `project.json` → `dependencies` |
| UiPath.System.Activities | 26.10.0 | `project.json` → `dependencies` |
| UiPath.UIAutomation.Activities | 26.10.6 | `project.json` → `dependencies` |
| Google Chrome | 156.0.8078.76 | Registro: `HKCU\Software\Google\Chrome\BLBeacon\version` y `HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome\DisplayVersion` |
| Extensión UiPath para Chrome | UiPath Browser Automation 26.10 — v26.10.2 | `manifest.json` de la extensión en el perfil de Chrome |
| UiPath CLI (validación y compilación) | `uip` 1.204.0-preview.250 | `uip --version` |
| Sistema operativo | Windows 11 Pro 10.0.26300 | Entorno de la máquina |

Configuración del proyecto (`project.json`): `targetFramework` = **Windows**, `dotNetVersion` = **net10.0**, `expressionLanguage` = **VisualBasic**, `outputType` = **Process**, `main` = **Main.xaml**, `projectVersion` = 1.0.0, `schemaVersion` = 4.1.

> Nota: en el registro también existe una entrada residual `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome` con versión 138.0.7204.158. La versión en uso es la 156.0.8078.76 (la de `BLBeacon` y la instalación de sistema).

---

## 2. Estructura del proyecto

```text
TA4_Bot_ConsultaRUC/
├── Main.xaml                      Flujo principal del bot (único workflow, punto de entrada)
├── project.json                   Manifiesto: dependencias, versión de Studio, lenguaje VB, framework
├── project.uiproj                 Archivo de proyecto de Studio
├── AGENTS.md / CLAUDE.md          Instrucciones del proyecto para asistentes
├── .gitignore                     Excluye .local/ y .storage/ (cachés generadas por Studio)
├── Data/
│   └── TA4_Resultados_RUC_Grupo1.xlsx   Entrada y salida del bot (hojas Consultas, Resumen, Instrucciones)
├── Evidencias/
│   ├── DOCUMENTACION_TECNICA.md   Este documento
│   ├── Capturas/                  20 capturas de la última ejecución (<RUC>.png)
│   └── Logs/
│       ├── 2026-11-01_Execution.log         Copia del log de UiPath del día (datos personales reemplazados)
│       └── Ejecucion_final_2026-11-01.log   Solo las líneas de la última ejecución completa
├── .objects/                      Object Repository (aplicación, pantallas y elementos de UI)
├── .screenshots/                  Imágenes de diseño de los elementos del Object Repository (15)
├── .settings/                     Configuración del proyecto (Debug / Design / Release, 16 archivos)
├── .project/                      Metadatos de Studio (design.json, PackageBindingsMetadata.json)
├── .tmh/, .entities/, .templates/ Carpetas estándar de Studio
├── .local/                        (generada, no se versiona)
└── .storage/                      (caché de diseño de Studio, no se versiona)
```

**Object Repository (`.objects/`):** una aplicación, *Usar navegador Chrome en Consulta RUC SUNAT*, con dos pantallas:

| Pantalla | Selector de ventana | Elementos |
|---|---|---|
| *Usar navegador Chrome en Consulta RUC SUNAT* (creada por Studio) | `<html app='chrome.exe' title='SUNAT - Consulta Ruc' />` | 5 elementos antiguos creados al indicar en Studio; **ya no se usan en Main.xaml** |
| *Ficha De Resultados RUC* | `<html app='chrome.exe' title='*Consulta RUC*' url='https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc*' />` | Encabezado Resultado De La Busqueda, Valor Numero De RUC, Valor Domicilio Fiscal, Valor Condicion Del Contribuyente, Valor Fecha Inicio Actividades, Mensaje RUC No Valido (**en uso**) |

**Hoja `Consultas` (fila 1, encabezados):** N.° · RUC · Razón social · Domicilio fiscal · Condición · Fecha inicio actividades · Estado consulta · Observación · Fecha y hora consulta · Captura · Formato RUC válido. El bot lee y escribe por nombre de columna; no toca la columna `Formato RUC válido` (fórmula).

---

## 3. Flujo de Main.xaml

Orden de ejecución con su anidamiento. Formato: **Tipo de actividad** — *DisplayName* (`IdRef`): qué hace.

- **Sequence** — *Bot consulta RUC SUNAT* (`Sequence_1`): secuencia principal; declara todas las variables.
  1. **Log Message** — *Registrar inicio del bot* (`LogMessage_1`): nivel Info, mensaje `"Inicio del bot de consulta RUC"`.
  2. **Use Excel File** (`ExcelApplicationCard`) — *Usar archivo Excel de resultados* (`ExcelApplicationCard_1`): abre `strRutaExcel`. *Crear si no existe* = No; *Guardar cambios* = Sí. Referencia del libro: `Excel`.
     - **Use Application/Browser** (`NApplicationCard`) — *Usar navegador Chrome en Consulta RUC SUNAT* (`NApplicationCard_1`): abre o reutiliza Chrome en `strUrlSunat`. Modo de adjuntar: `ByInstance`.
       - **For Each Excel Row** (`ExcelForEachRowX`) — *Recorrer filas de la hoja Consultas* (`ExcelForEachRowX_1`): recorre `Excel.Sheet("Consultas")` con encabezados; guarda el libro después de cada fila (`SaveAfterEachRow = True`); se detiene en la primera fila vacía. Variables de iteración: `CurrentRow`, `CurrentIndex`.
         - **Sequence** — *Procesar fila del RUC* (`Sequence_4`):
           1. **Assign** — *Leer RUC de la fila* (`Assign_1`): `strRuc = If(Convert.ToString(CObj(CurrentRow("RUC"))), "").Trim`.
           2. **If** — *¿La fila ya tiene Estado consulta OK?* (`If_1`): si `Estado consulta` (sin espacios, en mayúsculas) es `OK` → **Continue** *Saltar fila ya consultada* (`Continue_1`).
           3. **If** — *¿Celda RUC vacía?* (`If_2`): si `strRuc` está vacío → **Continue** *Saltar fila sin RUC* (`Continue_2`).
           4. **If** — *¿RUC con formato no válido?* (`If_3`): si `strRuc` no cumple la expresión regular `^(10|15|16|17|20)\d{9}$` → secuencia *Registrar RUC inválido* (`Sequence_7`):
              - **Write Cell** ×3 (`WriteCellX_1` a `WriteCellX_3`): Estado consulta = `"RUC INVÁLIDO"`, Observación = `"Formato no válido, no se consultó"`, Fecha y hora consulta = `Now.ToString("dd/MM/yyyy HH:mm:ss")`.
              - **Assign** — *Sumar consulta fallida (RUC inválido)* (`Assign_2`): `intFallidas + 1`.
              - **Log Message** — *Registrar RUC inválido en el log* (`LogMessage_2`): nivel Warn.
              - **Continue** — *Pasar a la siguiente fila* (`Continue_3`).
           5. **Try Catch** — *Consultar RUC en SUNAT (Try Catch)* (`TryCatch_1`):
              - **Try** (`Sequence_8`):
                1. **Multiple Assign** — *Limpiar datos de la consulta anterior* (`MultipleAssign_1`): vacía `strTextoRuc`, `strRazonSocial`, `strDomicilio`, `strCondicion`, `strFechaInicio`, `strEstado`, `strObservacion`, `strCaptura`, `strMensajeSunat` y pone `blnFichaEncontrada = False`.
                2. **Retry Scope** — *Reintentar búsqueda del RUC en SUNAT* (`RetryScope_1`): número de intentos 2, intervalo 5 s; reintenta solo si una actividad lanza excepción. Contiene la secuencia *Buscar RUC en SUNAT* (`Sequence_9`):
                   1. **Go To URL** — *Ir a la página de Consulta RUC* (`NGoToUrl_1`): navega a `strUrlSunat`.
                   2. **Check App State** — *Verificar que cargó el formulario de búsqueda* (`NCheckState_1`): espera hasta 10 s el botón «Por RUC». Si no aparece → **Message Box** *Pedir a una persona que resuelva el captcha* (`MessageBox_1`): `"Resuelva el captcha y presione OK"`.
                   3. **Click** — *Clic en búsqueda Por RUC* (`NClick_1`).
                   4. **Type Into** — *Escribir número de RUC* (`NTypeInto_1`): escribe `strRuc`; vacía el campo antes (`EmptyFieldMode = SingleLine`).
                   5. **Check App State** — *¿SUNAT muestra captcha de imagen?* (`NCheckState_2`): espera 2 s a que el campo del código captcha sea **visible**. Si aparece → **Message Box** *Pedir a una persona que resuelva el captcha de imagen* (`MessageBox_2`).
                   6. **Browser Dialog Scope** — *Buscar y capturar alerta de SUNAT* (`NBrowserDialogScope_1`): contiene **Click** *Clic en Buscar* (`NClick_2`). Si SUNAT responde con una alerta del navegador, la acepta (`OK`) y guarda su texto en `strMensajeSunat`. Espera la alerta 3 s; `ContinueOnError = True`.
                   7. **Check App State** — *Esperar ficha de resultados* (`NCheckState_3`): espera hasta 15 s el encabezado «Resultado de la Búsqueda»; guarda el resultado en `blnFichaEncontrada`.
                3. **Check App State** — *¿La página muestra la ficha del RUC?* (`NCheckState_4`): espera hasta 5 s el valor junto a la etiqueta «Número de RUC:» (solo existe en una ficha real) y sobrescribe `blnFichaEncontrada`.
                4. **If** — *¿Apareció la ficha del RUC?* (`If_4`), condición `blnFichaEncontrada`:
                   - **Then** — *Extraer y registrar datos de la ficha* (`Sequence_16`):
                     1. **Get Text** — *Obtener texto de Número de RUC* (`NGetText_1`) → `strTextoRuc` (p. ej. `20100055237 - ALICORP SAA`).
                     2. **Get Text** — *Obtener texto de Domicilio Fiscal* (`NGetText_2`) → `strDomicilio`.
                     3. **Get Text** — *Obtener texto de Condición del Contribuyente* (`NGetText_3`) → `strCondicion`.
                     4. **Get Text** — *Obtener texto de Fecha de Inicio de Actividades* (`NGetText_4`) → `strFechaInicio`.
                     5. **Multiple Assign** — *Separar razón social y definir estado y observación* (`MultipleAssign_2`):
                        - `strRazonSocial = strTextoRuc.Split({" - "}, 2, StringSplitOptions.None)(1).Trim`
                        - `strDomicilio = Regex.Replace(strDomicilio, "\s+", " ").Trim` (reduce los espacios internos que trae SUNAT)
                        - `strCondicion` = solo la primera línea del texto (SUNAT agrega una nota debajo de «NO HABIDO»)
                        - `strFechaInicio = strFechaInicio.Trim`
                        - `strEstado = "OK"`
                        - `strObservacion = "Campo sin dato en SUNAT"` si algún campo es `"-"`; si no, `"Sin observaciones"`
                        - `strCaptura = strRuc + ".png"`
                     6. **Take Screenshot** — *Tomar captura de la ficha del RUC* (`NTakeScreenshot_1`): guarda `Evidencias\Capturas\<RUC>.png`.
                     7. **Write Cell** ×8 (`WriteCellX_4` a `WriteCellX_11`): Razón social, Domicilio fiscal, Condición, Fecha inicio actividades, Estado consulta (`OK`), Observación, Fecha y hora consulta (`dd/MM/yyyy HH:mm:ss`) y Captura.
                     8. **Assign** — *Sumar consulta OK* (`Assign_3`): `intOk + 1`.
                     9. **Log Message** — *Registrar RUC y condición en el log* (`LogMessage_3`): Info, `"RUC <ruc> consultado OK. Condición: <condición>"`.
                   - **Else** — *Registrar RUC no encontrado* (`Sequence_17`):
                     1. **If** — *¿SUNAT no mostró alerta del navegador?* (`If_5`): si `strMensajeSunat` está vacío → **Get Text** *Obtener aviso de SUNAT en la página* (`NGetText_6`) → `strMensajeSunat` (timeout 5 s, `ContinueOnError = True`).
                     2. **Multiple Assign** — *Definir estado NO ENCONTRADO* (`MultipleAssign_3`): `strEstado = "NO ENCONTRADO"`; `strObservacion` = mensaje de SUNAT con espacios normalizados o, si no hay mensaje, `"SUNAT no mostró la ficha del RUC (ver captura)"`; `strCaptura = strRuc + ".png"`.
                     3. **Take Screenshot** — *Tomar captura del mensaje de SUNAT* (`NTakeScreenshot_2`).
                     4. **Write Cell** ×4 (`WriteCellX_12` a `WriteCellX_15`): Estado consulta, Observación, Fecha y hora consulta, Captura.
                     5. **Assign** — *Sumar consulta fallida (no encontrado)* (`Assign_4`): `intFallidas + 1`.
                     6. **Log Message** — *Registrar RUC no encontrado en el log* (`LogMessage_4`): Warn.
              - **Catch** `System.Exception` (variable `exception`) — *Registrar error de la consulta* (`Sequence_18`):
                1. **Multiple Assign** — *Definir estado ERROR* (`MultipleAssign_4`): `strEstado = "ERROR"`, `strObservacion = exception.Message`, `strCaptura = strRuc + "_error.png"`.
                2. **Take Screenshot** — *Tomar captura del error* (`NTakeScreenshot_3`): `ContinueOnError = True`.
                3. **Write Cell** ×4 (`WriteCellX_16` a `WriteCellX_19`): Estado consulta, Observación, Fecha y hora consulta y Captura (solo el nombre si el archivo existe).
                4. **Log Message** — *Registrar error en el log* (`LogMessage_5`): Error, `"Error al consultar el RUC <ruc>: <mensaje>"`.
                5. **Assign** — *Sumar consulta fallida (error)* (`Assign_5`): `intFallidas + 1`.
           6. **Delay** — *Pausa de 2 a 3 s entre consultas* (`Delay_1`): `TimeSpan.FromMilliseconds(New Random().Next(2000, 3001))`.
  3. **Log Message** — *Registrar resumen final* (`LogMessage_6`): Info, `"Fin. OK: " + intOk.ToString + " | Fallidas: " + intFallidas.ToString`. Se ejecuta después de cerrar el Excel.

---

## 4. Variables

Todas están declaradas en la secuencia principal *Bot consulta RUC SUNAT*.

| Nombre | Tipo | Valor inicial |
|---|---|---|
| strRutaExcel | String | `Data\TA4_Resultados_RUC_Grupo1.xlsx` |
| strUrlSunat | String | `https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp` |
| strRuc | String | (sin valor) |
| strTextoRuc | String | (sin valor) |
| strRazonSocial | String | (sin valor) |
| strDomicilio | String | (sin valor) |
| strCondicion | String | (sin valor) |
| strFechaInicio | String | (sin valor) |
| strEstado | String | (sin valor) |
| strObservacion | String | (sin valor) |
| strCaptura | String | (sin valor) |
| intOk | Int32 | 0 |
| intFallidas | Int32 | 0 |
| blnFichaEncontrada | Boolean | False |
| strMensajeSunat | String | (sin valor) |

Argumentos de ámbito (no son variables declaradas): `Excel` (libro abierto por *Use Excel File*), `WSSessionData` (cuerpo de *Use Browser*), `CurrentRow` y `CurrentIndex` (*For Each Excel Row*) y `exception` (*Catch*).

---

## 5. Targets de las actividades de UI

**Ventanas:**
- Página de búsqueda: `<html app='chrome.exe' title='SUNAT - Consulta Ruc' />`
- Página de resultados (ficha o aviso): `<html app='chrome.exe' title='*Consulta RUC*' url='https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc*' />`

| Actividad (`IdRef`) | Método | Selector final | Ancla | Origen del target / ¿re-indicado en Studio? |
|---|---|---|---|---|
| Usar navegador Chrome en Consulta RUC SUNAT (`NApplicationCard_1`) | Selector de aplicación + URL `strUrlSunat`, navegador Chrome | `<html app='chrome.exe' title='SUNAT - Consulta Ruc' />` | — | Título verificado con el HTML real de la página. Studio lo registró en el Object Repository (pantalla creada por Studio); el selector no cambió. |
| Verificar que cargó el formulario de búsqueda (`NCheckState_1`) | Selector estricto | `<webctrl id='btnPorRuc' tag='BUTTON' />` | — | `id` tomado del HTML real de la página de búsqueda. No se re-indicó. |
| Clic en búsqueda Por RUC (`NClick_1`) | Selector estricto | `<webctrl id='btnPorRuc' tag='BUTTON' />` | — | `id` tomado del HTML real. No se re-indicó. |
| Escribir número de RUC (`NTypeInto_1`) | Selector estricto | `<webctrl id='txtRuc' tag='INPUT' />` | — | `id` tomado del HTML real. No se re-indicó. |
| ¿SUNAT muestra captcha de imagen? (`NCheckState_2`) | Selector estricto + visibilidad | `<webctrl id='txtCodigo' tag='INPUT' />` | — | `id` tomado del HTML real (el campo existe pero está oculto). No se re-indicó. |
| Clic en Buscar (`NClick_2`) | Selector estricto | `<webctrl id='btnAceptar' tag='BUTTON' />` | — | `id` tomado del HTML real. No se re-indicó. |
| Esperar ficha de resultados (`NCheckState_3`) | Selector estricto | `<webctrl aaname='Resultado de la Búsqueda' class='*panel-heading*' parentclass='*panel-primary*' tag='DIV' />` | — | Capturado con la CLI de UiPath (`uia-configure-target`) sobre la ficha real; elemento *Encabezado Resultado De La Busqueda* del Object Repository. |
| ¿La página muestra la ficha del RUC? (`NCheckState_4`) | Selector difuso + ancla | `<webctrl tag='H4' />` | `<webctrl aaname='Número de RUC:' tag='H4' class='list-group-item-heading' … check:text='Número de RUC:' />` | Reutiliza el elemento *Valor Numero De RUC*. Probado en vivo: **no** se encuentra en la página de aviso «RUC no válido». |
| Obtener texto de Número de RUC (`NGetText_1`) | Selector difuso + ancla | `<webctrl tag='H4' />` | Etiqueta «Número de RUC:» (H4) | Studio lo indicó primero con una posición (`idx='2'`); se reemplazó con la CLI de UiPath por valor + ancla. Probado: `20100055237 - ALICORP SAA`. |
| Obtener texto de Domicilio Fiscal (`NGetText_2`) | Selector difuso + ancla | `<webctrl tag='P' />` | Etiqueta «Domicilio Fiscal:» (H4) | Studio lo indicó primero con `idx='8'`; reemplazado con la CLI por valor + ancla. Probado con la ficha de 20100055237. |
| Obtener texto de Condición del Contribuyente (`NGetText_3`) | Selector difuso + ancla | `<webctrl tag='P' />` | Etiqueta «Condición del Contribuyente:» (H4) | Studio lo indicó primero con `idx='7'`; reemplazado con la CLI. Probado: `HABIDO` (no el campo «Estado»). |
| Obtener texto de Fecha de Inicio de Actividades (`NGetText_4`) | Selector difuso + ancla | `<webctrl tag='P' />` | Etiqueta «Fecha de Inicio de Actividades:» (H4) | Studio lo indicó primero con `idx='5'`; reemplazado con la CLI. Probado: `16/07/1956` (no la Fecha de Inscripción). |
| Obtener aviso de SUNAT en la página (`NGetText_6`) | Selector estricto | `<webctrl parentclass='*text-center*' tag='STRONG' />` | — | Capturado con la CLI sobre la página real del aviso (RUC 20100128129); elemento *Mensaje RUC No Valido*. |
| Ir a la página de Consulta RUC (`NGoToUrl_1`) | Sin target (usa la pestaña de la tarjeta) | — | — | — |
| Buscar y capturar alerta de SUNAT (`NBrowserDialogScope_1`) | Sin target (captura alertas JavaScript) | — | — | — |
| Tomar captura … (`NTakeScreenshot_1` a `_3`) | Sin target (ventana de la tarjeta) | — | — | — |

Las anclas de los cuatro *Get Text* usan selector difuso sobre la etiqueta: `aaname` y `class='list-group-item-heading'` con coincidencia difusa y `check:text` igual al texto de la etiqueta. Por qué se reemplazaron las posiciones: la ficha de una empresa (RUC 20) no tiene la fila «Tipo de Documento» que sí tiene una persona natural (RUC 10), así que las posiciones cambian. Por ejemplo, la condición es el 7.º párrafo en la ficha de RUC 10 y el 6.º en la de RUC 20.

Ninguna actividad usa selectores semánticos, CV ni imagen.

---

## 6. Manejo de errores

### Estados posibles en la columna «Estado consulta»

| Estado | Cuándo | Observación escrita | Captura | Suma en |
|---|---|---|---|---|
| `OK` | Apareció la ficha y se extrajeron los 4 datos | `Sin observaciones` o `Campo sin dato en SUNAT` (si algún campo es `-`) | `<RUC>.png` | `intOk` |
| `NO ENCONTRADO` | SUNAT mostró una alerta del navegador (p. ej. «Por favor, ingrese número de RUC válido.») o la página de aviso («El número de RUC … consultado no es válido. Debe verificar el número y volver a ingresar.») | Texto exacto del mensaje de SUNAT; si no se pudo leer: `SUNAT no mostró la ficha del RUC (ver captura)` | `<RUC>.png` | `intFallidas` |
| `RUC INVÁLIDO` | El RUC no cumple el formato `^(10\|15\|16\|17\|20)\d{9}$`; no se consulta | `Formato no válido, no se consultó` | — | `intFallidas` |
| `ERROR` | Cualquier excepción dentro del Try (después de los reintentos) | `exception.Message` | `<RUC>_error.png` (si se pudo tomar) | `intFallidas` |

No se procesan (se saltan con *Continue*) las filas con `Estado consulta = OK` ni las filas sin RUC.

### Mecanismos
- **Reintentos:** el *Retry Scope* repite la búsqueda completa (ir a la URL, Por RUC, escribir, Buscar, esperar resultados) si alguna actividad lanza excepción. En UiPath, *Number of retries* = 2 significa **2 intentos en total** (1 intento + 1 reintento), con 5 s entre ellos. No reintenta por condición.
- **Try Catch por fila:** cada RUC está aislado. Si falla, la fila queda en `ERROR` con captura y el bot continúa con el siguiente RUC.
- **Continue on error:** `Buscar y capturar alerta de SUNAT` (si no aparece alerta, el flujo sigue), `Obtener aviso de SUNAT en la página` (si no hay aviso, se usa el texto por defecto) y `Tomar captura del error` (un fallo de la captura no impide registrar el error).
- **Intervención humana (captcha):** si el formulario de búsqueda no carga en 10 s (p. ej. una verificación humana de SUNAT o Cloudflare) o si el campo de captcha de imagen se vuelve visible, un *Message Box* pide «Resuelva el captcha y presione OK». El bot nunca resuelve ni evade captchas.
- **Pausa entre consultas:** 2 a 3 s aleatorios después de cada RUC consultado (OK, NO ENCONTRADO o ERROR) para no saturar el servicio público.

### Re-ejecución
- El libro se guarda después de cada fila (`SaveAfterEachRow = True`), así que una ejecución interrumpida conserva lo ya procesado.
- Al volver a ejecutar, solo se procesan las filas cuyo `Estado consulta` no es `OK`: vacías, `ERROR`, `NO ENCONTRADO` y `RUC INVÁLIDO`.
- Para repetir una fila que ya está en `OK`, basta con borrar su celda «Estado consulta».

---

## 7. Diferencias respecto a la especificación original

| Especificación original | Implementación | Motivo |
|---|---|---|
| `strRutaExcel = "Data\TA4_Resultados_RUC_GrupoX.xlsx"` | `Data\TA4_Resultados_RUC_Grupo1.xlsx` | Pedido del usuario: el archivo se renombró con el número de grupo. |
| `strRuc = CurrentRow("RUC").ToString.Trim` | `If(Convert.ToString(CObj(CurrentRow("RUC"))), "").Trim` (lo mismo para «Estado consulta», comparado en mayúsculas) | Una celda vacía llega como `Nothing`. En la primera prueba (13:23) el bot se detuvo con *Object reference not set to an instance of an object*. |
| Variables listadas en la especificación | Se agregaron `blnFichaEncontrada` (Boolean) y `strMensajeSunat` (String) | Guardar el resultado de *Check App State* y el texto de la alerta o aviso de SUNAT. |
| Retry Scope «2 reintentos, 5 s» | `NumberOfRetries = 2`, `RetryInterval = 00:00:05` | En UiPath el valor cuenta intentos totales: equivale a 1 intento + 1 reintento. |
| Captcha: Message Box si aparece | Dos verificaciones: formulario no cargado en 10 s y campo `txtCodigo` visible | Al descargar la página, SUNAT respondió con una verificación de Cloudflare; el campo de captcha de imagen existe oculto en el HTML. |
| Clic en «Buscar» | Dentro de *Browser Dialog Scope* (acepta la alerta y guarda su texto) | SUNAT valida el dígito verificador y muestra una alerta del navegador («Por favor, ingrese número de RUC válido.») que bloquea la página. |
| Check App State de la ficha + Get Text del mensaje de SUNAT | *Check App State* del encabezado «Resultado de la Búsqueda» + segundo *Check App State* del valor junto a «Número de RUC:»; el mensaje se toma de la alerta o del aviso en la página | La página de aviso «RUC no válido» también tiene el encabezado; sin la segunda verificación, 2 RUC terminaron en `ERROR` en la ejecución de las 14:32. |
| 4 Get Text con la etiqueta como ancla | Selector difuso del valor (`P` o `H4`) + ancla en la etiqueta, configurados con la CLI de UiPath sobre la ficha real | Al indicar en Studio se guardaron posiciones (`idx`) que cambian entre fichas de RUC 10 y RUC 20, y el analizador de UiPath las marcaba como error (UI-REL-001). |
| Selectores: ids del HTML real y «TODO: re-indicar» en lo no verificable | Página de búsqueda: ids del HTML. Página de resultados: capturada en vivo con la CLI de UiPath sobre páginas abiertas por el usuario. No quedan targets «TODO». | La ficha y los avisos solo se ven después de una consulta, que hizo el usuario a mano. |
| Escribir los datos tal cual | Domicilio con espacios internos normalizados; Condición solo con la primera línea | SUNAT entrega el domicilio con muchos espacios internos y agrega una nota de varias líneas debajo de «NO HABIDO». |
| NO ENCONTRADO: captura y siguiente fila | También suma en `intFallidas` y espera 2–3 s | Para que OK + Fallidas = RUC procesados, como en la hoja Resumen («Total procesados (OK + fallidas)»). |
| Delay de 2–3 s después de un OK | Delay después del Try Catch (OK, NO ENCONTRADO y ERROR) | No saturar el servicio en ningún caso. Las filas saltadas no esperan. |
| Columna Captura | Solo el nombre del archivo (`<RUC>.png`); en ERROR, solo si la captura existe | Igual al ejemplo de la hoja Instrucciones (`20XXXXXXXXX.png`); el archivo está en `Evidencias\Capturas\`. |
| Use Excel File «con guardado de cambios» | *Guardar cambios* = Sí, *Crear si no existe* = No, *Guardar después de cada fila* = Sí | Si falta el archivo, falla con un error claro; conservar el avance si se interrumpe. |
| Log Message final después del bucle | Se ejecuta después de cerrar *Use Excel File* | Sigue estando después del bucle, con el libro ya guardado. |
| — | Se reinician las variables al inicio de cada consulta | Evitar que datos de un RUC pasen al siguiente. |
| — | No se usa *Ámbito de proceso de Excel* | No estaba en la especificación; el validador solo lo muestra como advertencia. |

Validación del proyecto: `uip rpa validate` sin errores (1 advertencia: recomienda *Ámbito de proceso de Excel*) y `uip rpa build` exitoso (advertencias: verificación de ejecución no activada en Click/Type Into, URL de Automation Hub y uso de Delay).

---

## 8. Última ejecución completa

Fuente: `C:\Users\<usuario>\AppData\Local\UiPath\Logs\2026-11-01_Execution.log`, copiado a `Evidencias/Logs/2026-11-01_Execution.log`. En la copia se reemplazaron el usuario de Windows (`<usuario>`), el correo del robot (`<correo>`) y la clave `userKey` (`<redactado>`). El extracto de esta ejecución está en `Evidencias/Logs/Ejecucion_final_2026-11-01.log`.

| Dato | Valor |
|---|---|
| Inicio | 01/11/2026 14:52:40 (UTC−05:00) |
| Fin | 01/11/2026 14:58:07 (UTC−05:00) |
| Duración total | 5 min 27,5 s (≈ 16,4 s por RUC) |
| RUC procesados | 20 |
| OK | 19 |
| Fallidas | 1 (NO ENCONTRADO) |
| Errores (estado ERROR / nivel Error en el log) | Ninguno |
| Mensaje final del bot | `Fin. OK: 19 \| Fallidas: 1` |

**Fallida:** RUC 20100128129 → NO ENCONTRADO: «El número de RUC 20100128129 consultado no es válido. Debe verificar el número y volver a ingresar.» (mensaje mostrado por SUNAT).

Detalle por RUC según el log:

| Hora | RUC | Resultado | Condición |
|---|---|---|---|
| 14:52:57 | 10756668175 | OK | HABIDO |
| 14:53:14 | 20100055237 | OK | HABIDO |
| 14:53:29 | 20100053455 | OK | HABIDO |
| 14:53:45 | 20100070970 | OK | HABIDO |
| 14:54:02 | 20429683581 | OK | HABIDO |
| 14:54:18 | 20574650187 | OK | HABIDO |
| 14:54:34 | 20514680401 | OK | HABIDO |
| 14:54:50 | 20100105862 | OK | HABIDO |
| 14:55:05 | 20601013011 | OK | HABIDO |
| 14:55:21 | 20100154057 | OK | HABIDO |
| 14:55:37 | 20601233488 | OK | HABIDO |
| 14:55:53 | 20100047218 | OK | HABIDO |
| 14:56:09 | 20100130204 | OK | HABIDO |
| 14:56:25 | 20518779690 | OK | NO HABIDO |
| 14:56:41 | 20100128056 | OK | HABIDO |
| 14:57:00 | 20100128129 | NO ENCONTRADO | — |
| 14:57:16 | 20106897914 | OK | HABIDO |
| 14:57:32 | 20389230724 | OK | HABIDO |
| 14:57:48 | 20100043140 | OK | HABIDO |
| 14:58:04 | 20508565934 | OK | HABIDO |

**Ejecuciones previas del mismo día** (pruebas durante el desarrollo, en el mismo log):

| Inicio | Fin | Resultado |
|---|---|---|
| 13:23:53 | 13:23:57 | Se detuvo en la primera fila: *Object reference not set to an instance of an object* (celda vacía; corregido). |
| 14:30:11 | 14:30:42 | Solo procesó la primera fila (NO ENCONTRADO: «SUNAT no mostró la ficha del RUC (ver captura)») y terminó sin registrar el resumen. |
| 14:32:14 | 14:38:45 | `Fin. OK: 17 \| Fallidas: 3`. Dos ERROR por la página de aviso «RUC no válido» (corregido con `NCheckState_4` y `NGetText_6`) y un NO ENCONTRADO por alerta del navegador. |
| 14:49:22 | 14:50:19 | `Fin. OK: 2 \| Fallidas: 1`. Re-ejecución que procesó solo las 3 filas que no estaban en OK: 20100047218 y 20100128056 (OK) y 20100128129 (NO ENCONTRADO por el aviso «RUC no válido», ya manejado). |
| 14:52:40 | 14:58:07 | **Ejecución completa final** (detallada arriba). |
