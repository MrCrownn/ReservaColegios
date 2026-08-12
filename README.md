# Sistema de Reservas 

Sistema de reservas de **Tablets** y **Sala** sobre Google Sheets + Apps Script, con frontend embebido en Google Sites.

## Estructura

| Archivo | Descripción |
|---|---|
| `Tablets.gs` | Backend de Tablets en **un solo archivo**: `doGet` (CSV del horario) + `doPost` (registrar reserva + notificación al admin) + `enviarCorreoAlEditar` (trigger `onEdit`, envío de correos y escritura en el horario). |
| `Sala.gs` | Backend de Reserva de Sala: `doGet` (CSV del horario) + `doPost` (registrar reserva + notificación al admin) + `onEditSala` (trigger `onEdit`, envío de correos y escritura en el horario). |
| `index.html` | Frontend (Tablets y Sala en una sola página). Se embebe tal cual en Google Sites. |
| `SelectorUI.html` | Diálogo de selección de profesor cuando hay múltiples coincidencias de correo (usado por `Tablets.gs`). |
| `Plantilla Tablets.xlsx` | Plantilla del spreadsheet de Tablets (estructura de hojas sin datos reales). |
| `Plantilla Sala.xlsx` | Plantilla del spreadsheet de Sala (estructura de hojas sin datos reales). |

## Configuración necesaria

Antes de desplegar, reemplaza los valores sensibles (marcados con `TU_...`):

1. **`Tablets.gs` y `Sala.gs`**
   ```javascript
   const CONFIG_EMAIL_ADMIN = "TU_CORREO_ADMIN@COLEGIO.CL";
   ```
   Correo que recibe el aviso de cada nueva reserva.

2. **`index.html`**
   ```javascript
   const CSV_URL = "https://TU_URL_WEBAPP_TABLETS/exec";
   const CSV_URL_CRA = "https://TU_URL_WEBAPP_SALA/exec";
   ```
   URL de la implementación (Web App) de Apps Script de Tablets y de Sala.

## Estructura de hojas (Google Sheets)

### Tablets
- **Solicitudes** — pedidos registrados: `Profesor  (Columna A)`, `Cantidad  (Columna B) `, `Fecha  (Columna C)`, `Curso  (Columna D)`, `HoraInicio  (Columna E)`, `HoraFin  (Columna F)`, `Estado  (Columna G)`, `Comentarios (Columna I)`, `Checkbox (Columna J)`, `Errores (Columna K)`.
- **Profes** — diccionario de profesores: `Nombre`, `Apellido`, `Correo`.
- **Horario semanal** — columna `Hora` + columnas por día. Celdas coloreadas = reserva.

### Sala
- **Solicitudes** — pedidos registrados: `Profesor`, `Curso`, `Fecha`, `HoraInicio`, `HoraFin`, `ReservAceptada (F)`, `Observaciones (G)`.
- **Profes** — diccionario de profesores: `Nombre`, `Apellido`, `Correo`.
- **Horario** — columna `Hora` + columnas por día. Celdas con fondo verde = reserva.

## Despliegue

1. Sube cada backend (`.gs`) + `SelectorUI.html` al proyecto Apps Script de su hoja de cálculo.
2. Crea el trigger instalable (en cada proyecto):
   - Tablets: **`enviarCorreoAlEditar` → al editar (onEdit) → Head**.
   - Sala: **`onEditSala` → al editar (onEdit) → Head**.
3. Implementa como **Web App** (`Ejecutar como: Yo`, `Acceso: Cualquier persona`).
4. Copia cada URL de Web App a `index.html` (`CSV_URL` / `CSV_URL_CRA`).
5. Embebe `index.html` en Google Sites (Insertar → HTML o página embebida).

## Plantillas

Las plantillas muestran la estructura exacta de cada spreadsheet (nombres de hojas, columnas, fórmulas, colores y validaciones) sin datos reales. Úsalas como referencia para replicar la estructura en Google Sheets.
## Demo

Para ver el funcionamiento de este proyecto se puede visitar la url.

   https://sites.google.com/view/prueba-pag-gestor/web-prueba
   
   Además, si se quiere ver los google sheets utilizados.
   
   https://docs.google.com/spreadsheets/d/1Q-zxLnxYTneZm-qBFq1PcGDXVN8lHe7sAi3kssPDzkg/edit?usp=sharing
   
   https://docs.google.com/spreadsheets/d/10Ca80sLRdRskZLzuZTHC3hmZNB9tG_4Ak_u1QuBH2Fo/edit?usp=sharing
   
   

