# BioGrow Analytics — Aplicación de Escritorio

Monitor de Germinación y Biometría de Semillas.
Aplicación de escritorio para registrar experimentos de germinación (lotes de hasta 27 semillas) y consultar su historial.

Desarrollada en **Python + CustomTkinter/Tkinter + MySQL**.

---

## Funcionalidades

- **Datos del experimento (lote):** identificador del lote, variedad de semilla, fecha de registro, tiempo de tratamiento, temperatura, **tipo de gas** (Sin gas, Aire, Oxígeno, Helio, Argón, Nitrógeno), **flujo (L/min)**, **orientación** de la semilla (Aleatorio, Embrión, Endospermo) y **morfología** (Plana, Bola).
- **Variedades disponibles:** `Semilla de maíz blanco`, `Semilla de maíz rosa`, `Semilla de maíz cacahuazintle`, `Semilla de maíz criollo`, `Semilla de lenteja`.
- **Tabla de captura** de 27 muestras por defecto; la **cantidad de semillas por lote es configurable** (1–300, botón «Aplicar»):

  | No. | Longitud Raíz (mm) | Longitud Tallo (mm) | Cantidad de hojas | Estado de Germinación |
  |---|---|---|---|---|

- **Estado de germinación automático** por cada semilla.
- **Historial consolidado** (una fila por semilla medida) con búsqueda y filtros.
- **Navegación por teclado** tipo hoja de cálculo.

---

## Estructura de `BioGrow_Desktop/`

```
BioGrow_Desktop/
├── main.py                  → Punto de entrada; muestra un mensaje si MySQL no está disponible
├── interfaz.py              → Interfaz gráfica (CustomTkinter + ttk.Treeview)
├── database.py              → Acceso a MySQL (DatabaseManager)
├── models.py                → Modelos de datos (dataclasses)
├── utils.py                 → Validaciones, regla de germinación y formateo
├── motor_biogrow.py         → Cálculos del dashboard (generado, no editar)
├── analisis_biogrow.py      → Flujo de análisis, machote de Excel, nombres técnicos
├── graficas_biogrow.py      → Gráficas del tablero
├── exportar_biogrow.py      → Exportación de varias gráficas a PDF/Excel
├── reporte_biogrow.py       → Reporte PDF de los cálculos
├── generar_motor.py         → Regenera motor_biogrow.py desde dashboardst.py
├── requirements.txt         → Dependencias de Python
├── BioGrow_Analytics.spec   → Configuración de PyInstaller para generar el .exe
├── assets/                  → Recursos (si existe assets/icon.ico se usa como icono)
├── experimentos.db          → HEREDADO: base SQLite de la versión anterior (solo respaldo, no se usa)
└── README.md
```

`build/`, `dist/` y `__pycache__/` son generados automáticamente y están excluidos de Git (`.gitignore`).

---

## Requisitos

- Python 3 (desarrollado y probado con **Python 3.11**).
- Servidor **MySQL** en ejecución, con la base de datos `biogrow`.
- Dependencias de Python (ver `requirements.txt`), entre ellas:
  - `customtkinter`
  - `mysql-connector-python`

### Instalar dependencias

Desde la carpeta `BioGrow_Desktop`:

```bash
pip install -r requirements.txt
```

---

## Base de datos MySQL

La aplicación usa **exclusivamente MySQL**. La base de datos se llama **`biogrow`** y contiene dos tablas:

```sql
CREATE TABLE experimentos (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nombre_lote VARCHAR(150) NOT NULL,
    tipo_semilla VARCHAR(100) NOT NULL,
    fecha_registro DATETIME NOT NULL,
    tiempo_tratamiento VARCHAR(100) NOT NULL
);

CREATE TABLE mediciones_semillas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    experimento_id INT NOT NULL,
    num_muestra INT NOT NULL,
    longitud_raiz DECIMAL(10,2) NOT NULL,
    longitud_tallo DECIMAL(10,2) NOT NULL,
    cantidad_hojas INT NULL,
    germino BOOLEAN NOT NULL,
    CONSTRAINT fk_mediciones_experimento
        FOREIGN KEY (experimento_id) REFERENCES experimentos(id)
        ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT uq_muestra_experimento
        UNIQUE (experimento_id, num_muestra)
);
```

- `experimentos`: un registro por lote/experimento.
- `mediciones_semillas`: un registro por semilla medida, con su propia cantidad de hojas.
- Al iniciar, la aplicación ejecuta `CREATE TABLE IF NOT EXISTS` para ambas tablas. **No modifica tablas que ya existen** ni inserta datos de prueba.

> La base puede contener además la tabla `respaldo_hojas_experimento`, creada durante la migración que movió `cantidad_hojas` de `experimentos` a `mediciones_semillas`. Es solo un respaldo; la aplicación no la usa.

La administración de la base se realiza con **MySQL Workbench**.

### Configuración de la conexión (variables de entorno)

La conexión se configura con variables de entorno (ver `database.py`):

| Variable | Valor por defecto | Descripción |
|---|---|---|
| `BIOGROW_DB_HOST` | `localhost` | Servidor MySQL |
| `BIOGROW_DB_PORT` | `3306` | Puerto |
| `BIOGROW_DB_USER` | `root` | Usuario |
| `BIOGROW_DB_PASSWORD` | *(ninguno — obligatoria)* | Contraseña del usuario MySQL |
| `BIOGROW_DB_NAME` | `biogrow` | Nombre de la base de datos |

**La contraseña nunca se escribe en el código.** Si `BIOGROW_DB_PASSWORD` no está definida, la aplicación no se conecta y muestra un mensaje indicándolo.

Definirla de forma permanente para el usuario de Windows (cmd o PowerShell):

```bat
setx BIOGROW_DB_PASSWORD "TU_CONTRASEÑA"
```

`setx` solo afecta a las ventanas que se abran **después**: cierra y vuelve a abrir la terminal (y VS Code, si ejecutas desde ahí).

Definirla solo para la sesión actual de PowerShell:

```powershell
$env:BIOGROW_DB_PASSWORD = "TU_CONTRASEÑA"
```

Sustituye `TU_CONTRASEÑA` por la contraseña real. No la guardes en archivos del proyecto.

---

## Ejecutar desde el código fuente

Desde la carpeta `BioGrow_Desktop`:

```bash
python main.py
```

Flujo de arranque: `DatabaseManager` → `init_db()` → ventana principal (`BioGrowApp`).

### Comprobar la conexión

Desde la carpeta `BioGrow_Desktop` (no inserta datos):

```bash
python -c "from database import DatabaseManager; DatabaseManager().init_db(); print('Conexión con MySQL correcta')"
```

Si la conexión falla, el comando muestra el error correspondiente. Al abrir la aplicación, si no puede conectarse, **no se abre la interfaz** y aparece un mensaje de error (ver la sección siguiente).

### Solución de errores de conexión

| Mensaje / detalle | Causa probable | Qué hacer |
|---|---|---|
| "…falta configurar la variable de entorno BIOGROW_DB_PASSWORD" | La variable no existe en el entorno del proceso | Definirla con `setx` y reabrir la terminal/VS Code |
| `2003 (HY000): Can't connect to MySQL server on 'localhost:3306'` | MySQL no está iniciado, o host/puerto incorrectos | Iniciar el servicio MySQL; revisar `BIOGROW_DB_HOST` / `BIOGROW_DB_PORT` |
| `1045 (28000): Access denied for user …` | Usuario o contraseña incorrectos | Revisar `BIOGROW_DB_USER` / `BIOGROW_DB_PASSWORD` |
| `1049 (42000): Unknown database 'biogrow'` | La base no existe o el nombre no coincide | Crear la base `biogrow` en MySQL Workbench o revisar `BIOGROW_DB_NAME` |
| "…no pudo iniciar la interfaz… error al cargar los datos" (p. ej. `1054 Unknown column`) | Conectó, pero el esquema de las tablas no coincide con el esperado | Comparar las tablas con el esquema de este README (`SHOW CREATE TABLE …`) |

---

## Registro de mediciones

### Reglas de la tabla de captura

- La numeración **1–27 es fija**; una fila conserva su número aunque haya filas vacías antes.
- **Fila completamente vacía:** no se guarda.
- **Fila completa** (raíz, tallo y cantidad de hojas): se guarda como una medición con su número de muestra original. Ejemplo: si se capturan las filas 1, 3 y 5, se guardan las muestras 1, 3 y 5.
- **Fila incompleta** (falta cualquiera de los tres valores): bloquea el guardado con el mensaje *"La muestra N está incompleta. Debes capturar raíz, tallo y cantidad de hojas."*
- **Vacío y cero son distintos:** `0.00` en raíz/tallo y `0` en hojas son valores válidos.
- **Raíz y tallo:** número mayor o igual a 0; se acepta punto o coma decimal. Se muestran con 2 decimales (MySQL los almacena como `DECIMAL(10,2)`).
- **Cantidad de hojas:** número entero mayor o igual a 0. Se rechazan negativos, decimales y texto.
- Un valor inválido se señala al salir de la celda y la celda permanece en edición para corregirlo.
- Debe capturarse al menos una medición para guardar el lote.
- El guardado es una única transacción: se guarda el experimento con todas sus mediciones, o no se guarda nada.

### Regla de germinación

Definida en un único lugar: `calcular_germino()` en `utils.py`.

```
Germinó     si  longitud_raiz > 0.00  O  longitud_tallo > 0.00
No germinó  si  longitud_raiz = 0.00  Y  longitud_tallo = 0.00
```

La cantidad de hojas no interviene en la regla. El estado se muestra en la tabla en cuanto la fila tiene raíz y tallo.

### Navegación por teclado (tabla de captura)

| Tecla | Acción |
|---|---|
| `Doble clic` | Editar celda (Raíz, Tallo o Cantidad de hojas) |
| `Enter` | Confirmar y pasar a la misma columna de la fila siguiente |
| `Tab` | Raíz → Tallo → Cantidad de hojas → Raíz de la fila siguiente |
| `↑` `↓` | Mover entre filas |
| `←` `→` | Mover entre columnas editables |
| `Escape` | Cancelar la edición de la celda |

---

## Historial consolidado

Muestra una fila por cada semilla medida:

| Lote Origen | Tipo Semilla | No. | Tratamiento | Cant. Hojas | Raíz (mm) | Tallo (mm) | Resultado |
|---|---|---|---|---|---|---|---|

- **Búsqueda:** por lote, tratamiento o tipo de semilla.
- **Filtro por tipo de semilla.**
- **Filtro por germinación:** `1` = germinó, `0` = no germinó, vacío = todos.
- Mediciones registradas antes de la migración pueden no tener cantidad de hojas; se muestran con esa celda vacía.
- El historial es de solo lectura.

---

## Generar el ejecutable (.exe)

Desde la carpeta `BioGrow_Desktop`:

```bash
pip install pyinstaller
pyinstaller BioGrow_Analytics.spec
```

El ejecutable se genera en `dist/BioGrow_Analytics.exe`.

- Usa el archivo `BioGrow_Analytics.spec` (no el comando `pyinstaller main.py`): incluye los módulos y bibliotecas de `mysql.connector` que PyInstaller no detecta automáticamente.
- El `.exe` **también lee la contraseña de `BIOGROW_DB_PASSWORD`**: el equipo donde se ejecute debe tener la variable definida y acceso al servidor MySQL.

---

## Archivos heredados

- `experimentos.db` (en esta carpeta y en la raíz del proyecto): base SQLite de la versión anterior. Se conserva como respaldo; la aplicación actual **no** la utiliza.
- La versión web original (Flask: `app.py`, `templates/`, `styles/` en la raíz del proyecto) no forma parte de la aplicación de escritorio.

---

## Cálculos y Modelos / Tablero & Gráficas (análisis del dashboard)

Las pestañas 2 y 3 hacen **los mismos cálculos que el dashboard de Streamlit (`dashboardst.py`)**:

| Archivo | Qué hace |
|---|---|
| `motor_biogrow.py` | Copia fiel de las funciones de cálculo de `dashboardst.py` (lector de Excel, descartes, promedio por caja, descriptivos, ANOVA, Tukey, LSD, letras, cuartiles). Generado con `generar_motor.py`; no se edita a mano. |
| `analisis_biogrow.py` | Convierte los registros de la base (pestaña 1) al mismo formato que los Excel y ejecuta el flujo de análisis del dashboard. |
| `graficas_biogrow.py` | Las 12 gráficas del tablero (matplotlib): las 9 del dashboard (incluida cuartiles con su descripción) más Barras + Puntos, Tendencia de grupos y Combinada (cuartiles + puntos + promedio ± DE + letras). |
| `exportar_biogrow.py` | Exporta una o varias gráficas a un solo PDF o Excel (con hoja «Proyecto», imágenes de las gráficas y, si se elige, las tablas de cálculos). |
| `reporte_biogrow.py` | Reporte PDF de los cálculos (botón «📄 Reporte PDF de los cálculos» en Cálculos y Modelos). |
| `generar_motor.py` | Regenera `motor_biogrow.py` si cambia un cálculo en `dashboardst.py`: `python generar_motor.py ruta/a/dashboardst.py motor_biogrow.py` |

- **Fuente de datos:** la base de datos (lo registrado en la pestaña 1), los Excel cargados, o ambos. Los Excel se leen con el lector del dashboard (todos sus formatos, incluido el exportado por esta app).
- **Cálculos y Modelos:** variable (raíz, tallo, germinación), comparar por, variedad, cada dato es caja o semilla (si cada grupo tiene una sola caja se usa la semilla automáticamente), unidad mm/cm, α 5 % o 1 % y descartes. Arriba, el dictamen del ANOVA y el resumen de los datos. Sub-pestañas: Resumen, Promedio/DE/varianza, ANOVA, Tukey y LSD. El botón **📄 Reporte PDF de los cálculos** genera un PDF breve: dictamen del ANOVA, ranking con letras y gráfica de promedios, promedio/DE/varianza, tabla ANOVA, y Tukey y LSD (valor crítico, GL, letras y comparaciones).
- **Tablero & Gráficas:** KPIs y las gráficas con letras de Tukey o LSD, más la gráfica de cuartiles (Q1, mediana, Q3, IQR y atípicos) con una descripción de los datos; usan la misma configuración de Cálculos y Modelos, más los filtros rápidos de variedad, tratamiento y fechas. La barra de herramientas permite zoom y desplazamiento, y la gráfica se guarda en PNG/PDF a 300 DPI.
- `analytics_engine.py` ya no se usa (quedó incompleto); se conserva solo como referencia.

### Excel en Cálculos y Modelos: recordar e importar

- **Recordar Excel** (casilla, activada por defecto): la lista de Excel cargados se guarda en `biogrow_excels.json`, junto a `biogrow.db`, y se vuelve a cargar sola al abrir la app. Si un archivo ya no existe, se omite. El archivo está en `.gitignore` porque guarda rutas de cada computadora.
- **📥 Importar a la base:** pasa las semillas de los Excel cargados a `biogrow.db`, con un lote por caja. Después aparecen en el Historial y se pueden editar o borrar como cualquier registro. Antes de guardar muestra un resumen y pide confirmación; se guarda todo o nada.
  - Las longitudes se guardan en mm y se conserva el estado de germinación del Excel. La cantidad de hojas se importa cuando el Excel la trae (tablas con una semilla por fila); si no, queda vacía.
  - No se repiten lotes que ya están en la base (mismo nombre y variedad), ni se importan cajas vacías de plantilla ni hojas con un tratamiento por columna (ahí raíz y tallo no son de las mismas semillas).
  - Los lotes sin fecha en el Excel se registran con la fecha del día de la importación.
  - La base no tiene campo de parte expuesta: en archivos con Embrión / Endospermo, esa información queda solo en el nombre del lote (p. ej. `T-1M · rep 1 · …`).
  - Los archivos importados se quitan de «Excel cargados» para no contarlos dos veces.

### Novedades: factores de tratamiento, combinaciones y exportación

- **Machote (plantilla) de Excel:** botón «📋 Plantilla Excel» en Cálculos y Modelos. Genera un Excel vacío con las columnas que entiende la app (Lote, Variedad, No. Muestra, Tratamiento, Temperatura, Tipo de gas, Flujo, Orientación, Morfología, Hojas, Raíz, Tallo, Estado, Fecha) y una hoja «Léeme» con la descripción de cada columna. «Exportar Excel» del historial usa el mismo formato, así que se puede volver a cargar.
- **Nombres técnicos de variedades:** botón «🏷 Nombres técnicos». Asocia el nombre común con el técnico (p. ej. *Maíz blanco → PBB-DH-183*); se guardan en `biogrow_variedades.json` y aparecen en tablas, gráficas y reportes.
- **Combinación de tratamientos ("lenteja"):** la opción «Combinación de tratamiento» junta tiempo + temperatura + gas + flujo + orientación + morfología (solo los factores que cambian), para comparar cada combinación como un grupo. También se puede comparar por gas, flujo, orientación o morfología por separado.
- **Tarjetas en Cálculos y Modelos:** semillas, germinación, grupos y resultado del ANOVA de un vistazo.
- **Tablero:** filtro por tipo de gas; gráfica de **Tendencia** (eje X: tiempo, temperatura, flujo o fecha; una línea por variedad, gas, orientación, morfología o temperatura, y botón «Elegir series…»); **Barras + Puntos** y **Combinada** (varios análisis en una sola gráfica).
- **Guardar gráfica** en PNG, PDF o **Excel**; **🗂 Varias gráficas → PDF / Excel** junta las gráficas elegidas en un solo archivo, con la información del proyecto (variedades, lotes, semillas, tratamientos, gases, flujos, fechas) y, en Excel, las tablas de cálculos y los datos de cada grupo.
- El reporte PDF de los cálculos incluye la tabla «Información del proyecto».
- La base se actualiza sola al abrir la app (agrega las columnas `gas`, `flujo`, `orientacion`, `morfologia`, `cantidad_semillas` sin borrar datos). Los registros anteriores quedan como *Sin gas · Aleatorio · Plana*.
