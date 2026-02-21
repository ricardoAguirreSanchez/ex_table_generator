# Exp Table Generator

Herramienta que genera documentos Word con tablas de experiencias a partir de un Excel.
Lee un **template Word** para detectar las columnas automáticamente y las mapea al Excel.

Funciona 100% en el navegador — no requiere instalación. Ideal para [GitHub Pages](https://pages.github.com/).

---

## Uso web (recomendado)

1. Abrí `index.html` en tu navegador (doble clic o servidor local)
2. Subí un template Word (.docx)
3. Subí un Excel (.xlsx) y hacé clic en "Cargar Excel"
4. Revisá o ajustá el mapeo de columnas
5. Ingresá las filas a incluir (ej: `50, 51` o `10-15`)
6. Clic en **Generar Word** — se descargará el archivo

---

## Desplegar en GitHub Pages

1. Subí el proyecto a un repositorio de GitHub
2. En el repo: **Settings** → **Pages**
3. En **Source** elegí la rama (p. ej. `main`) y la carpeta raíz
4. Guardá — la página quedará disponible en:
   `https://<usuario>.github.io/<repo>/`

### Estructura para GitHub Pages

```
ex_table_generator/
├── index.html
├── css/
│   └── styles.css
├── js/
│   └── app.js
└── README.md
```

---

## Cómo funciona

### Paso 1 — Template Word
Elegís el Word que sirve de modelo. La herramienta detecta automáticamente las columnas, anchos y formato de la tabla.

### Paso 2 — Excel
Elegís el Excel con los datos. Indicás la hoja y la fila de encabezado. Se muestra una vista previa de las filas.

### Paso 3 — Mapeo
Se mapea cada columna del template a la del Excel. Podés ajustar el mapeo y los formatos.

| Opción | Qué hace |
|--------|----------|
| `(vacío)` | La columna queda vacía en el Word |
| `(auto-incremento)` | Numera las filas 1, 2, 3... |
| `(extraer país)` | Detecta el país a partir de la entidad contratante |
| `valor_tal_cual` | Copia el valor tal cual del Excel |
| `short_date` | Convierte "Agosto 2021" → "ago-21" |

### Paso 4 — Filas
Formato de entrada:
- `50, 51` — filas individuales
- `10 20 30` — separadas por espacio
- `10-20` — rango
- `5, 10-15, 20` — combinación

---

## Versión Python (escritorio)

También existe una versión de escritorio con interfaz gráfica:

```
python exp_table_generator.py
```

Requisitos: Python 3.9+, `pip install -r requirements.txt`

---

## Privacidad

Todo el procesamiento se hace en tu navegador. Los archivos no se envían a ningún servidor.
