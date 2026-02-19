# Exp Table Generator

Herramienta que genera documentos Word con tablas de experiencias a partir de un Excel.
Lee un **template Word** para detectar las columnas automáticamente y las mapea al Excel.

---

## Requisitos

- Windows 10 o superior
- Python 3.9 o superior ([descargar](https://www.python.org/downloads/))

> Al instalar Python, **marcar la casilla "Add Python to PATH"**.

---

## Instalación

1. Clonar o descargar este repositorio
2. Abrir CMD o PowerShell en la carpeta del proyecto
3. Ejecutar:

```
pip install -r requirements.txt
```

---

## Uso directo (con Python)

```
python exp_table_generator.py
```

---

## Generar .exe (opcional)

Si querés un ejecutable que funcione sin tener Python instalado:

```
pyinstaller --onefile --windowed --name "exp_table_generator" exp_table_generator.py
```

El .exe queda en `dist\exp_table_generator.exe`.
Se puede copiar a cualquier PC con Windows y funciona sin instalar nada.

---

## Cómo funciona

Al abrir la herramienta aparece una ventana con 4 pasos:

### Paso 1 — Seleccionar template Word

Elegís el Word que sirve de modelo. La herramienta detecta automáticamente
las columnas, anchos y formato de la tabla.

### Paso 2 — Seleccionar Excel

Elegís el Excel con los datos. Se muestra una vista previa de las filas disponibles.

### Paso 3 — Revisar mapeo de columnas

La herramienta mapea automáticamente cada columna del template a la columna
del Excel que mejor coincida. Podés ajustar cualquier mapeo con los dropdowns.

Opciones especiales disponibles:

| Opción | Qué hace |
|--------|----------|
| `(auto-incremento)` | Numera las filas 1, 2, 3... |
| `(extraer país)` | Detecta el país a partir de la entidad contratante |
| `short_date` (formato) | Convierte "Agosto 2021" → "ago-21" |

### Paso 4 — Elegir filas y generar

Escribís los números de fila del Excel que querés incluir y hacés click en **Generar Word**.

Formatos válidos:
- `50, 51` — filas individuales
- `10 20 30` — separadas por espacio
- `10-20` — rango de filas
- `5, 10-15, 20` — combinación

---

## Cambiar el template

No hay que modificar código. Si el template cambia (columnas distintas, otro formato):

1. Abrí la herramienta
2. Seleccioná el nuevo Word como template
3. Las columnas se detectan automáticamente
4. Ajustá el mapeo si es necesario
5. Generá
