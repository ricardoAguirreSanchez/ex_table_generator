# Exp Table Generator — Versión Escritorio (Python)

Versión con interfaz gráfica (tkinter) para generar documentos Word con tablas de experiencias a partir de un Excel.

---

## Requisitos

- Windows 10 o superior (o macOS/Linux con Python)
- Python 3.9 o superior ([descargar](https://www.python.org/downloads/))

> Al instalar Python, **marcar la casilla "Add Python to PATH"**.

---

## Instalación

1. Abrir terminal/CMD en esta carpeta (`desktop_version`)
2. Crear entorno virtual (opcional pero recomendado):

   ```
   python -m venv venv
   venv\Scripts\activate
   ```

3. Instalar dependencias:

   ```
   pip install -r requirements.txt
   ```

---

## Uso

```
python exp_table_generator.py
```

Se abrirá una ventana con 4 pasos:

1. **Template Word** — Seleccionar el .docx modelo
2. **Excel** — Seleccionar el .xlsx y configurar hoja/fila encabezado
3. **Mapeo** — Revisar o ajustar la correspondencia de columnas
4. **Filas** — Indicar qué filas incluir (ej: `50, 51` o `10-15`)

---

## Generar ejecutable (.exe)

Para tener un .exe que funcione sin Python instalado:

```
pyinstaller --onefile --windowed --name "exp_table_generator" exp_table_generator.py
```

El archivo queda en `dist\exp_table_generator.exe`.

---

## Opciones de mapeo

| Opción | Qué hace |
|--------|----------|
| `(vacío)` | La columna queda vacía |
| `(auto-incremento)` | Numera 1, 2, 3... |
| `(extraer país)` | Detecta país desde entidad contratante |
| `valor_tal_cual` | Copia el valor tal cual (ideal para fórmulas) |
| `fecha_corta` | Convierte "Agosto 2021" → "ago-21" |

## Filas a incluir

- `50, 51` — filas individuales
- `10 20 30` — separadas por espacio
- `10-20` — rango
- `5, 10-15, 20` — combinación
