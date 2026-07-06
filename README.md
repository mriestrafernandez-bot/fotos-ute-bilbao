# Fotos UTE Bilbao

Webapp local para leer datos de una foto JPEG y exportarlos a Excel.

## Requisitos

- Python 3.9 o superior.
- Una clave de OpenAI API con acceso a modelos con vision.
- Conexion a internet mientras se convierte la foto.

## Instalacion en este ordenador

1. Abre Terminal en esta carpeta:

   ```bash
   cd "/Users/mariariestrafernandez/Documents/UTE bilbao"
   ```

2. Crea y activa un entorno virtual:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. Instala dependencias:

   ```bash
   pip install -r requirements.txt
   ```

4. Crea un archivo `.env` con tu clave. Puedes copiar `.env.example` y cambiar el valor:

   ```bash
   OPENAI_API_KEY=tu_clave_aqui
   OPENAI_MODEL=gpt-5-mini
   ```

5. Ejecuta la app:

   ```bash
   streamlit run app.py
   ```

6. Abre la direccion que aparece en Terminal, normalmente:

   ```text
   http://localhost:8501
   ```

## Uso

La app tiene dos apartados:

### Excel preformado

1. Sube la foto JPEG.
2. Sube el Excel preformado donde quieres cargar los datos. Puede ser `.xlsx` o `.xls`.
3. Elige el mes de referencia en la barra lateral.
4. Pulsa `Convertir a tabla`.
5. Revisa y corrige la tabla si alguna lectura manuscrita no es perfecta.
6. Pulsa `Descargar Excel completado`.

### Foto a Excel

1. Sube la foto JPEG.
2. Normalmente deja desmarcada la opcion `Indicar numero de columnas manualmente` para que la app detecte las columnas.
3. Si una foto ancha no detecta todas las columnas, activa esa opcion e indica el numero total de columnas visibles.
4. Ajusta el ancho aproximado de la columna Fecha si hace falta.
5. Elige el mes de referencia en la barra lateral.
6. Pulsa `Convertir foto a Excel`.
7. Revisa y corrige la tabla.
8. Pulsa `Descargar Excel de la foto (.xlsx)`.

## Usar la webapp en otro ordenador

1. Copia esta carpeta completa al otro ordenador.
2. Instala Python 3.9 o superior desde <https://www.python.org/downloads/>.
3. Abre Terminal o PowerShell dentro de la carpeta copiada.
4. Crea el entorno virtual:

   macOS/Linux:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

   Windows PowerShell:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

5. Instala dependencias:

   ```bash
   pip install -r requirements.txt
   ```

6. Crea un archivo `.env` dentro de la carpeta. Puedes copiar `.env.example`:

   ```text
   OPENAI_API_KEY=tu_clave_aqui
   OPENAI_MODEL=gpt-5-mini
   ```

7. Ejecuta:

   ```bash
   streamlit run app.py
   ```

## Notas importantes

- La lectura de escritura a mano no es infalible. La app incluye una tabla editable para verificar antes de generar el Excel.
- Si el Excel tiene encabezados como `fecha`, `numero`, `confianza` o `notas`, la app escribira debajo de esos encabezados. Si no los encuentra, escribira en las primeras columnas libres de la hoja activa.
- Si el Excel ya tiene informacion, los datos nuevos se anadiran en las filas inferiores, debajo de la ultima fila con contenido.
- Si subes `.xlsx`, descargara `.xlsx`. Si subes `.xls`, descargara `.xls`.
- Las cifras dudosas se cargan en rojo cuando la confianza es baja, cuando el numero queda vacio o cuando hay una nota de duda.
- En el apartado `Foto a Excel`, la app intenta reproducir las columnas visibles de la foto. Las celdas dudosas se marcan en rojo en el Excel descargado.
- Para tablas anchas, puedes indicar manualmente el numero de columnas. En ese modo la app lee la foto por tramos para no perder las columnas de la derecha.
- Si una foto esta muy oscura, torcida o borrosa, conviene repetirla con buena luz y buena resolucion.
- Puedes cambiar `OPENAI_MODEL` por otro modelo con vision si tu cuenta lo permite.

