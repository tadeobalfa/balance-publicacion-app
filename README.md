
# Balance Publicación App

Aplicación web para convertir balances de publicación desde Excel a Word.

## Método

El Excel debe tener:

1. Hoja de carátula con 1 área de impresión.
2. Hoja del balance con varias áreas de impresión.
3. Cada área de impresión se convierte en una página del Word.

La app:
- Lee las áreas de impresión.
- Permite corregir orientación vertical/horizontal.
- Usa LibreOffice en el servidor para renderizar cada área.
- Genera un Word descargable.

## Archivos

- `app.py`: aplicación Streamlit.
- `requirements.txt`: librerías Python.
- `Dockerfile`: instalación de Python + LibreOffice para Render.
- `.dockerignore`: archivos a ignorar al desplegar.

## Despliegue en Render

1. Crear repositorio en GitHub.
2. Subir estos archivos:
   - app.py
   - requirements.txt
   - Dockerfile
   - .dockerignore
   - README.md
3. En Render crear un Web Service.
4. Seleccionar el repositorio.
5. Elegir Runtime: Docker.
6. Crear el servicio.

## Uso

1. Entrar a la URL de Render.
2. Cargar Excel.
3. Seleccionar hoja de carátula y hoja del balance.
4. Verificar áreas detectadas.
5. Generar Word.
6. Descargar Word.
