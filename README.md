
# Generador Balance Publicación - Áreas de impresión

Esta versión trabaja con el método definitivo:

- Hoja `Carátula`: 1 área de impresión.
- Hoja `Bal Public ...`: varias áreas de impresión.
- Cada área de impresión se convierte en una página del Word.

No abre Excel ni Word. Usa LibreOffice en modo oculto para renderizar cada área como PDF/imagen.

## Requisitos

1. Python instalado.
2. LibreOffice instalado.
3. Instalar librerías:

```cmd
pip install -r requirements.txt
```

## Ejecutar

```cmd
streamlit run app.py
```

## Cómo preparar el Excel

### Carátula
Definir 1 área de impresión.

### Balance
En la misma hoja del balance, seleccionar cada página y usar:

`Diseño de página > Área de impresión > Agregar al área de impresión`

La app lee automáticamente esas áreas.

## Orientación

La app detecta vertical/horizontal automáticamente, pero permite editar:

```text
A18:E85 | V
A561:L595 | H
A602:P642 | H
```

- `V` = vertical
- `H` = horizontal
- `AUTO` = detección automática

## Importante

Si LibreOffice no está instalado, la app no puede generar el Word.
