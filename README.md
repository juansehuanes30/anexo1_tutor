# Anexo 1 Tutor v16

Aplicación Streamlit para diligenciar el Anexo 1 PTAFI sobre la plantilla oficial.

## Ajuste v16

Esta versión parte de la versión 15 funcional y realiza únicamente la corrección del módulo **Editar registro existente**.

### Corrección aplicada

- Al seleccionar un registro para editar, ahora se cargan correctamente las actividades ya guardadas con “Sí”.
- El tutor puede agregar o quitar actividades del registro existente.
- Al guardar, se actualiza el registro en memoria sin crear duplicados.
- Se muestra confirmación de actualización.
- Se reinicia el editor para evitar estados anteriores de Streamlit.
- Se agrega botón “Cancelar edición”.

No se modifica la lógica de generación del Excel, validaciones, listas desplegables, flujo de semanas, finalización de entrega ni estructura funcional existente.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```
