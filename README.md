# Anexo 1 Tutor v9

Aplicación Streamlit para diligenciar el Anexo 1 PTAFI sobre la plantilla oficial.

## Mejoras v9

- Paso 0 bloqueado hasta completar todos los datos obligatorios.
- Opción "Selecciona entrega" para obligar la selección de E1 a E10.
- Botón Continuar con alertas de datos faltantes.
- Modo de registro: Individual, Grupal y Todos.
- En modo Todos se selecciona toda la base y se permite retirar docentes antes de agregar.
- Limpieza automática de docente(s) y actividades después de agregar registros.
- Motor de Excel ajustado para conservar validaciones x14 y listas desplegables de la plantilla oficial.
- Escritura directa sobre la plantilla original sin reconstruir hojas.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```
