# Anexo 1 Tutor v10

Versión ajustada con:

- Columna `CODIGO DANE SEDE` tomada desde la base docente para diligenciar la columna J del Anexo 1.
- Se conserva el campo `Código DANE del EE` en el Paso 0 para diligenciar la columna I.
- Se elimina el campo manual de Código DANE sede.
- Validación de duplicados por combinación exacta `semana de columna A + cédula`.
- Sección para editar registros existentes y agregar/quitar actividades sin crear duplicados.
- Flujo de continuidad: continuar en semana, finalizar semana y continuar con otra, finalizar entrega.
- La descarga se habilita únicamente después de finalizar la entrega.
- Se conserva el motor XML de escritura sobre plantilla oficial para mantener estructura, listas desplegables y evitar mensajes de reparación.

## Ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py
```
