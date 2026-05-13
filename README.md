# Anexo 1 Tutor v5

Aplicación web en Streamlit para diligenciar el Anexo 1 PTAFI 3.0 sobre la plantilla oficial del Ministerio, sin alterar su estructura.

## Archivos principales

- `app.py`: aplicación principal.
- `requirements.txt`: dependencias para Streamlit Cloud.
- `assets/logo_ptafi_transparente.png`: logo usado en el banner.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Cloud

1. Subir todos estos archivos al repositorio de GitHub.
2. Entrar a Streamlit Cloud.
3. Crear o actualizar la app usando:
   - Repository: `anexo1_tutor`
   - Branch: `main`
   - Main file path: `app.py`
4. Presionar **Deploy** o **Reboot app**.

## Flujo de uso

1. Configurar tutor, entrega y DANE.
2. Cargar Anexo 1 oficial.
3. Cargar base docente.
4. Seleccionar semana.
5. Seleccionar docente individual o grupo de docentes.
6. Marcar únicamente las actividades realizadas con “Sí”.
7. Descargar el archivo final con nombre `E#_DANE.xlsx`.


## v5
Corrige la generación del Excel para evitar reparación en Microsoft Excel. Conserva listas desplegables, formatos y hoja oculta escribiendo sobre la plantilla oficial con openpyxl.
