# Anexo 1 Tutor

Aplicación web local para diligenciar el Anexo 1 PTAFI 3.0 sin alterar la plantilla oficial del Ministerio.

## Instalación

1. Instale Python 3.10 o superior.
2. Abra una terminal en esta carpeta.
3. Ejecute:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Flujo de uso

1. Escriba el nombre del tutor.
2. Seleccione la entrega: E1 a E10.
3. Ingrese Código DANE del EE y de la sede.
4. Cargue el archivo Anexo 1 oficial del Ministerio.
5. Cargue la base de docentes en Excel o CSV.
6. Seleccione la semana.
7. Seleccione docentes en modo individual o grupal.
8. Marque solo las actividades realizadas; las demás se guardan como “No”.
9. Agregue más registros o descargue el archivo final.

El archivo final se descarga con el nombre `E#_DANE.xlsx`, por ejemplo `E1_218150000578.xlsx`.
