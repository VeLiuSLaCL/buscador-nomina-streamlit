# App Streamlit para búsqueda de empleado en Excel

## Archivos
- `app.py`
- `requirements.txt`

## Ejecución local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicación en Streamlit Community Cloud
1. Sube `app.py` y `requirements.txt` a un repositorio en GitHub.
2. Entra a Streamlit Community Cloud.
3. Elige **New app**.
4. Selecciona tu repositorio, rama y archivo principal: `app.py`.
5. Publica la app.

## Qué hace
- Permite subir un archivo `.xls`, `.xlsx` o `.xlsm`.
- Busca el número de empleado en la primera columna de todas las hojas.
- Devuelve:
  - Hoja
  - Número de empleado
  - Nombre
  - Mes
  - Periodo de nómina
  - UUID Vigente
  - /559 Transferencia
