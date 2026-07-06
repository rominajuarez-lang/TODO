# Dashboard de inventarios, forecast y ahorro

Aplicación en Streamlit para:

- Ver la simulación de inventarios por SKU.
- Comparar demanda real, pronóstico empresa y pronóstico propuesto.
- Calcular ahorro económico en soles del pronóstico propuesto.

## Archivos para GitHub

- `app.py`
- `requirements.txt`

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud

1. Sube `app.py` y `requirements.txt` a GitHub.
2. En Streamlit Cloud selecciona tu repositorio.
3. En Main file path coloca:

```text
app.py
```

4. Abre la app y sube tu Excel desde el botón de carga.

## Hojas esperadas del Excel

- `Demanda`: columnas `date`, `product_id`, `demand_real`.
- `Datos`: columnas `product_id`, `initial_stock`, `lead_time_months`, `ss`, `lot_size`, `unit_value`, entre otras.
- `Forescast_Comercial`: columnas `date`, `product_id`, `forecast_company`.
