# Análisis Integral de la Flota Vehicular – PESV

Dashboard en Streamlit (Python) que reemplaza al dashboard original en React/JSX.
Carga un Excel/CSV de bitácora de flota y muestra KPIs, destacados, recorrido,
combustible/eficiencia, seguridad, estacionamiento y huella de carbono.

## Archivos
- `app.py` — la aplicación completa.
- `requirements.txt` — dependencias.
- `informe_junio_demo.xlsx` — dataset de ejemplo precargado (el mismo que traía el dashboard original).

## Estructura de columnas esperada en el archivo a cargar
Placa ID | Fecha | Kilometraje(km) | Exceso de velocidad | Estacionamiento | Combustible(gal (us))

(La app también acepta variaciones del nombre de columna, ej. "Placa", "Combustible (gal)", etc.,
siempre que contengan esas palabras clave.)

---

## Opción A — Publicarlo como página web gratis (recomendado)

1. Crea un repositorio en GitHub y sube estos 3 archivos (`app.py`, `requirements.txt`,
   `informe_junio_demo.xlsx`).
2. Entra a https://share.streamlit.io con tu cuenta de GitHub.
3. Clic en "New app", selecciona el repo, la rama y `app.py` como archivo principal.
4. Deploy. Streamlit Community Cloud te da una URL pública gratuita
   (algo como `tuapp.streamlit.app`) que puedes compartir con quien quieras.

## Opción B — Correrlo localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en `http://localhost:8501`.

## Opción C — Correrlo desde Google Colab

Streamlit no corre "dentro" de una celda de Colab como un notebook normal —
se lanza como servidor y se expone con un túnel. En una celda de Colab:

```python
!pip install streamlit pandas numpy plotly openpyxl pyngrok -q

# sube app.py, requirements.txt e informe_junio_demo.xlsx a /content primero
# (panel izquierdo de Colab -> Archivos -> subir)

!streamlit run /content/app.py &>/content/log.txt &

from pyngrok import ngrok
ngrok.set_auth_token("TU_TOKEN_DE_NGROK")  # gratis en https://ngrok.com
public_url = ngrok.connect(8501)
print(public_url)
```

Nota: para uso serio y para compartir con clientes/inversionistas, la Opción A
(Streamlit Community Cloud) es más estable y no depende de que Colab siga la sesión activa.
