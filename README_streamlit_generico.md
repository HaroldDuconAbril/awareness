# Generador Flexible de Awareness / Recordación

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run streamlit_awareness_generico.py
```

## Streamlit Cloud

Main file path:

```text
streamlit_awareness_generico.py
```

## Notas

- Incluye modo Bancos y modo Personalizado.
- En modo Bancos, Davibank / DaviBank no se cuenta como Davivienda.
- No usa `DataFrame.applymap`, por compatibilidad con pandas nuevos.
