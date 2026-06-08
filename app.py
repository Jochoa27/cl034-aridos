"""
app.py — Entry point alternativo para Streamlit Cloud
Delega todo a dashboard_aridos_v5.py (evita doble set_page_config)
Uso: py -m streamlit run app.py --server.port 8505
"""
import runpy
from pathlib import Path

_target = str(Path(__file__).parent / "dashboard_aridos_v5.py")
runpy.run_path(_target, init_globals={"__file__": _target}, run_name="__main__")
