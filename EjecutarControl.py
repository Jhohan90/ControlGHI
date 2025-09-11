# EjecutarControl.py
import streamlit as st
import runpy
from io import StringIO
import contextlib
from pathlib import Path
import json
import traceback

# ===== NUEVO: helpers de geolocalización (sutil) =====
def _load_get_geolocation():
    try:
        from streamlit_js_eval import get_geolocation  # dependencia ligera
        return get_geolocation
    except Exception:
        return None

def request_user_location(key_suffix: str = "main"):
    """
    Dispara el diálogo del navegador para compartir ubicación (HTTPS o localhost).
    Guarda un dict minimal en st.session_state['loc_<suffix>'] o no hace nada si se deniega.
    """
    get_geolocation = _load_get_geolocation()
    if not get_geolocation:
        st.toast("Geolocalización no disponible (falta dependencia o contexto).", icon="⚠️")
        return

    try:
        loc = get_geolocation()
        if loc and isinstance(loc, dict) and loc.get("coords"):
            st.session_state[f"loc_{key_suffix}"] = {
                "latitude":  loc["coords"].get("latitude"),
                "longitude": loc["coords"].get("longitude"),
                "accuracy":  loc["coords"].get("accuracy"),
                "ts": loc.get("timestamp"),
                "source": "browser",
            }
            st.toast("Ubicación guardada ✅", icon="✅")
        else:
            st.toast("No se obtuvo la ubicación (quizá denegada).", icon="⚠️")
    except Exception:
        st.toast("No fue posible obtener la ubicación.", icon="⚠️")

def sidebar_geoloc(minimal: bool = True):
    """
    Inserta un botón compacto 📍 en la barra lateral (no bloquea la UI).
    Muestra un indicador de estado sin revelar coordenadas.
    """
    with st.sidebar:
        cols = st.columns([1, 9]) if minimal else st.columns([1, 7, 2])
        with cols[0]:
            pressed = st.button("📍", help="Compartir ubicación (opcional)", key="geo_btn")
        with cols[-1]:
            st.markdown(
                "<div style='font-size:0.85rem; opacity:0.85'>Ubicación (opcional)</div>",
                unsafe_allow_html=True
            )

        if pressed:
            request_user_location("main")

        status = "✔️" if st.session_state.get("loc_main") else "—"
        st.caption(f"Estado: {status}")
# ===== FIN helpers de geolocalización =====

st.set_page_config(page_title="Control GHI", page_icon="📊", layout="wide")

# Inserción sutil: añade el botón en la barra lateral SIN cambiar tu layout
sidebar_geoloc(minimal=True)

st.title("Control GHI")
st.write("Presiona el botón para actualizar los datos ingresados")

# Ruta del script original
SCRIPT_PATH = Path(__file__).parent / "ControlGHI.py"

# Nombre de archivo que tu script original espera
JSON_FILENAME = "Llave_JSON.json"
JSON_PATH = Path.cwd() / JSON_FILENAME

def escribir_llave_desde_secrets():
    """
    Crea 'Llave_JSON.json' a partir de st.secrets['google_service_account'].
    """
    if "google_service_account" not in st.secrets:
        raise RuntimeError(
            "No encontré 'google_service_account' en Secrets.\n"
            "Ve a Settings → Secrets y pega tu llave en formato TOML:\n\n"
            "[google_service_account]\n"
            'type = "service_account"\n...'
        )
    data = dict(st.secrets["google_service_account"])
    JSON_PATH.write_text(json.dumps(data), encoding="utf-8")

if st.button("Actualizar Datos"):
    if not SCRIPT_PATH.exists():
        st.error(f"No se encontró el archivo: {SCRIPT_PATH}")
    else:
        st.info("Ejecutando actualización... esto puede tardar algunos minutos.")
        out_buf, err_buf = StringIO(), StringIO()

        try:
            # 1) Crear la llave para que ControlGHI.py la encuentre
            escribir_llave_desde_secrets()

            # (Opcional) Si quisieras usar la ubicación dentro de tu ejecución:
            # loc = st.session_state.get("loc_main")
            # if loc:
            #     # Ejemplo: escribir un JSON efímero con coords para que lo lea ControlGHI.py
            #     Path("user_location.json").write_text(json.dumps(loc), encoding="utf-8")

            # 2) Ejecutar el script original como si fuera: python ControlGHI.py
            with contextlib.redirect_stdout(out_buf), contextlib.redirect_stderr(err_buf):
                runpy.run_path(str(SCRIPT_PATH), run_name="__main__")

            st.success("¡Actualización completada!")
        except Exception:
            st.error("Ocurrió un error durante la ejecución.")
            st.code(traceback.format_exc())
        finally:
            # 3) Limpieza: eliminar la llave del disco
            try:
                if JSON_PATH.exists():
                    JSON_PATH.unlink()
            except Exception:
                pass
