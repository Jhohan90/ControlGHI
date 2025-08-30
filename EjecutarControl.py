# EjecutarControl.py
import streamlit as st
import runpy
from io import StringIO
import contextlib
from pathlib import Path
import json
import traceback

st.set_page_config(page_title="Control HGI", page_icon="📊", layout="wide")
st.title("Control HGI")
st.write("Presiona el botón para ejecutar actualizar los datos ingresados")

# Ruta del script original
SCRIPT_PATH = Path(__file__).parent / "ControlHGI.py"

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
            # 1) Crear la llave para que ControlHGI.py la encuentre
            escribir_llave_desde_secrets()

            # 2) Ejecutar el script original como si fuera: python ControlHGI.py
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

        # Mostrar logs capturados
        out, err = out_buf.getvalue().strip(), err_buf.getvalue().strip()
        if out:
            st.subheader("Salida (stdout)")
            st.code(out)
        if err:
            st.subheader("Errores/Advertencias (stderr)")
            st.code(err)

