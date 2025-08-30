# Importar librerias
import streamlit as st
import runpy
from io import StringIO
import contextlib
from pathlib import Path
import tempfile, json, os, shutil
import traceback

st.set_page_config(page_title="Control HGI", page_icon="📊", layout="wide")
st.title("App de Control HGI")
st.write("Presiona el botón para ejecutar tu script original.")

# Ajusta si tu script está en otra carpeta
SCRIPT_PATH = Path(__file__).parent / "ControlHGI.py"

# Nombre de archivo que tu script original espera (no subas este archivo al repo)
JSON_FILENAME = "Llave_JSON.json"

def write_service_account_json(target_path: Path):
    """
    Crea el archivo JSON de credencial a partir de st.secrets.
    Espera que exista st.secrets['google_service_account'] con todo el contenido del JSON.
    """
    if "google_service_account" not in st.secrets:
        raise RuntimeError(
            "No encontré 'google_service_account' en Secrets. "
            "Configúralo en Settings → Secrets."
        )
    data = dict(st.secrets["google_service_account"])
    target_path.write_text(json.dumps(data), encoding="utf-8")

if st.button("Actualizar Controles"):
    if not SCRIPT_PATH.exists():
        st.error(f"No encuentro el archivo: {SCRIPT_PATH}")
    else:
        st.info("Ejecutando script...")

        stdout_buffer, stderr_buffer = StringIO(), StringIO()

        # Creamos una carpeta temporal para la llave,
        # y la ubicamos junto al script (o en cwd) con el nombre esperado.
        temp_dir = Path(tempfile.mkdtemp())
        json_path = Path.cwd() / JSON_FILENAME  # el script lo buscará aquí

        try:
            # Escribimos la credencial desde los secrets
            write_service_account_json(json_path)

            # (Opcional) también puedes exponerla como variable de entorno estándar de Google:
            # os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = str(json_path)

            with contextlib.redirect_stdout(stdout_buffer), contextlib.redirect_stderr(stderr_buffer):
                runpy.run_path(str(SCRIPT_PATH), run_name="__main__")

            st.success("¡Actualización completada!")
        except Exception:
            st.error("Ocurrió un error durante la ejecución.")
            st.code(traceback.format_exc())
        finally:
            # Limpieza: borra la llave del disco
            try:
                if json_path.exists():
                    json_path.unlink()
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass

        out, err = stdout_buffer.getvalue().strip(), stderr_buffer.getvalue().strip()
        if out:
            st.subheader("Salida (stdout)")
            st.code(out)
        if err:
            st.subheader("Errores/Advertencias (stderr)")
            st.code(err)
