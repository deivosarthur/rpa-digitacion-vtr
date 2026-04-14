import tkinter as tk
import subprocess
import threading
import queue
import os
from pathlib import Path
import sys
# Configuración de colores
COLOR_OFF = "red"
COLOR_ON_LIGHT = "#90EE90"  # Verde claro
COLOR_ON_DARK = "#228B22"   # Verde oscuro

q = queue.Queue()
info_label = None

# Diccionario para rastrear si cada monitor está activo
running_states = {1: False, 2: False, 3: False, 4: False, 5: False}
# Diccionario para guardar las referencias a los Canvas (círculos)
status_indicators = {}

def toggle_blink(monitor_id, state):
    """Hace que el círculo parpadee si el proceso está activo."""
    if not running_states[monitor_id]:
        status_indicators[monitor_id].itemconfig("circle", fill=COLOR_OFF)
        return

    # Alternar entre verde claro y oscuro
    new_color = COLOR_ON_DARK if state else COLOR_ON_LIGHT
    status_indicators[monitor_id].itemconfig("circle", fill=new_color)
    
    # Programar el siguiente cambio en 500ms
    root.after(500, lambda: toggle_blink(monitor_id, not state))

def run_digitacion_thread(monitor_id, script_path):
    """Función genérica para ejecutar los scripts."""
    running_states[monitor_id] = True
    toggle_blink(monitor_id, True) # Iniciar parpadeo
    
    try:
        process = subprocess.Popen(
            [sys.executable, script_path],  # Ejecuta el script con el mismo intérprete de Python
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )
        
        for line in iter(process.stdout.readline, ''):
            q.put(f"NDC{monitor_id} -> {line.strip()}")
            
        process.stdout.close()
        process.wait()
        q.put(f"--- Monitor {monitor_id} Finalizado ---")
    except Exception as e:
        q.put(f"Error en M{monitor_id}: {e}")
    finally:
        running_states[monitor_id] = False # Detiene el parpadeo

def create_status_light(parent):
    """Crea un pequeño círculo indicador."""
    canvas = tk.Canvas(parent, width=20, height=20, highlightthickness=0)
    canvas.create_oval(2, 2, 14, 14, fill=COLOR_OFF, tags="circle", outline="")
    return canvas

def update_gui():
    global info_label
    try:
        while True:
            line = q.get_nowait()
            if info_label is None:
                info_label = tk.Label(frame, text="", justify="left", wraplength=350, font=("Consolas", 8))
                info_label.pack(pady=5)
            
            current_text = info_label.cget("text")
            # Limitar el log para que no crezca infinitamente (últimas 15 líneas)
            lines = (current_text + "\n" + line).split('\n')[-15:]
            info_label.config(text="\n".join(lines))
            q.task_done()
    except queue.Empty:
        pass
    root.after(100, update_gui)

# --- Funciones de disparo ---

from pathlib import Path

def ejecutar_monitor(id):


    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent

    script = base_dir / f"digitacion{id}.py"
    print(sys.executable)
    if not script.exists():
        q.put(f"❌ No existe: {script}")
        print(sys.executable)
        return

    if not running_states[id]:
        threading.Thread(
            target=run_digitacion_thread,
            args=(id, str(script)),
            daemon=True
        ).start()

# --- Interfaz Principal ---
root = tk.Tk()
root.title("TG Bots (Digitación)")
root.geometry("450x600")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(expand=True, fill="both")

tk.Button(frame, text="Regularización Material", bg="#f0f0f0").pack(pady=5, fill=tk.X)

# Generar filas para los 4 monitores
for i in [1, 2, 3, 4, 5]:
    f = tk.Frame(frame)
    f.pack(pady=5, fill=tk.X)
    
    # Indicador de luz
    light = create_status_light(f)
    light.pack(side=tk.LEFT, padx=5)
    status_indicators[i] = light
    
    # Botón Principal
    btn = tk.Button(f, text=f"Iniciar Declaración Monitor {i}", 
                    command=lambda idx=i: ejecutar_monitor(idx))
    btn.pack(side=tk.LEFT, expand=True, fill=tk.X)
    
    # Botones + y -
    path_log = f'.\\Digitacion\\log_ejecucion_monitor{i}.csv'
    tk.Button(f, text=" + ", command=lambda p=path_log: os.startfile(p) if os.path.exists(p) else None).pack(side=tk.LEFT, padx=2)
    tk.Button(f, text=" - ", command=lambda p=path_log: os.remove(p) if os.path.exists(p) else None).pack(side=tk.LEFT, padx=2)

tk.Button(frame, text="Salir", command=root.destroy, bg="#bcb9b9").pack(pady=8, fill=tk.X)
tk.Label(frame, text="---- Log Operacional ----", font=("Arial", 9, "bold")).pack()

root.after(100, update_gui)
root.mainloop()

