import tkinter as tk
import subprocess
import threading
import queue
import os

# Configuración de colores
COLOR_OFF = "red"
COLOR_ON_LIGHT = "#90EE90"  # Verde claro
COLOR_ON_DARK = "#228B22"   # Verde oscuro

q = queue.Queue()
info_label = None

# Diccionario para rastrear si cada monitor está activo (ahora hasta 10)
running_states = {i: False for i in range(1, 11)}
# Diccionario para guardar las referencias a los Canvas (círculos)
status_indicators = {}

def toggle_blink(monitor_id, state):
    """Hace que el círculo parpadee si el proceso está activo."""
    if not running_states[monitor_id]:
        status_indicators[monitor_id].itemconfig("circle", fill=COLOR_OFF)
        return

    new_color = COLOR_ON_DARK if state else COLOR_ON_LIGHT
    status_indicators[monitor_id].itemconfig("circle", fill=new_color)
    root.after(500, lambda: toggle_blink(monitor_id, not state))

def run_digitacion_thread(monitor_id, script_path):
    """Función genérica para ejecutar los scripts."""
    running_states[monitor_id] = True
    root.after(0, lambda: toggle_blink(monitor_id, True)) 
    
    try:
        process = subprocess.Popen(
            ['python', script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )
        
        for line in iter(process.stdout.readline, ''):
            q.put(f"M{monitor_id} -> {line.strip()}")
            
        process.stdout.close()
        process.wait()
        q.put(f"--- Monitor {monitor_id} Finalizado ---")
    except Exception as e:
        q.put(f"Error en M{monitor_id}: {e}")
    finally:
        running_states[monitor_id] = False 

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
                info_label.pack(pady=5, side=tk.BOTTOM)
            
            current_text = info_label.cget("text")
            lines = (current_text + "\n" + line).split('\n')[-12:]
            info_label.config(text="\n".join(lines))
            q.task_done()
    except queue.Empty:
        pass
    root.after(100, update_gui)

def ejecutar_monitor(id):
    script = f'./Digitacion/digitacion{id}.py'
    if not running_states[id]:
        threading.Thread(target=run_digitacion_thread, args=(id, script), daemon=True).start()

# --- Interfaz Principal ---
root = tk.Tk()
root.title("TG Bots (Digitación)")
root.geometry("400x600")

# Frame principal con pack
frame = tk.Frame(root, padx=25, pady=25)
frame.pack(expand=True, fill="both")

tk.Button(frame, text="Regularización Material", bg="#f0f0f0").pack(pady=10, fill=tk.X)

# --- CONTENEDOR PARA EL GRID ---
# Creamos un frame específico para los monitores para evitar el error de mezcla
grid_container = tk.Frame(frame)
grid_container.pack(fill=tk.X, pady=5)

grid_container.columnconfigure(0, weight=1)
grid_container.columnconfigure(1, weight=1)

# Generar 10 monitores en 2 columnas
for i in range(1, 11):
    fila = (i - 1) // 2
    columna = (i - 1) % 2
    
    # Frame de celda
    f = tk.Frame(grid_container, bd=1, relief=tk.RIDGE, padx=5, pady=5)
    f.grid(row=fila, column=columna, sticky="nsew", padx=3, pady=3)
    
    light = create_status_light(f)
    light.pack(side=tk.LEFT, padx=2)
    status_indicators[i] = light
    
    btn = tk.Button(f, text=f" Monitor {i}", font=("Arial", 8),
                    command=lambda idx=i: ejecutar_monitor(idx))
    btn.pack(side=tk.LEFT, expand=True, fill=tk.X)
    
    path_log = f'.\\Digitacion\\log_ejecucion_monitor{i}.csv'
    
    tk.Button(f, text="+", width=2,
              command=lambda p=path_log: os.startfile(p) if os.path.exists(p) else None).pack(side=tk.LEFT, padx=1)
              
    tk.Button(f, text="-", width=2,
              command=lambda p=path_log: os.remove(p) if os.path.exists(p) else None).pack(side=tk.LEFT, padx=1)

# Botones finales y Log
tk.Button(frame, text="Salir", command=root.destroy, bg="#bcb9b9").pack(pady=5, fill=tk.X)
tk.Label(frame, text="---- Log Operacional ----", font=("Arial", 9, "bold")).pack()

root.after(100, update_gui)
root.mainloop()