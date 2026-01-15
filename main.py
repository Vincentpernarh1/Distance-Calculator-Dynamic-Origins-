import os
import json
import time
import math
import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import queue

# ======================================================
# CONFIG
# ======================================================
EXCEL_FILE = os.path.join(os.getcwd(), "routes.xlsx")
OUTPUT_FILE = os.path.join(os.getcwd(), "routes_with_distance.xlsx")
MATRIX_URL = None

CHUNK_SIZE  = 50
MAX_RETRIES = 5
BACKOFF_FACTOR = 2

# ======================================================
# LOAD API KEY
# ======================================================
with open(os.path.join(os.getcwd(), "credencial.json"), "r") as f:
    cred = json.load(f)
    API_KEY = cred["api_key"]
    MATRIX_URL = cred["url"]
    
HEADERS = {
    "Authorization": API_KEY,
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "application/json, application/geo+json, application/gpx+xml, img/png; charset=utf-8",
    "User-Agent": "curl/8.0.1"
}

# ======================================================
# HELPERS
# ======================================================
def parse_lon_lat(value):
    lon, lat = value.split("|")
    lon = float(lon.replace(",", "."))
    lat = float(lat.replace(",", "."))
    # lon = lon / (10 ** 14)
    return lon, lat

def post_with_retry(payload):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = requests.post(
                MATRIX_URL,
                headers=HEADERS,
                json=payload,
                timeout=60
            )
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"Attempt {attempt} failed: {e}")
            if attempt == MAX_RETRIES:
                raise
            time.sleep(BACKOFF_FACTOR ** attempt)

# ======================================================
# MAIN PROCESS
# ======================================================
def process_file():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name="Base")
        
        # Normalize numbers
        df["Longitude"] = df["Longitude"].astype(str).str.replace(",", ".").astype(float)
        df["Latitude"] = df["Latitude"].astype(str).str.replace(",", ".").astype(float)

        # ----------------------------------
        # BUILD ORIGINS FROM FILE (UNIQUE)
        # ----------------------------------
        origins = (
            df[["Origin", "Long|Lat"]]
            .drop_duplicates()
            .assign(coords=lambda x: x["Long|Lat"].apply(parse_lon_lat))
        )

        origin_map = dict(zip(origins["Origin"], origins["coords"]))

        df["distance_km"] = None

        # Count total chunks (for progress)
        total_chunks = 0
        for origin in origin_map:
            total_chunks += math.ceil(
                len(df[df["Origin"] == origin]) / CHUNK_SIZE
            )

        current_step = 0

        # ----------------------------------
        # PROCESS EACH ORIGIN DYNAMICALLY
        # ----------------------------------
        for origin_name, origin_coord in origin_map.items():
            subset = df[df["Origin"] == origin_name]

            indices = subset.index.tolist()
            destinations = subset[["Longitude", "Latitude"]].values.tolist()

            for i in range(0, len(destinations), CHUNK_SIZE):
                chunk_dest = destinations[i:i + CHUNK_SIZE]
                chunk_idx = indices[i:i + CHUNK_SIZE]

                locations = [list(origin_coord)] + chunk_dest

                # print(f"Locations count: {len(locations)}")
                # print(f"Origin: {locations[0]}")
                # print(f"Destinations: {locations[1:]}")

                payload = {
                    "locations": locations,
                    "metrics": ["distance"]
                }

                result = post_with_retry(payload)
                distances_m = result["distances"][0][1:]

                for idx, dist in zip(chunk_idx, distances_m):
                    if dist is not None:
                        df.at[idx, "distance_km"] = round(dist / 1000, 2)
                    else:
                        df.at[idx, "distance_km"] = None

                current_step += 1
                # print(f"Processing {origin_name} ({current_step}/{total_chunks})")

                # Save partial results to file
                df.to_excel(OUTPUT_FILE, index=False)
                print(f"Partial results saved to {OUTPUT_FILE}")

                time.sleep(2)  # respect 40 requests per minute limit

        df.to_excel(OUTPUT_FILE, index=False)
        print(f"Finished! Saved as {OUTPUT_FILE}")

    except Exception as e:
        print(f"Error: {e}")

def main_process(queue):
    try:
        queue.put({'type': 'status', 'text': 'Iniciando processo...'})
        df = pd.read_excel(EXCEL_FILE, sheet_name="Base")
        
        # Normalize numbers
        df["Longitude"] = df["Longitude"].astype(str).str.replace(",", ".").astype(float)
        df["Latitude"] = df["Latitude"].astype(str).str.replace(",", ".").astype(float)

        # ----------------------------------
        # BUILD ORIGINS FROM FILE (UNIQUE)
        # ----------------------------------
        origins = (
            df[["Origin", "Long|Lat"]]
            .drop_duplicates()
            .assign(coords=lambda x: x["Long|Lat"].apply(parse_lon_lat))
        )

        origin_map = dict(zip(origins["Origin"], origins["coords"]))

        df["distance_km"] = None

        # Count total chunks (for progress)
        total_chunks = 0
        for origin in origin_map:
            total_chunks += math.ceil(
                len(df[df["Origin"] == origin]) / CHUNK_SIZE
            )

        current_step = 0

        # ----------------------------------
        # PROCESS EACH ORIGIN DYNAMICALLY
        # ----------------------------------
        for origin_name, origin_coord in origin_map.items():
            subset = df[df["Origin"] == origin_name]

            indices = subset.index.tolist()
            destinations = subset[["Longitude", "Latitude"]].values.tolist()

            for i in range(0, len(destinations), CHUNK_SIZE):
                chunk_dest = destinations[i:i + CHUNK_SIZE]
                chunk_idx = indices[i:i + CHUNK_SIZE]

                locations = [list(origin_coord)] + chunk_dest

                payload = {
                    "locations": locations,
                    "metrics": ["distance"]
                }

                result = post_with_retry(payload)
                distances_m = result["distances"][0][1:]

                for idx, dist in zip(chunk_idx, distances_m):
                    if dist is not None:
                        df.at[idx, "distance_km"] = round(dist / 1000, 2)
                    else:
                        df.at[idx, "distance_km"] = None
                    queue.put({'type': 'log', 'text': f"{origin_name} - {df.at[idx, 'Destino']} - {df.at[idx, 'distance_km']} km"})

                current_step += 1
                progress = (current_step / total_chunks) * 100
                queue.put({'type': 'progress', 'value': progress})
                queue.put({'type': 'log', 'text': f"Processing {origin_name} ({current_step}/{total_chunks})"})

                # Save partial results to file
                df.to_excel(OUTPUT_FILE, index=False)
                queue.put({'type': 'log', 'text': f"Partial results saved to {OUTPUT_FILE}"})

                time.sleep(2)  # respect 40 requests per minute limit

        df.to_excel(OUTPUT_FILE, index=False)
        queue.put({'type': 'log', 'text': f"Finished! Saved as {OUTPUT_FILE}"})
        queue.put({'type': 'status', 'text': 'Processo concluído.'})
        queue.put({'type': 'done'})
    except Exception as e:
        queue.put({'type': 'log', 'text': f"Error: {e}"})
        if 'permission' in str(e).lower() or 'limit' in str(e).lower() or '403' in str(e):
            queue.put({'type': 'status', 'text': 'Limite atingido'})
        else:
            queue.put({'type': 'status', 'text': 'Erro ocorrido'})
        queue.put({'type': 'done'})

def update_gui(q, status_label, progress_bar, log_text, button, root):
    try:
        msg = q.get_nowait()
        if msg['type'] == 'status':
            status_label.config(text=msg['text'])
        elif msg['type'] == 'progress':
            progress_bar['value'] = msg['value']
        elif msg['type'] == 'log':
            log_text.insert(tk.END, msg['text'] + '\n')
            log_text.see(tk.END)
        elif msg['type'] == 'done':
            button.config(state="normal")
    except queue.Empty:
        pass
    root.after(100, update_gui, q, status_label, progress_bar, log_text, button, root)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("calcule a distância entre cidades - DHL")
        self.root.geometry("700x550")
        
        stellantis_blue = "#003DA5"
        stellantis_orange = "#FF6600"
        dhl_yellow = "#FFCC00"
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', background=stellantis_blue, foreground="white", relief="flat", padding=6)
        style.map('TButton', background=[('active', stellantis_orange)])
        style.configure('TProgressbar', background=stellantis_blue, troughcolor='#E8E8E8')
        style.configure('Title.TLabel', font=("Segoe UI", 16, "bold"), foreground=stellantis_blue)
        
        self.queue = queue.Queue()
        
        container = tk.Frame(root, bg="white")
        container.pack(fill=tk.BOTH, expand=True)
        
        header_frame = tk.Frame(container, bg=stellantis_blue, height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(header_frame, text="🤖 CALCULO DE DISTANCIA ENTRE CIDADES", font=("Segoe UI", 16, "bold"), fg="white", bg=stellantis_blue)
        title_label.pack(anchor="w", padx=15, pady=(10, 2))
        
        subtitle_label = tk.Label(header_frame, text="Processamento Inteligente de Processos Manuais", font=("Segoe UI", 9), fg=dhl_yellow, bg=stellantis_blue)
        subtitle_label.pack(anchor="w", padx=15, pady=(0, 10))
        
        main_frame = ttk.Frame(container, padding="13")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar.", font=("Segoe UI", 11), foreground=stellantis_blue)
        self.status_label.pack(pady=(2, 5), fill=tk.X)
        
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, fill=tk.X)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=4, fill=tk.X)
        
        self.process_button = ttk.Button(button_frame, text="▶ Processar", command=self.start_processing_thread)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Atividades", padding="13")
        log_frame.pack(pady=0, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=10, font=("Consolas", 11), bg="#F5F5F5", fg="#333333")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Footer section with DHL/STELLANTIS branding
        footer_frame = tk.Frame(container, bg=stellantis_blue, height=34)
        footer_frame.pack(fill=tk.X, padx=0, pady=0, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Left side - DHL -> PHILIPS
        left_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        left_footer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)
        
        # DHL Logo/Text (DHL Yellow)
        dhl_label = tk.Label(left_footer, text="🚚 DHL", font=("Segoe UI", 11, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        dhl_label.pack(side=tk.LEFT, padx=1)
        
        arrow_label = tk.Label(left_footer, text="→", font=("Segoe UI", 12, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        arrow_label.pack(side=tk.LEFT, padx=3)
        
        # PHILIPS Logo/Text (STELLANTIS Orange accent)
        philips_label = tk.Label(left_footer, text="PHILIPS 🏢", font=("Segoe UI", 11, "bold"), fg=stellantis_orange, bg=stellantis_blue)
        philips_label.pack(side=tk.LEFT, padx=3)
        
        # Right side - Developer credit
        right_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        right_footer.pack(side=tk.RIGHT, padx=15, pady=10)
        
        footer_label = tk.Label(right_footer, text="Desenvolvido por: Vincent Pernarh", font=("Segoe UI", 9), fg="white", bg=stellantis_blue)
        footer_label.pack(anchor="e")
        
    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.process_button, self.root)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
