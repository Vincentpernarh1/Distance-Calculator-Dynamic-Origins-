import os
import json
import time
import math
import threading
import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# ======================================================
# CONFIG
# ======================================================
EXCEL_FILE = os.path.join(os.getcwd(), "routes.xlsx")
OUTPUT_FILE = os.path.join(os.getcwd(), "routes_with_distance.xlsx")
MATRIX_URL = None

CHUNK_SIZE = 2000
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
    "Content-Type": "application/json"
}

# ======================================================
# HELPERS
# ======================================================
def parse_lon_lat(value):
    lon, lat = value.split("|")
    lon = float(lon.replace(",", "."))
    lat = float(lat.replace(",", "."))
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
        except Exception:
            if attempt == MAX_RETRIES:
                raise
            time.sleep(BACKOFF_FACTOR ** attempt)

# ======================================================
# MAIN PROCESS
# ======================================================
def process_file(progress_bar, status_label):
    try:
        df = pd.read_excel(EXCEL_FILE)

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

        # Count total chunks (for progress bar)
        total_chunks = 0
        for origin in origin_map:
            total_chunks += math.ceil(
                len(df[df["Origin"] == origin]) / CHUNK_SIZE
            )

        progress_bar["maximum"] = total_chunks
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
                    "sources": [0],
                    "destinations": list(range(1, len(locations))),
                    "metrics": ["distance"]
                }

                result = post_with_retry(payload)
                distances_m = result["distances"][0]

                for idx, dist in zip(chunk_idx, distances_m):
                    df.at[idx, "distance_km"] = round(dist / 1000, 2)

                current_step += 1
                progress_bar["value"] = current_step
                status_label.config(
                    text=f"Processing {origin_name} ({current_step}/{total_chunks})"
                )
                progress_bar.update_idletasks()

                time.sleep(1)  # respect free tier

        df.to_excel(OUTPUT_FILE, index=False)
        messagebox.showinfo("Done", f"Finished!\nSaved as {OUTPUT_FILE}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ======================================================
# TKINTER UI
# ======================================================
def start_process():
    threading.Thread(
        target=process_file,
        args=(progress_bar, status_label),
        daemon=True
    ).start()

root = tk.Tk()
root.title("OpenRouteService Distance Calculator")
root.geometry("540x230")

ttk.Label(
    root,
    text="Distance Calculator (Dynamic Origins)",
    font=("Arial", 14)
).pack(pady=10)

status_label = ttk.Label(root, text="Ready")
status_label.pack(pady=5)

progress_bar = ttk.Progressbar(root, length=450, mode="determinate")
progress_bar.pack(pady=10)

ttk.Button(
    root,
    text="Start Processing",
    command=start_process
).pack(pady=15)

root.mainloop()
