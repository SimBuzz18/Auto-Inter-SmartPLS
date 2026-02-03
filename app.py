
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os
import json
import warnings
from smartpls_logic import SmartPLSReader

# Suppress openpyxl style warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# Configuration
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("SmartPLS Interpreter")
        self.geometry("600x400")
        
        # Set Icon
        try:
            self.iconbitmap(r"c:\Users\Admin\OneDrive\Dokumen\Asyraf\CODING\Inter-SmartPLS\icon.ico")
        except:
            pass
            
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header
        self.label_title = ctk.CTkLabel(self, text="Interpretasi Hasil SmartPLS", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Input Frame
        self.frame_input = ctk.CTkFrame(self)
        self.frame_input.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.frame_input.grid_columnconfigure(0, weight=1)

        self.entry_path = ctk.CTkEntry(self.frame_input, placeholder_text="Pilih folder hasil export...")
        
        # Load config
        last_path = self.load_config()
        if last_path and os.path.exists(last_path):
            self.entry_path.insert(0, last_path)
        else:
            self.entry_path.insert(0, r"C:\Users\Admin\OneDrive\Dokumen\Asyraf\CODING\Inter-SmartPLS")
            
        self.entry_path.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="ew")

        self.btn_browse = ctk.CTkButton(self.frame_input, text="Browse", width=80, command=self.browse_folder)
        self.btn_browse.grid(row=0, column=1, padx=(5, 10), pady=10)

        # Action Button & Log
        self.btn_process = ctk.CTkButton(self, text="Proses Interpretasi", command=self.start_process, font=ctk.CTkFont(size=16))
        self.btn_process.grid(row=2, column=0, padx=20, pady=10, sticky="n")
        
        self.textbox_log = ctk.CTkTextbox(self, width=500, height=200)
        self.textbox_log.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.textbox_log.insert("0.0", "Siap. Silakan pilih folder dan klik Proses.\n")

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, folder_selected)
            self.log(f"Selected: {folder_selected}")
            self.save_config(folder_selected)

    def log(self, message):
        self.textbox_log.insert(tk.END, message + "\n")
        self.textbox_log.see(tk.END)

    def start_process(self):
        folder = self.entry_path.get()
        if folder:
            self.save_config(folder)
            
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Silakan pilih folder yang valid.")
            return

        self.btn_process.configure(state="disabled")
        self.log(f"\nMemulai proses di: {folder}...")
        
        # Threading to prevent UI freeze
        thread = threading.Thread(target=self.run_logic, args=(folder,))
        thread.start()

    def run_logic(self, folder):
        try:
            reader = SmartPLSReader(folder)
            out_path = reader.process()
            self.log(f"\nBerhasil! Output disimpan di:\n{out_path}")
            # Optional: Open file
            # os.startfile(out_path) 
        except Exception as e:
            self.log(f"Error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.btn_process.configure(state="normal")

    def load_config(self):
        """Load last used path from config.json"""
        config_path = "config.json"
        if os.path.exists(config_path):
            try:
                with open(config_path, "r") as f:
                    data = json.load(f)
                    return data.get("last_path", "")
            except:
                pass
        return ""

    def save_config(self, path):
        """Save path to config.json"""
        config_path = "config.json"
        try:
            with open(config_path, "w") as f:
                json.dump({"last_path": path}, f)
        except Exception as e:
            print(f"Failed to save config: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
