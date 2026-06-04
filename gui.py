import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading
import os
import sys
import time
from datetime import datetime
from PIL import Image, ImageTk
import io
from contextlib import redirect_stdout
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from process_data import charger_et_nettoyer
from mapper_data import run_test, export_per_scope
from extract_files_data import lancer_le_bilan
from calcul_data import calculer_emissions, agreger_par_scope, afficher_bilan, calculer_incertitude_bilan




class BilanCarboneGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("HDQUALITE Auto_Bilan_Carbone")
        self.root.geometry("1100x780")
        self.root.resizable(True, True)

        self.dossier_source  = None
        self.data_nettoyee   = None
        self.data_extraite   = None
        self.resultats       = None
        self.bilan           = None

        self.TRANSP = self.root.cget('bg')
        self.original_image = None
        self.setup_background("fond.jpg")
        self.setup_ui()
        self.root.bind("<Configure>", self.resize_background)

    def setup_background(self, image_path):
        try:
            self.original_image = Image.open(image_path)
            w = self.root.winfo_screenwidth()
            h = self.root.winfo_screenheight()
            img = self.original_image.resize((w, h), Image.Resampling.LANCZOS)
            self.bg_photo = ImageTk.PhotoImage(img)
            self.bg_label = tk.Label(self.root, image=self.bg_photo, bd=0)
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            self.bg_label.lower()
        except Exception as e:
            print(f"[FOND] Image non chargée : {e}")
            self.root.configure(bg="#1a1a2e")

    def resize_background(self, event):
        if self.original_image is None:
            return
        try:
            if event.width > 1 and event.height > 1:
                img = self.original_image.resize(
                    (event.width, event.height), Image.Resampling.LANCZOS)
                self.bg_photo = ImageTk.PhotoImage(img)
                self.bg_label.config(image=self.bg_photo)
        except Exception:
            pass


    def setup_ui(self):
        T = self.TRANSP

        self.main = tk.Frame(self.root, bg="#495899")
        self.main.place(relx=0.5, rely=0.5, anchor="center", width=1000, height=600)

        tk.Label(self.main, text="🌿  Bilan Carbone Automatisé",
                 font=("Arial", 22, "bold"), bg="#6841FF", fg="white",
                 padx=10, pady=5).pack(pady=(10, 5))

        tk.Label(self.main,
                 text="Outil de calcul des émissions GHG Protocol — ADEME",
                 font=("Arial", 10, "italic"), bg="#971515",
                 fg="#dfe6e9").pack(pady=(0, 15))

        btn_frame = tk.Frame(self.main, bg="#000874")
        btn_frame.pack(pady=5)

        btn_cfg = {"width": 26, "height": 2, "font": ("Arial", 10, "bold"),
                   "relief": "flat", "cursor": "hand2", "bd": 0}

        self.btn_choisir = tk.Button(
            btn_frame, text="📁  1. Choisir le dossier",
            command=self.choisir_dossier, bg="#2980b9", fg="white", **btn_cfg)
        self.btn_choisir.grid(row=0, column=0, padx=6, pady=4)

        self.btn_nettoyer = tk.Button(
            btn_frame, text="🧹  2. Nettoyer les fichiers",
            command=self.lancer_nettoyage, bg="#27ae60", fg="white",
            state=tk.DISABLED, **btn_cfg)
        self.btn_nettoyer.grid(row=0, column=1, padx=6, pady=4)

        self.btn_mapper = tk.Button(
            btn_frame, text="🗂️  3. Classifier par scope",
            command=self.lancer_mapping, bg="#e67e22", fg="white",
            state=tk.DISABLED, **btn_cfg)
        self.btn_mapper.grid(row=0, column=2, padx=6, pady=4)

        self.btn_calculer = tk.Button(
            btn_frame, text="⚡  4. Calculer les émissions",
            command=self.lancer_calcul, bg="#c0392b", fg="white",
            state=tk.DISABLED, **btn_cfg)
        self.btn_calculer.grid(row=0, column=3, padx=6, pady=4)

        self.label_dossier = tk.Label(self.main,
                                      text="Aucun dossier sélectionné",
                                      font=("Arial", 9), bg="#942328", fg="white")
        self.label_dossier.pack(pady=(5, 2))

        log_outer = tk.Frame(self.main, bg="white", bd=1, relief="solid")
        log_outer.pack(fill=tk.BOTH, expand=True, padx=5, pady=8)

        tk.Label(log_outer, text="Journal d'exécution",
                 font=("Arial", 9, "bold"), bg="white",
                 fg="white").pack(fill=tk.X, padx=8, pady=(5, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_outer, width=110, height=16, font=("Courier", 9),
            bg="#f8f9fa", fg="#000000", relief="flat", bd=0)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        res_frame = tk.Frame(self.main, bg="white", bd=1, relief="solid")
        res_frame.pack(fill=tk.X, padx=5, pady=(0, 5))

        self.result_label = tk.Label(res_frame, text="",
                                     font=("Arial", 18, "bold"),
                                     fg="#27ae60", bg="white")
        self.result_label.pack(pady=(8, 2))

        self.detail_label = tk.Label(res_frame, text="",
                                     font=("Arial", 10), fg="#555", bg="white")
        self.detail_label.pack(pady=(0, 8))

        tk.Label(self.main,
                 text=f"© {datetime.now().year}  HDQUALITE — Outil Bilan Carbone",
                 font=("Arial", 8), bg="#942238", fg="white").pack(pady=(5, 0))

    # LOGS

    def log(self, message, level="INFO"):
        colors = {"INFO": "#2d3436", "SUCCESS": "#00b894",
                  "WARNING": "#e17055", "ERROR": "#d63031"}
        timestamp = datetime.now().strftime("%H:%M:%S")
        tag = f"tag_{level}"
        self.log_text.tag_config(tag, foreground=colors.get(level, "#2d3436"))
        self.log_text.insert(tk.END, f"[{timestamp}] ", "gray")
        self.log_text.insert(tk.END, f"[{level}] ", tag)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.tag_config("gray", foreground="#b2bec3")
        self.log_text.see(tk.END)
        self.root.update()

    # ACTIONS 

    def choisir_dossier(self):
        d = filedialog.askdirectory(title="Sélectionner le dossier source")
        if d:
            self.dossier_source = d
            short = d if len(d) < 70 else "..." + d[-67:]
            self.label_dossier.config(text=f"📂  {short}", fg="#dfe6e9")
            self.btn_nettoyer.config(state=tk.NORMAL)
            self.log(f"Dossier sélectionné : {d}")

    def lancer_nettoyage(self):
        self.btn_nettoyer.config(state=tk.DISABLED)
        self.log("═" * 60)
        self.log("ÉTAPE 1 — Nettoyage des fichiers", "INFO")

        def run():
            try:
                os.chdir(self.dossier_source)

                t0 = time.time()
                self.data_nettoyee, errors = charger_et_nettoyer(self.dossier_source)
                duree = time.time() - t0

                self.log(f"Nettoyage terminé en {duree:.1f}s — "
                         f"{len(self.data_nettoyee)} fichiers traités", "SUCCESS")
                if errors:
                    self.log(f"{len(errors)} erreur(s) :", "WARNING")
                    for err in errors[:5]:
                        self.log(f"  ↳ {err}", "WARNING")

                self.btn_mapper.config(state=tk.NORMAL)
            except Exception as e:
                self.log(f"Erreur critique : {e}", "ERROR")
                self.btn_nettoyer.config(state=tk.NORMAL)

        threading.Thread(target=run, daemon=True).start()

    def lancer_mapping(self):
        self.btn_mapper.config(state=tk.DISABLED)
        self.log("═" * 60)
        self.log("ÉTAPE 2 — Classification par scope", "INFO")

        def run():
            try:
                os.chdir(self.dossier_source)

                cleaned_folder = os.path.join(self.dossier_source, "cleaned_files")

                t0 = time.time()
                results = run_test(folder=cleaned_folder)
                export_per_scope(results,
                                 source_folder=cleaned_folder,
                                 dest_base=self.dossier_source)
                duree = time.time() - t0

                counts = {}
                for scope in results.values():
                    counts[scope] = counts.get(scope, 0) + 1

                self.log(f"Classification terminée en {duree:.1f}s", "SUCCESS")
                for scope, n in sorted(counts.items()):
                    self.log(f"  ↳ {scope} : {n} fichier(s)", "INFO")
                self.log(f"Dossiers SCOPE créés dans : {self.dossier_source}", "INFO")

                self.btn_calculer.config(state=tk.NORMAL)
            except Exception as e:
                self.log(f"Erreur critique : {e}", "ERROR")
                self.btn_mapper.config(state=tk.NORMAL)

        threading.Thread(target=run, daemon=True).start()

    def lancer_calcul(self):
        self.btn_calculer.config(state=tk.DISABLED)
        self.log("═" * 60)
        self.log("ÉTAPE 3 — Extraction & calcul des émissions", "INFO")

        def run():
            try:
                os.chdir(self.dossier_source)

                t0 = time.time()

                self.log("Extraction des données en cours...", "INFO")
                self.data_extraite, _ = lancer_le_bilan(base_path=self.dossier_source)
                self.log(f"{len(self.data_extraite)} valeur(s) extraite(s)", "SUCCESS")

                self.log("Calcul des émissions CO₂...", "INFO")
                self.resultats, non_calc = calculer_emissions(self.data_extraite)
                self.bilan = agreger_par_scope(self.resultats)
                incertitude = calculer_incertitude_bilan(self.bilan, self.data_extraite)
                duree = time.time() - t0

                tampon = io.StringIO()
                with redirect_stdout(tampon):
                    afficher_bilan(self.bilan, incertitude)
                self.log(tampon.getvalue(), "INFO")

                total = self.bilan['TOTAL']['emissions_tCO2e']
                self.log(f"Calcul terminé en {duree:.1f}s", "SUCCESS")
                self.log(f"TOTAL : {total:.4f} tCO₂e", "SUCCESS")

                if non_calc:
                    self.log(f"{len(non_calc)} valeur(s) non calculée(s) :", "WARNING")
                    for nc in non_calc:
                        self.log(f"  ↳ {nc['designation']} — {nc['raison']}", "WARNING")

                self.result_label.config(text=f"Total : {total:.4f} tCO₂e")
                details = "   |   ".join(
                    f"{scope} : {data['emissions_tCO2e']:.4f} tCO₂e"
                    for scope, data in self.bilan.items()
                    if scope != "TOTAL"
                )
                self.detail_label.config(text=details)

            except Exception as e:
                self.log(f"Erreur critique : {e}", "ERROR")
            finally:
                self.btn_calculer.config(state=tk.NORMAL)

        threading.Thread(target=run, daemon=True).start()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = BilanCarboneGUI()
    app.run()