import customtkinter as ctk
from tkcalendar import Calendar
from PIL import Image, ImageTk
import os
from datetime import datetime
from logic import (
    ensure_full_path, 
    get_excel_path, 
    list_photos_to_process, 
    process_photo,
    undo_process,
    search_brinco,
    FARM_MAPPING
)
from tkinter import messagebox
from pathlib import Path
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return Path(base_path) / relative_path


# Optional HEIC/HEIF support via pillow-heif
try:
    from pillow_heif import register_heif_opener

    register_heif_opener()
    HEIF_ENABLED = True
except Exception:
    HEIF_ENABLED = False

# Farm Palette
COLOR_FOREST_GREEN = "#2E7D32"
COLOR_EARTH_BROWN = "#5D4037"
COLOR_CREAM = "#D7CCC8"
COLOR_BG = "#1A1A1A" # Slightly darker for more contrast
COLOR_ACCENT = "#8D6E63"

class RPAFazendaApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("RPA Foto Fazenda v3.0")
        self.geometry("1000x800")
        ctk.set_appearance_mode("Dark")
        
        # Set Icon
        icon_path = resource_path("assets/ox_icon.png")
        if icon_path.exists():
            try:
                img = Image.open(icon_path)
                photo = ImageTk.PhotoImage(img)
                self.after(200, lambda: self.wm_iconphoto(True, photo))
            except Exception as e:
                print(f"Error loading icon: {e}")

        
        self.selected_date = None
        self.selected_farm = None
        self.selected_batch = None
        
        self.photos_to_process = []
        self.processed_history = []
        self.current_photo_index = 0
        self.total_photos = 0
        
        # UI Container
        self.main_container = ctk.CTkFrame(self, fg_color=COLOR_BG)
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.show_dashboard()

    def clear_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()

    def show_dashboard(self):
        self.clear_container()
        
        # Title with theme
        title_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        title_frame.pack(pady=(50, 40))
        
        ctk.CTkLabel(
            title_frame, 
            text="Fazenda Digital", 
            font=("Roboto", 48, "bold"),
            text_color=COLOR_FOREST_GREEN
        ).pack()
        
        ctk.CTkLabel(
            title_frame, 
            text="Sistema de Registro e Gerenciamento", 
            font=("Roboto", 18),
            text_color=COLOR_CREAM
        ).pack(pady=5)
        
        # Dashboard Buttons
        btn_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        btn_frame.pack(pady=20, expand=True)
        
        process_btn = ctk.CTkButton(
            btn_frame, 
            text="📥 Processar Fotos", 
            command=self.show_setup_screen,
            width=300,
            height=80,
            font=("Roboto", 24, "bold"),
            fg_color=COLOR_FOREST_GREEN,
            hover_color="#1B5E20"
        )
        process_btn.grid(row=0, column=0, padx=20, pady=20)
        
        search_btn = ctk.CTkButton(
            btn_frame, 
            text="🔍 Buscar Brinco", 
            command=self.show_search_screen,
            width=300,
            height=80,
            font=("Roboto", 24, "bold"),
            fg_color=COLOR_EARTH_BROWN,
            hover_color="#3E2723"
        )
        search_btn.grid(row=0, column=1, padx=20, pady=20)

    # --- SETUP SCREEN ---
    def show_setup_screen(self):
        self.clear_container()
        
        # Back to Dashboard
        ctk.CTkButton(self.main_container, text="← Início", command=self.show_dashboard, width=80, fg_color=COLOR_EARTH_BROWN).pack(anchor="nw", padx=10, pady=10)

        label = ctk.CTkLabel(self.main_container, text="Configuração do Lote", font=("Roboto", 28, "bold"), text_color=COLOR_CREAM)
        label.pack(pady=(10, 20))
        
        setup_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        setup_frame.pack(pady=10, fill="both", expand=True)
        
        # Left: Calendar
        left_side = ctk.CTkFrame(setup_frame, fg_color="transparent")
        left_side.pack(side="left", padx=30, expand=True)
        
        ctk.CTkLabel(left_side, text="1. Selecione a Data:", font=("Roboto", 16, "bold"), text_color=COLOR_CREAM).pack(pady=5)
        self.cal = Calendar(left_side, selectmode='day', background=COLOR_EARTH_BROWN, foreground="white", headersbackground=COLOR_FOREST_GREEN)
        self.cal.pack(pady=10)
        
        # Right: Info
        right_side = ctk.CTkFrame(setup_frame, fg_color="transparent")
        right_side.pack(side="left", padx=30, expand=True)
        
        ctk.CTkLabel(right_side, text="2. Escolha a Fazenda:", font=("Roboto", 16, "bold"), text_color=COLOR_CREAM).pack(pady=(0, 5))
        self.farm_var = ctk.StringVar(value="VAL")
        self.farm_menu = ctk.CTkOptionMenu(right_side, values=["VAL", "BARRA", "TRES"], variable=self.farm_var, fg_color=COLOR_FOREST_GREEN, button_color=COLOR_EARTH_BROWN)
        self.farm_menu.pack(pady=(0, 20), fill="x")
        
        ctk.CTkLabel(right_side, text="3. Digite o Lote:", font=("Roboto", 16, "bold"), text_color=COLOR_CREAM).pack(pady=(0, 5))
        self.batch_entry = ctk.CTkEntry(right_side, placeholder_text="Ex: LOTE 01", font=("Roboto", 14))
        self.batch_entry.pack(pady=(0, 20), fill="x")
        
        confirm_btn = ctk.CTkButton(self.main_container, text="Abrir Pasta de Fotos", command=self.start_processing, fg_color=COLOR_FOREST_GREEN, height=50, font=("Roboto", 18, "bold"))
        confirm_btn.pack(pady=30)

    # --- SEARCH SCREEN ---
    def show_search_screen(self):
        self.clear_container()
        
        ctk.CTkButton(self.main_container, text="← Início", command=self.show_dashboard, width=80, fg_color=COLOR_EARTH_BROWN).pack(anchor="nw", padx=10, pady=10)

        ctk.CTkLabel(self.main_container, text="Pesquisa de Animais", font=("Roboto", 28, "bold"), text_color=COLOR_CREAM).pack(pady=10)
        
        search_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        search_frame.pack(pady=20, fill="x", padx=100)
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="Digite o número do brinco...", height=40, font=("Roboto", 16))
        self.search_entry.pack(side="left", expand=True, fill="x", padx=10)
        self.search_entry.bind("<Return>", lambda e: self.perform_search())
        
        search_btn = ctk.CTkButton(search_frame, text="🔍 Buscar", command=self.perform_search, width=100, height=40, font=("Roboto", 14, "bold"), fg_color=COLOR_FOREST_GREEN)
        search_btn.pack(side="left", padx=10)
        
        # Results area
        self.results_box = ctk.CTkTextbox(self.main_container, font=("Roboto", 14), fg_color="#242424")
        self.results_box.pack(pady=10, fill="both", expand=True, padx=50)
        self.results_box.configure(state="disabled")

    def perform_search(self):
        query = self.search_entry.get().strip()
        if not query: return
        
        exact, near = search_brinco(query)
        
        self.results_box.configure(state="normal")
        self.results_box.delete("1.0", "end")
        
        if not exact and not near:
            self.results_box.insert("end", "❌ Nenhum animal encontrado com este número.\n\n")
        
        if exact:
            self.results_box.insert("end", "✅ RESULTADO EXATO:\n", "bold")
            for m in exact:
                self.results_box.insert("end", f"   • {m['brinco']} -> Fazenda: {m['farm']} | Lote: {m['batch']} | Dia: {m['date']}\n")
            self.results_box.insert("end", "\n")
            
        if near:
            self.results_box.insert("end", "❓ RESULTADOS SIMILARES (1 dígito diferente):\n", "bold")
            for m in near:
                # Group by brinco to avoid clutter if needed, but here simple list
                self.results_box.insert("end", f"   • {m['brinco']} -> Fazenda: {m['farm']} | Lote: {m['batch']} | Dia: {m['date']}\n")
        
        self.results_box.configure(state="disabled")

    # --- PROCESSING FLOW ---
    def start_processing(self):
        self.selected_date = self.cal.get_date()
        try:
            date_obj = datetime.strptime(self.selected_date, '%m/%d/%y')
        except: # Fallback for different locales
            date_obj = datetime.now()
        self.date_str = date_obj.strftime("%d-%m-%Y")
        
        self.selected_farm = self.farm_var.get()
        self.selected_batch = self.batch_entry.get().strip()
        
        if not self.selected_batch:
            messagebox.showwarning("Atenção", "Por favor, informe o lote.")
            return
            
        self.target_path = ensure_full_path(self.selected_farm, self.selected_batch, self.date_str)
        self.excel_path = get_excel_path(self.target_path, self.date_str)
        
        self.photos_to_process = list_photos_to_process()
        self.total_photos = len(self.photos_to_process)
        self.current_photo_index = 0
        self.processed_history = []

        heic_present = any(Path(p).suffix.lower() in (".heic", ".heif") for p in self.photos_to_process)
        if heic_present and not HEIF_ENABLED:
            messagebox.showerror(
                "Erro",
                "Fotos HEIC/HEIF encontradas, mas o suporte nao esta instalado. Instale pillow-heif."
            )
            return
        
        if self.total_photos == 0:
            messagebox.showinfo("Aviso", "Sem fotos na pasta 'processar'")
            return
        
        self.show_processing_screen()

    def show_processing_screen(self):
        self.clear_container()
        
        if self.current_photo_index >= self.total_photos:
            self.show_final_screen()
            return
            
        current_photo_path = self.photos_to_process[self.current_photo_index]
        
        ctk.CTkLabel(self.main_container, text=f"Fazenda: {self.selected_farm} | Lote: {self.selected_batch}", font=("Roboto", 14), text_color="#AAAAAA").pack(pady=(0, 5))
        
        self.counter_label = ctk.CTkLabel(self.main_container, text=f"Foto {self.current_photo_index + 1} / {self.total_photos}", font=("Roboto", 24, "bold"), text_color=COLOR_CREAM)
        self.counter_label.pack(pady=(5, 5))
        
        self.img_frame = ctk.CTkFrame(self.main_container, fg_color=COLOR_EARTH_BROWN)
        self.img_frame.pack(pady=10, fill="both", expand=True)
        
        try:
            img = Image.open(current_photo_path)
            img.thumbnail((800, 450))
            self.photo_img = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
            self.img_label = ctk.CTkLabel(self.img_frame, image=self.photo_img, text="")
            self.img_label.pack(expand=True)
        except Exception as e:
            ctk.CTkLabel(self.img_frame, text=f"Erro ao carregar imagem: {e}").pack(expand=True)
            
        input_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        input_frame.pack(pady=20)
        
        ctk.CTkLabel(input_frame, text="Brinco:", font=("Roboto", 22, "bold"), text_color=COLOR_CREAM).pack(side="left", padx=10)
        
        self.brinco_entry = ctk.CTkEntry(input_frame, width=200, font=("Roboto", 22))
        self.brinco_entry.pack(side="left", padx=10)
        self.brinco_entry.bind("<Return>", lambda e: self.process_current())
        self.brinco_entry.focus_set()

        btn_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        if self.processed_history:
            ctk.CTkButton(btn_frame, text=" ← Voltar", command=self.undo_last, fg_color="#8B0000", font=("Roboto", 14), width=120).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Próximo →", command=self.process_current, fg_color=COLOR_FOREST_GREEN, font=("Roboto", 14), width=120).pack(side="left", padx=10)

    def process_current(self):
        brinco = self.brinco_entry.get().strip()
        if not brinco: return
            
        orig_path = self.photos_to_process[self.current_photo_index]
        dest_path = self.target_path / f"{brinco}{Path(orig_path).suffix}"
        
        success, message = process_photo(orig_path, self.target_path, brinco, self.excel_path)
        
        if success:
            self.processed_history.append({'original_path': orig_path, 'dest_path': str(dest_path), 'brinco_number': brinco, 'excel_path': str(self.excel_path)})
            self.current_photo_index += 1
            self.show_processing_screen()
        else:
            messagebox.showerror("Erro", message)
            self.brinco_entry.delete(0, 'end')

    def undo_last(self):
        if not self.processed_history: return
        last_item = self.processed_history.pop()
        success, message = undo_process(last_item)
        if success:
            self.current_photo_index -= 1
            self.show_processing_screen()
            self.brinco_entry.insert(0, last_item['brinco_number'])
            self.brinco_entry.select_range(0, 'end')
        else:
            messagebox.showerror("Erro", message)
            self.processed_history.append(last_item)

    def show_final_screen(self):
        self.clear_container()
        ctk.CTkLabel(self.main_container, text="Processamento Concluído!", font=("Roboto", 32, "bold"), text_color=COLOR_FOREST_GREEN).pack(pady=(80, 20))
        ctk.CTkButton(self.main_container, text="Sair", command=self.destroy, fg_color=COLOR_FOREST_GREEN, width=200).pack(pady=20)
        ctk.CTkButton(self.main_container, text="Início", command=self.show_dashboard, fg_color="transparent", border_width=2, border_color=COLOR_EARTH_BROWN).pack(pady=10)

if __name__ == "__main__":
    app = RPAFazendaApp()
    app.mainloop()
