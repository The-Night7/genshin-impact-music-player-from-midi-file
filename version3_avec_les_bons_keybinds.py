import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os
import ctypes

try:
    import mido
    import keyboard
except ImportError:
    import tkinter.messagebox as msg

    root = tk.Tk()
    root.withdraw()
    msg.showerror("Erreur de dépendance",
                  "Veuillez installer les modules requis en exécutant:\npip install mido keyboard")
    exit()

# Dictionnaires de mappage des notes MIDI selon le type de clavier et d'instrument
KEYMAPS = {
    "QWERTY": {
        "3_octaves": {
            # Basses (Z X C V B N M)
            48: 'z', 50: 'x', 52: 'c', 53: 'v', 55: 'b', 57: 'n', 59: 'm',
            # Moyennes (A S D F G H J)
            60: 'a', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j',
            # Hautes (Q W E R T Y U)
            72: 'q', 74: 'w', 76: 'e', 77: 'r', 79: 't', 81: 'y', 83: 'u'
        },
        "2_octaves": {
            # Moyennes (A S D F G H J)
            60: 'a', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j',
            # Hautes (Q W E R T Y U)
            72: 'q', 74: 'w', 76: 'e', 77: 'r', 79: 't', 81: 'y', 83: 'u'
        },
        "drum": ['a', 's', 'k', 'l']  # KA(L), DON(L), DON(R), KA(R)
    },
    "AZERTY": {
        "3_octaves": {
            # Basses (W X C V B N M)
            48: 'w', 50: 'x', 52: 'c', 53: 'v', 55: 'b', 57: 'n', 59: 'm',
            # Moyennes (Q S D F G H J)
            60: 'q', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j',
            # Hautes (A Z E R T Y U)
            72: 'a', 74: 'z', 76: 'e', 77: 'r', 79: 't', 81: 'y', 83: 'u'
        },
        "2_octaves": {
            # Moyennes (Q S D F G H J)
            60: 'q', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j',
            # Hautes (A Z E R T Y U)
            72: 'a', 74: 'z', 76: 'e', 77: 'r', 79: 't', 81: 'y', 83: 'u'
        },
        "drum": ['q', 's', 'k', 'l']  # KA(L), DON(L), DON(R), KA(R)
    }
}


class GenshinMidiPlayer:
    def __init__(self, root):
        self.root = root
        self.root.title("Genshin Impact MIDI Player")
        self.root.geometry("400x420")  # Agrandissement de la fenêtre pour le nouveau menu
        self.root.resizable(False, False)

        self.midi_file_path = None
        self.is_playing = False
        self.play_thread = None

        self.create_widgets()
        self.check_admin()

        # Raccourci d'arrêt d'urgence global
        keyboard.add_hotkey('F8', self.stop_playback)

    def check_admin(self):
        """Vérifie si le script est lancé en tant qu'administrateur et prévient l'utilisateur."""
        if os.name == 'nt':
            try:
                if not ctypes.windll.shell32.IsUserAnAdmin():
                    messagebox.showwarning(
                        "Droits Administrateur Requis",
                        "Genshin Impact bloque les macros si le logiciel n'est pas lancé en tant qu'Administrateur.\n\n"
                        "Si la musique ne se joue pas en jeu, fermez ce programme et relancez votre invite de commande/IDE ou le script avec un clic droit -> 'Exécuter en tant qu'administrateur'."
                    )
            except:
                pass

    def create_widgets(self):
        # Titre
        title_label = tk.Label(self.root, text="Lyre Auto-Player", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # Sélection du clavier
        layout_frame = tk.Frame(self.root)
        layout_frame.pack(pady=5)
        tk.Label(layout_frame, text="Clavier :").pack(side=tk.LEFT)

        self.layout_var = tk.StringVar(value="AZERTY")
        self.layout_combo = ttk.Combobox(layout_frame, textvariable=self.layout_var, state="readonly", width=20)
        self.layout_combo['values'] = ("AZERTY", "QWERTY")
        self.layout_combo.pack(side=tk.LEFT, padx=5)

        # Sélection de l'instrument
        inst_frame = tk.Frame(self.root)
        inst_frame.pack(pady=5)
        tk.Label(inst_frame, text="Instrument :").pack(side=tk.LEFT)

        self.instrument_var = tk.StringVar(value="Windsong Lyre (3 Octaves)")
        self.inst_combo = ttk.Combobox(inst_frame, textvariable=self.instrument_var, state="readonly", width=25)
        self.inst_combo['values'] = (
            "Windsong Lyre (3 Octaves)",
            "Vintage Lyre (3 Octaves)",
            "Floral Zither (3 Octaves)",
            "Nightwing Horn (2 Octaves)",
            "Festive Drum (4 Touches)"
        )
        self.inst_combo.pack(side=tk.LEFT, padx=5)

        # Sélection du fichier
        self.file_label = tk.Label(self.root, text="Aucun fichier sélectionné", fg="gray", wraplength=380)
        self.file_label.pack(pady=5)

        browse_btn = tk.Button(self.root, text="Parcourir (.mid)", command=self.browse_file, width=20)
        browse_btn.pack(pady=5)

        # Transposition
        trans_frame = tk.Frame(self.root)
        trans_frame.pack(pady=10)
        tk.Label(trans_frame, text="Transposition (demi-tons) :").pack(side=tk.LEFT)

        self.transpose_var = tk.IntVar(value=0)
        trans_spin = tk.Spinbox(trans_frame, from_=-36, to=36, textvariable=self.transpose_var, width=5)
        trans_spin.pack(side=tk.LEFT, padx=5)

        # Boutons de contrôle
        ctrl_frame = tk.Frame(self.root)
        ctrl_frame.pack(pady=15)

        self.play_btn = tk.Button(ctrl_frame, text="Jouer (F5)", command=self.start_playback, bg="#4CAF50", fg="white",
                                  width=12, font=("Helvetica", 10, "bold"))
        self.play_btn.pack(side=tk.LEFT, padx=10)
        # Raccourci pour lancer
        keyboard.add_hotkey('F5', self.start_playback_hotkey)

        self.stop_btn = tk.Button(ctrl_frame, text="Stop (F8)", command=self.stop_playback, state=tk.DISABLED,
                                  bg="#F44336", fg="white", width=12, font=("Helvetica", 10, "bold"))
        self.stop_btn.pack(side=tk.LEFT, padx=10)

        # Statut
        self.status_label = tk.Label(self.root, text="Prêt.", font=("Helvetica", 10))
        self.status_label.pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers MIDI", "*.mid *.midi")])
        if file_path:
            self.midi_file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=filename, fg="black")

    def start_playback_hotkey(self):
        # Wrapper pour s'assurer que l'appel depuis keyboard hotkey est thread-safe
        if not self.is_playing:
            self.root.after(0, self.start_playback)

    def start_playback(self):
        if not self.midi_file_path:
            messagebox.showwarning("Attention", "Veuillez d'abord sélectionner un fichier MIDI.")
            return

        if self.is_playing:
            return

        self.is_playing = True
        self.play_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        self.play_thread = threading.Thread(target=self.playback_loop)
        self.play_thread.daemon = True
        self.play_thread.start()

    def stop_playback(self):
        if self.is_playing:
            self.is_playing = False
            self.root.after(0, self.update_ui_stopped)

    def update_ui_stopped(self):
        self.play_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Lecture arrêtée.", fg="red")

    def get_genshin_key(self, midi_note):
        """Convertit une note MIDI en touche du clavier en fonction du layout et de l'instrument."""
        layout = self.layout_var.get()
        inst = self.instrument_var.get()

        # Déterminer si on utilise la map Drum, 2 octaves ou 3 octaves
        if "Drum" in inst:
            # Le Festive Drum utilise 4 touches (KA L, DON L, DON R, KA R)
            # On répartit cycliquement toutes les notes MIDI sur ces 4 touches
            drum_keys = KEYMAPS[layout]["drum"]
            return drum_keys[midi_note % 4]
        elif "2 Octaves" in inst:
            # On ramène les notes hors de portée dans la plage 60-83 (Do4 à Si5)
            note_a_jouer = midi_note
            while note_a_jouer < 60:
                note_a_jouer += 12
            while note_a_jouer > 83:
                note_a_jouer -= 12
            current_map = KEYMAPS[layout]["2_octaves"]
        else:
            note_a_jouer = midi_note
            current_map = KEYMAPS[layout]["3_octaves"]

        # Cherche la note exacte (Pour les instruments normaux)
        if note_a_jouer in current_map:
            return current_map[note_a_jouer]
        # Si c'est un dièse/bémol, on prend la note naturelle juste en dessous
        elif note_a_jouer - 1 in current_map:
            return current_map[note_a_jouer - 1]

        return None

    def press_key_for_game(self, key):
        """Maintient la touche enfoncée 50ms pour que le jeu puisse la lire, puis la relâche."""
        keyboard.press(key)
        time.sleep(0.05)
        keyboard.release(key)

    def playback_loop(self):
        # Compte à rebours
        for i in range(3, 0, -1):
            if not self.is_playing: return
            self.root.after(0, self.status_label.config,
                            {"text": f"La lecture commence dans {i} secondes... Allez en jeu !", "fg": "orange"})
            time.sleep(1)

        if not self.is_playing: return
        self.root.after(0, self.status_label.config,
                        {"text": "Lecture en cours... (Appuyez sur F8 pour arrêter)", "fg": "green"})

        try:
            mid = mido.MidiFile(self.midi_file_path)
            transposition = self.transpose_var.get()

            for msg in mid.play():
                if not self.is_playing:
                    break

                if msg.type == 'note_on' and msg.velocity > 0:
                    adjusted_note = msg.note + transposition
                    key = self.get_genshin_key(adjusted_note)

                    if key:
                        # Exécuté dans un thread pour ne pas désynchroniser le rythme de la musique MIDI
                        threading.Thread(target=self.press_key_for_game, args=(key,), daemon=True).start()

            if self.is_playing:
                self.is_playing = False
                self.root.after(0, lambda: self.status_label.config(text="Lecture terminée.", fg="blue"))
                self.root.after(0, self.update_ui_stopped)

        except Exception as e:
            self.is_playing = False
            self.root.after(0, self.update_ui_stopped)
            self.root.after(0,
                            lambda: messagebox.showerror("Erreur de lecture", f"Une erreur est survenue :\n{str(e)}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = GenshinMidiPlayer(root)
    root.attributes('-topmost', True)
    root.mainloop()