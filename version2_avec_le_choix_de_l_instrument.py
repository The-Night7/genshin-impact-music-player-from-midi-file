import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os

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

# Dictionnaire de mappage des notes MIDI vers les touches de Genshin Impact
# La lyre de Genshin utilise une gamme de Do majeur sur 3 octaves.
# C3 (48) à B5 (83)
GENSHIN_KEYMAP = {
    # Octave basse (C3 - B3)
    48: 'z', 50: 'x', 52: 'c', 53: 'v', 55: 'b', 57: 'n', 59: 'm',
    # Octave moyenne (C4 - B4) - C4 est le Do central (Note 60)
    60: 'a', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j',
    # Octave haute (C5 - B5)
    72: 'q', 74: 'w', 76: 'e', 77: 'r', 79: 't', 81: 'y', 83: 'u'
}


class GenshinMidiPlayer:
    def __init__(self, root):
        self.root = root
        self.root.title("Genshin Impact MIDI Player")
        self.root.geometry("400x350")
        self.root.resizable(False, False)

        self.midi_file_path = None
        self.is_playing = False
        self.play_thread = None

        self.create_widgets()

        # Raccourci d'arrêt d'urgence global
        keyboard.add_hotkey('F8', self.stop_playback)

    def create_widgets(self):
        # Titre
        title_label = tk.Label(self.root, text="Lyre Auto-Player", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # Sélection de l'instrument
        inst_frame = tk.Frame(self.root)
        inst_frame.pack(pady=5)
        tk.Label(inst_frame, text="Instrument :").pack(side=tk.LEFT)

        self.instrument_var = tk.StringVar(value="Windsong Lyre (3 Octaves)")
        self.inst_combo = ttk.Combobox(inst_frame, textvariable=self.instrument_var, state="readonly", width=30)
        self.inst_combo['values'] = (
            "Windsong Lyre (3 Octaves)",
            "Vintage Lyre (3 Octaves)",
            "Floral Zither (3 Octaves)",
            "Nightwing Horn (3 Octaves)",
            "Festive Drum (1 Octave)"
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
        # Wrapper pour s'assurer que l'appel depuis keyboard hotkey est thread-safe avec Tkinter
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

        # Lancement dans un thread séparé pour ne pas bloquer l'interface
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
        """Convertit une note MIDI en touche du clavier pour Genshin selon l'instrument."""
        inst = self.instrument_var.get()

        if "1 Octave" in inst:
            # Festive Drum (Aratambour) : on ramène toutes les notes sur l'octave du milieu (Do4 à Si4)
            note_dans_octave = midi_note % 12
            note_ramenee = 60 + note_dans_octave  # 60 = Do central

            mapping_1_octave = {60: 'a', 62: 's', 64: 'd', 65: 'f', 67: 'g', 69: 'h', 71: 'j'}

            if note_ramenee in mapping_1_octave:
                return mapping_1_octave[note_ramenee]
            # Si c'est un dièse/bémol, on prend la note naturelle juste en dessous
            elif note_ramenee - 1 in mapping_1_octave:
                return mapping_1_octave[note_ramenee - 1]
            return None

        else:
            # Instruments à 3 octaves (Windsong Lyre, Vintage Lyre, etc.)
            if midi_note in GENSHIN_KEYMAP:
                return GENSHIN_KEYMAP[midi_note]
            # Si c'est un dièse/bémol, on prend la note naturelle juste en dessous
            elif midi_note - 1 in GENSHIN_KEYMAP:
                return GENSHIN_KEYMAP[midi_note - 1]

        return None  # La note est en dehors de la gamme jouable (3 octaves)

    def playback_loop(self):
        # Compte à rebours pour laisser le temps de retourner sur la fenêtre du jeu
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

            # mid.play() gère automatiquement les timings et les tempos du fichier MIDI
            for msg in mid.play():
                if not self.is_playing:
                    break  # L'utilisateur a cliqué sur Stop ou appuyé sur F8

                # On ne traite que les messages de notes pressées
                if msg.type == 'note_on' and msg.velocity > 0:
                    adjusted_note = msg.note + transposition
                    key = self.get_genshin_key(adjusted_note)

                    if key:
                        # Simule une frappe de touche instantanée
                        keyboard.send(key)

            # Fin naturelle de la lecture
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

    # Force la fenêtre à rester au premier plan (optionnel mais pratique)
    root.attributes('-topmost', True)

    root.mainloop()