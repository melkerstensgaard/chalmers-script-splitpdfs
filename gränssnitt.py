import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import queue

from bearbetningsskript import process_all, get_default_config
from valideringsskript import validate_output
# -------------------------
# LIST EDITOR COMPONENT
# -------------------------

class ListEditor(tk.Frame):
    def __init__(self, parent, title):
        super().__init__(parent)

        tk.Label(self, text=title, font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.listbox = tk.Listbox(self, height=4) # Detta är längden på text-input rutorna
        self.listbox.pack(fill="both", expand=True, padx=2, pady=2)

        entry_frame = tk.Frame(self)
        entry_frame.pack(fill="x")

        self.entry = tk.Entry(entry_frame)
        self.entry.pack(side="left", fill="x", expand=True)

        tk.Button(entry_frame, text="+", command=self.add_item).pack(side="left")
        tk.Button(entry_frame, text="-", command=self.remove_selected).pack(side="left")

    def add_item(self):
        value = self.entry.get().strip()
        if value:
            self.listbox.insert(tk.END, value)
            self.entry.delete(0, tk.END)

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)

    def set_values(self, values):
        self.listbox.delete(0, tk.END)
        for v in values:
            self.listbox.insert(tk.END, v)

    def get_values(self):
        return list(self.listbox.get(0, tk.END))


# -------------------------
# MAIN GUI APP
# -------------------------

class CertificateGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gränssnitt för bearbetningsskript")
        self.root.geometry("550x800")
        # Progress bar
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()

        self.config = get_default_config()

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()


        self.skip_unreadable = tk.BooleanVar(value=True)
        self.require_second_page = tk.BooleanVar(value=True)

        self.allow_personnummer = tk.BooleanVar(
            value=False
        )
        # Validation
        self.run_validation = tk.BooleanVar(
            value=True
        )
        self.build_ui()
        self.load_config_into_ui()

        self.poll_queues()



    # -------------------------
    # UI LAYOUT
    # -------------------------

    def build_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(fill="x", padx=10, pady=10)

        # INPUT
        tk.Label(frame, text="Input").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.input_path, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame, text="Bläddra", command=self.select_input).grid(row=0, column=2)

        # OUTPUT
        tk.Label(frame, text="Output").grid(row=1, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.output_path, width=60).grid(row=1, column=1, padx=5)
        tk.Button(frame, text="Bläddra", command=self.select_output).grid(row=1, column=2)


        # Checkbox controls
        
        tk.Checkbutton(
            self.root,
            text="Tillåt personnummer på första sidan",
            variable=self.allow_personnummer
        ).pack(anchor="w", padx=10)

        # Validation button
        tk.Checkbutton(
            self.root,
            text="Kör slutvalidering",
            variable=self.run_validation
        ).pack(anchor="w", padx=10)

        '''
        tk.Checkbutton(
            self.root,
            text="Hoppa över oläsbara volymer",
            variable=self.skip_unreadable
        ).pack(anchor="w", padx=10)

        tk.Checkbutton(
            self.root,
            text="Kräv bekräftelse från sida 2",
            variable=self.require_second_page
        ).pack(anchor="w", padx=10)
        '''


        # CONFIG EDITORS
        self.degree_editor = ListEditor(self.root, "Markör för bevisets första sida") #(START_PAGE_DEGREE_MARKERS)
        self.degree_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.award_editor = ListEditor(self.root, "Stark indikation att beviset är en första sida") #START_PAGE_STRONG_AWARD_MARKERS
        self.award_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.special_editor = ListEditor(self.root, "Markör för undantagsbevis") #SPECIAL_START_PAGE_IDENTIFIERS
        self.special_editor.pack(fill="both", expand=True, padx=10, pady=5)

        self.nonstart_editor = ListEditor(self.root, "Markör för att inte vara en första sida") #NON_START_PAGE_PATTERNS
        self.nonstart_editor.pack(fill="both", expand=True, padx=10, pady=5)

        # RUN BUTTON
        tk.Button(
            self.root,
            text="Starta skriptet",
            command=self.run_processing,
            bg="#2d7d46",
            fg="white",
            height=2
        ).pack(fill="x", padx=10, pady=10)
        ### Progress bar
        self.progress = ttk.Progressbar(
            self.root,
            orient="horizontal",
            mode="determinate"
        )

        self.progress.pack(
            fill="x",
            padx=10,
            pady=5
        )

        self.progress_label = tk.Label(
            self.root,
            text="0 / 0"
        )

        self.progress_label.pack()

        # LOG box
        self.log_box = tk.Text(self.root, height=8)
        self.log_box.pack(fill="both", padx=10, pady=10)

    # -------------------------
    # CONFIG HANDLING
    # -------------------------

    def load_config_into_ui(self):
        self.degree_editor.set_values(self.config["start_page_degree_markers"])
        self.award_editor.set_values(self.config["start_page_strong_award_markers"])
        self.special_editor.set_values(self.config["special_start_page_identifiers"])
        self.nonstart_editor.set_values(self.config["non_start_page_patterns"])

    def collect_config(self):
        return {
            "start_page_degree_markers": self.degree_editor.get_values(),
            "start_page_strong_award_markers": self.award_editor.get_values(),
            "special_start_page_identifiers": self.special_editor.get_values(),
            "non_start_page_patterns": self.nonstart_editor.get_values(),
            # Keep weak markers from default config unless you expose them
            "start_page_weak_markers": self.config["start_page_weak_markers"],
            "allow_personnummer_on_start_page": self.allow_personnummer.get(),
        }

    # -------------------------
    # FILE PICKERS
    # -------------------------

    def select_input(self):
        path = filedialog.askdirectory(title="Select input folder")
        if path:
            self.input_path.set(path)

    def select_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_path.set(path)


    # -------------------------
    # LOGGING
    # -------------------------

    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    # -------------------------
    # RUN PROCESS
    # -------------------------

    def run_processing(self):
        input_folder = self.input_path.get().strip()
        output_folder = self.output_path.get().strip()

        if not input_folder:
            messagebox.showwarning("Missing", "Select input folder")
            return

        if not output_folder:
            messagebox.showwarning("Missing", "Select output folder")
            return

        config = self.collect_config()

        self.log("Starting processing...")
        self.log(f"Input: {input_folder}")
        self.log(f"Output: {output_folder}")

        try:
            threading.Thread(
                target=self.run_processor_thread,
                args=(
                    input_folder,
                    output_folder,
                    config
                ),
                daemon=True
            ).start()



        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

# ---------
# Progress bar
# -----------

    def processor_log_callback(self, message):
        self.log_queue.put(message)

    def processor_progress_callback(self, completed, total):
        self.progress_queue.put((completed, total))

    def update_progress(self, completed, total):

        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = completed

        self.progress_label.config(
            text=f"{completed} / {total}"
        )

    def poll_queues(self):

        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log(msg)

        while not self.progress_queue.empty():
            completed, total = self.progress_queue.get()
            self.update_progress(completed, total)

        self.root.after(100, self.poll_queues)

    def run_processor_thread(
            self,
            input_folder,
            output_folder,
            config
    ):

            try:

                process_all(
                    input_folder,
                    output_folder,
                    config,
                    log_callback=self.processor_log_callback,
                    progress_callback=self.processor_progress_callback
                )

                if self.run_validation.get():
                    validate_output(
                        input_root=input_folder,
                        output_root=output_folder,
                        log_callback=self.processor_log_callback
                    )

                self.log_queue.put(
                    "Processing complete"
                )

            except Exception as e:

                self.log_queue.put(
                    f"ERROR: {e}"
                )

# -------------------------
# ENTRY
# -------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateGUI(root)
    root.mainloop()