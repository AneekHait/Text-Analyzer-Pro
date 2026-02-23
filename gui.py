import os
import webbrowser
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from cluster_tool import (
    load_excel,
    coerce_text_column,
    preprocess_texts,
    vectorize_texts,
    cluster_texts,
    get_top_keywords_per_cluster,
    assign_cluster_names,
    visualize_embeddings,
    save_results_excel,
)


class ClusterGUI:
    def __init__(self, master):
        self.master = master
        self.app_title = "Text Clustering Tool - v0.1"
        master.title(self.app_title)
        try:
            master.wm_title(self.app_title)
        except Exception:
            pass
        # Ownership / imprint information (customize as requested)
        self.owner_name = "Aneek Hait"
        self.owner_contact = "https://www.linkedin.com/in/aneekhait/"
        self.owner_website = "https://aneekhait.me"
        self.owner_bmc = "https://buymeacoffee.com/aneekh"

        # Menu with Imprint / About
        menubar = tk.Menu(master)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_imprint)
        menubar.add_cascade(label="Help", menu=helpmenu)
        master.config(menu=menubar)

        # File selection
        self.file_label = tk.Label(master, text="No file selected")
        self.file_label.grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=6)
        self.file_btn = tk.Button(master, text="Select Excel file...", command=self.select_file)
        self.file_btn.grid(row=0, column=3, padx=6, pady=6)

        # Column selection
        tk.Label(master, text="Text column:").grid(row=1, column=0, sticky="e", padx=6)
        self.col_var = tk.StringVar(master)
        self.col_menu = tk.OptionMenu(master, self.col_var, "")
        self.col_menu.grid(row=1, column=1, sticky="w", padx=6)

        # Algorithm
        tk.Label(master, text="Algorithm:").grid(row=1, column=2, sticky="e", padx=6)
        self.alg_var = tk.StringVar(master)
        self.alg_var.set("kmeans")
        tk.OptionMenu(master, self.alg_var, "kmeans", "dbscan", "agglomerative").grid(row=1, column=3, sticky="w", padx=6)

        # n_clusters
        tk.Label(master, text="n_clusters:").grid(row=2, column=0, sticky="e", padx=6)
        self.k_entry = tk.Entry(master, width=6)
        self.k_entry.insert(0, "5")
        self.k_entry.grid(row=2, column=1, sticky="w", padx=6)

        # name options
        tk.Label(master, text="name top N:").grid(row=2, column=2, sticky="e", padx=6)
        self.name_top_entry = tk.Entry(master, width=6)
        self.name_top_entry.insert(0, "3")
        self.name_top_entry.grid(row=2, column=3, sticky="w", padx=6)

        tk.Label(master, text="joiner:").grid(row=3, column=0, sticky="e", padx=6)
        self.joiner_entry = tk.Entry(master, width=6)
        self.joiner_entry.insert(0, "_")
        self.joiner_entry.grid(row=3, column=1, sticky="w", padx=6)

        # Output path
        tk.Label(master, text="Output file:").grid(row=4, column=0, sticky="e", padx=6)
        self.out_entry = tk.Entry(master, width=50)
        self.out_entry.grid(row=4, column=1, columnspan=3, sticky="w", padx=6)

        # Buttons
        self.run_btn = tk.Button(master, text="Run clustering", command=self.run_clustering_thread)
        self.run_btn.grid(row=5, column=1, pady=8)
        self.save_btn = tk.Button(master, text="Save results (use edited names)", command=self.save_with_names, state="disabled")
        self.save_btn.grid(row=5, column=2, pady=8)

        # Log / status
        self.log = tk.Text(master, height=12, width=80)
        self.log.grid(row=6, column=0, columnspan=4, padx=6, pady=6)

        # Frame for editable cluster names
        self.names_frame = tk.Frame(master)
        self.names_frame.grid(row=7, column=0, columnspan=4, sticky="we", padx=6, pady=6)

        # Footer imprint label
        self.imprint_label = tk.Label(master, text=f"© {self.owner_name} — {self.owner_website}", fg="gray")
        self.imprint_label.grid(row=8, column=0, columnspan=4, sticky="w", padx=6, pady=(0,6))

        self.df = None
        self.labels = None
        self.cluster_names = {}
        self.top_keywords = {}

    def log_msg(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def show_imprint(self):
        win = tk.Toplevel(self.master)
        # Use ASCII hyphen in title to avoid display issues on some window managers
        win.title("About - Text Clustering Tool - v0.1")
        # Make dialog modal
        win.transient(self.master)
        win.grab_set()

        text = """Text Clustering Tool — v0.1

A compact desktop GUI to cluster short text from Excel files, extract representative keywords,
assign human-readable cluster names, visualize embeddings, and export results.

Features:
- Load Excel (.xlsx/.xls) and select text column
- Preprocessing + TF-IDF vectorization
- Multiple clustering algorithms: k-means, DBSCAN, Agglomerative
- Automatic top-keyword extraction and editable cluster naming
- 2D embedding visualization (matplotlib)
- Save annotated results to Excel

Author:
- %s
- LinkedIn: %s
- Website: %s
- Support: %s

Requirements:
- Python 3.8+
- pandas, scikit-learn, openpyxl, matplotlib, seaborn, joblib, tkinter

License:
- MIT — see LICENSE file

Privacy:
- All processing happens locally; no data is sent to external services.

Support & feedback:
- File issues or suggestions on the GitHub repository.
- For quick support or sponsorship, visit the Buy Me a Coffee link above.
""" % (self.owner_name, self.owner_contact, self.owner_website, self.owner_bmc)

        # Use a scrollable text widget so long About text is readable; buttons remain fixed below
        content_frame = tk.Frame(win, padx=12, pady=12)
        content_frame.pack(fill="both", expand=True)

        # Header: show the tool name prominently
        header = tk.Label(content_frame, text="Text Clustering Tool — v0.1", font=(None, 16, "bold"))
        header.pack(side="top", anchor="w", pady=(0, 8))

        text_widget = tk.Text(content_frame, wrap="word", state="normal")
        text_widget.insert("1.0", text)
        text_widget.config(state="disabled")
        text_widget.pack(side="left", fill="both", expand=True)

        # Vertical scrollbar for the text
        vsb = tk.Scrollbar(content_frame, orient="vertical", command=text_widget.yview)
        vsb.pack(side="right", fill="y")
        text_widget.configure(yscrollcommand=vsb.set)

        # Buttons frame fixed at the bottom so buttons are always visible
        btn_frame = tk.Frame(win, pady=8)
        btn_frame.pack(side="bottom", fill="x")

        if self.owner_contact:
            tk.Button(btn_frame, text="Open LinkedIn", command=lambda: webbrowser.open(self.owner_contact)).pack(side="left", padx=6)

        if self.owner_website:
            tk.Button(btn_frame, text="Open Website", command=lambda: webbrowser.open(self.owner_website)).pack(side="left", padx=6)

        if self.owner_bmc:
            tk.Button(btn_frame, text="Buy Me a Coffee", command=lambda: webbrowser.open(self.owner_bmc)).pack(side="left", padx=6)

        tk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right", padx=6)

        # Make dialog wider by default and give a larger minimum size so the text fits
        win.minsize(900, 520)
        # Center the dialog over the main window with a reasonable offset
        self.master.update_idletasks()
        x = self.master.winfo_rootx()
        y = self.master.winfo_rooty()
        w = self.master.winfo_width()
        h = self.master.winfo_height()
        # Default geometry: wide and tall
        win.geometry(f"900x600+{x + max(10, w//12)}+{y + max(10, h//12)}")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not path:
            return
        self.file_label.config(text=path)
        # Update the window title to include the selected file for easier identification
        try:
            name = os.path.basename(path)
            self.master.title(f"{self.app_title} - {name}")
            self.master.wm_title(f"{self.app_title} - {name}")
        except Exception:
            pass
        self.out_entry.delete(0, tk.END)
        base, ext = os.path.splitext(path)
        self.out_entry.insert(0, base + "_clustered.xlsx")
        try:
            df = load_excel(path)
            self.df = df
            cols = list(df.columns)
            menu = self.col_menu["menu"]
            menu.delete(0, "end")
            for c in cols:
                menu.add_command(label=c, command=lambda value=c: self.col_var.set(value))
            if cols:
                self.col_var.set(cols[0])
            self.log_msg(f"Loaded file with {len(df)} rows and columns: {cols}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel: {e}")

    def run_clustering_thread(self):
        t = threading.Thread(target=self.run_clustering)
        t.start()

    def run_clustering(self):
        if self.df is None:
            messagebox.showwarning("No file", "Please select an Excel file first")
            return
        col = self.col_var.get()
        if not col:
            messagebox.showwarning("No column", "Please select a text column")
            return

        try:
            texts = coerce_text_column(self.df[col]).tolist()
            processed = preprocess_texts(texts)
            vectorizer, X = vectorize_texts(processed)
            alg = self.alg_var.get()
            n_clusters = int(self.k_entry.get() or 5)
            model, labels = cluster_texts(X, algorithm=alg, n_clusters=n_clusters)
            self.labels = labels
            self.df["cluster_label"] = labels

            # top keywords and names
            top_n = int(self.name_top_entry.get() or 3)
            self.top_keywords = get_top_keywords_per_cluster(vectorizer, X, labels, top_n=10)
            self.cluster_names = assign_cluster_names(self.top_keywords, name_top_n=top_n, joiner=self.joiner_entry.get())

            # populate editable names
            self.populate_name_entries()

            self.log_msg("Clustering finished. Edit names below and click 'Save results'.")
            # enable save btn
            self.save_btn.config(state="normal")
        except Exception as e:
            messagebox.showerror("Error during clustering", str(e))

    def populate_name_entries(self):
        # Clear frame
        for w in self.names_frame.winfo_children():
            w.destroy()

        sorted_ids = sorted(self.cluster_names.keys())
        self.name_entries = {}
        for i, cid in enumerate(sorted_ids):
            tk.Label(self.names_frame, text=f"{cid}:").grid(row=i, column=0, sticky="e")
            ent = tk.Entry(self.names_frame, width=40)
            ent.insert(0, self.cluster_names[cid])
            ent.grid(row=i, column=1, sticky="w", padx=6, pady=2)
            # show top keywords as label
            kw = ", ".join([t for t, s in self.top_keywords.get(cid, [])])
            tk.Label(self.names_frame, text=kw).grid(row=i, column=2, sticky="w", padx=6)
            self.name_entries[cid] = ent

    def save_with_names(self):
        if self.df is None or self.labels is None:
            messagebox.showwarning("Nothing to save", "Run clustering first")
            return
        # read edited names
        final_names = {}
        for cid, ent in self.name_entries.items():
            name = ent.get().strip()
            final_names[cid] = name

        # apply to df
        self.df["cluster_name"] = [final_names.get(int(l), "") for l in self.labels]
        out = self.out_entry.get().strip()
        if not out:
            messagebox.showwarning("No output", "Provide an output filepath")
            return
        try:
            save_results_excel(self.df, out)
            messagebox.showinfo("Saved", f"Saved results to {out}")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))


def main():
    root = tk.Tk()
    # Ensure the main window has an explicit title (some WMs require title set on root)
    root.title("Text Clustering Tool - v0.1")
    root.wm_title("Text Clustering Tool - v0.1")
    # Try to load an application icon if one is available.
    # Looks for 'icon.png' or 'assets/icon.png' next to this file. If found, sets it via iconphoto.
    def _set_app_icon(root_window):
        here = os.path.dirname(__file__)
        candidates = [
            os.path.join(here, "icon.png"),
            os.path.join(here, "assets", "icon.png"),
            os.path.join(here, "icon.ico"),
            os.path.join(here, "assets", "icon.ico"),
        ]
        for fp in candidates:
            try:
                if os.path.exists(fp):
                    # PhotoImage supports PNG/GIF; try iconphoto first
                    img = tk.PhotoImage(file=fp)
                    root_window.iconphoto(True, img)
                    # keep a reference to prevent GC
                    root_window._icon_image = img
                    return True
            except Exception:
                # fallback: try iconbitmap for .ico
                try:
                    root_window.iconbitmap(fp)
                    return True
                except Exception:
                    continue
        return False

    _set_app_icon(root)

    # Withdraw the window first on Crostini so the WM has time to register it;
    # then show it after a short delay with a temporary topmost toggle to force decorations.
    try:
        root.withdraw()
    except Exception:
        pass

    app = ClusterGUI(root)

    def _show_root():
        try:
            # Create a tiny temporary Toplevel to nudge the window manager into drawing
            try:
                tmp = tk.Toplevel(root)
                tmp.overrideredirect(True)
                tmp.geometry("1x1+0+0")
                tmp.update_idletasks()
                # Destroy the tiny helper after a short delay to give the WM time to register it
                # (60 ms is a good compromise between speed and reliability on Crostini)
                root.after(60, lambda: (tmp.destroy() if tmp.winfo_exists() else None))
            except Exception:
                # ignore failures
                pass

            root.deiconify()
            root.lift()
            root.attributes("-topmost", True)
            # clear topmost after a short delay
            root.after(150, lambda: root.attributes("-topmost", False))
        except Exception:
            try:
                root.deiconify()
                root.lift()
            except Exception:
                pass

    # Schedule showing the window shortly after start so WM can decorate it (helps Crostini)
    root.after(80, _show_root)

    root.mainloop()


if __name__ == "__main__":
    main()
