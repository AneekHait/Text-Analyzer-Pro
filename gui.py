import os
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
        master.title("Text Clustering GUI")

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

        self.df = None
        self.labels = None
        self.cluster_names = {}
        self.top_keywords = {}

    def log_msg(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not path:
            return
        self.file_label.config(text=path)
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
    app = ClusterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
