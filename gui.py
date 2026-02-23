import os
import webbrowser
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import joblib
from datetime import datetime
import numpy as np

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
        self.app_title = "Text Analyzer Pro - v1.5"
        master.title(self.app_title)
        try:
            master.wm_title(self.app_title)
        except Exception:
            pass
        
        # Configure professional styling
        self._configure_styles()
        
        # Ownership / imprint information
        self.owner_name = "Aneek Hait"
        self.owner_contact = "https://www.linkedin.com/in/aneekhait/"
        self.owner_website = "https://aneekhait.github.io"
        self.owner_bmc = "https://buymeacoffee.com/aneekh"

        # Menu with Imprint / About
        menubar = tk.Menu(master, bg="#f0f0f0")
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_imprint)
        menubar.add_cascade(label="Help", menu=helpmenu)
        master.config(menu=menubar)

        # ===== FILE SELECTION FRAME ====
        file_frame = ttk.Frame(master, style="Card.TFrame")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        
        # File selection row
        file_row = ttk.Frame(file_frame, style="Card.TFrame")
        file_row.pack(side="top", fill="x", padx=0, pady=(8, 4))
        
        file_label_text = ttk.Label(file_row, text="📁 File:", style="Header.TLabel")
        file_label_text.pack(side="left", padx=(8, 6), pady=4)
        
        self.file_label = ttk.Label(file_row, text="No file selected", foreground="#666666", style="Header.TLabel")
        self.file_label.pack(side="left", fill="x", expand=True, padx=0, pady=4)
        
        self.file_btn = ttk.Button(file_row, text="📂 Select Excel file...", command=self.select_file)
        self.file_btn.pack(side="right", padx=(6, 8), pady=4)
        
        # Sheet selection row
        sheet_row = ttk.Frame(file_frame, style="Card.TFrame")
        sheet_row.pack(side="top", fill="x", padx=0, pady=(4, 8))
        
        sheet_label_text = ttk.Label(sheet_row, text="📄 Sheet:", style="Section.TLabel")
        sheet_label_text.pack(side="left", padx=(8, 6), pady=4)
        
        self.sheet_var = tk.StringVar(master)
        self.sheet_menu = ttk.OptionMenu(sheet_row, self.sheet_var, "")
        self.sheet_menu.pack(side="left", padx=0, pady=4)
        
        # Store file path for sheet loading
        self.current_file_path = None

        # ===== PARAMETERS FRAME ====
        params_frame = ttk.LabelFrame(master, text="⚙️  Clustering Parameters", style="TLabelframe", padding=12)
        params_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        
        # Row 1: Text column, Algorithm
        ttk.Label(params_frame, text="Text column:", style="Section.TLabel").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.col_var = tk.StringVar(master)
        self.col_menu = ttk.OptionMenu(params_frame, self.col_var, "")
        self.col_menu.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        
        ttk.Label(params_frame, text="Algorithm:", style="Section.TLabel").grid(row=0, column=2, sticky="e", padx=8, pady=6)
        self.alg_var = tk.StringVar(master)
        self.alg_var.set("kmeans")
        ttk.OptionMenu(params_frame, self.alg_var, "kmeans", "kmeans", "dbscan", "agglomerative").grid(row=0, column=3, sticky="w", padx=8, pady=6)
        
        # Row 2: n_clusters, name top N
        ttk.Label(params_frame, text="n_clusters:", style="Section.TLabel").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        self.k_entry = ttk.Entry(params_frame, width=10)
        self.k_entry.insert(0, "5")
        self.k_entry.grid(row=1, column=1, sticky="w", padx=8, pady=6)
        # Ensure the algorithm-change callback only runs after related widgets exist
        self.alg_var.trace_add("write", self._on_alg_change)
        
        ttk.Label(params_frame, text="name top N:", style="Section.TLabel").grid(row=1, column=2, sticky="e", padx=8, pady=6)
        self.name_top_entry = ttk.Entry(params_frame, width=10)
        self.name_top_entry.insert(0, "3")
        self.name_top_entry.grid(row=1, column=3, sticky="w", padx=8, pady=6)
        
        # Row 3: joiner, visualization
        ttk.Label(params_frame, text="joiner:", style="Section.TLabel").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        self.joiner_entry = ttk.Entry(params_frame, width=10)
        self.joiner_entry.insert(0, "_")
        self.joiner_entry.grid(row=2, column=1, sticky="w", padx=8, pady=6)
        
        ttk.Label(params_frame, text="Visualization:", style="Section.TLabel").grid(row=2, column=2, sticky="e", padx=8, pady=6)
        self.vis_var = tk.StringVar(master)
        self.vis_var.set("pca")
        ttk.OptionMenu(params_frame, self.vis_var, "pca", "pca", "tsne").grid(row=2, column=3, sticky="w", padx=8, pady=6)
        
        # Row 4: Output file (full width)
        ttk.Label(params_frame, text="Output file:", style="Section.TLabel").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        self.out_entry = ttk.Entry(params_frame)
        self.out_entry.grid(row=3, column=1, columnspan=3, sticky="ew", padx=8, pady=6)
        params_frame.columnconfigure(1, weight=1)

        # ===== ACTION BUTTONS FRAME =====
        btn_frame = ttk.Frame(master, style="TFrame")
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=12)
        
        self.run_btn = ttk.Button(btn_frame, text="▶️  Run Clustering", command=self.run_clustering_thread)
        self.run_btn.pack(side="left", padx=4)
        
        self.save_btn = ttk.Button(btn_frame, text="💾  Save Results", command=self.save_with_names, state="disabled")
        self.save_btn.pack(side="left", padx=4)
        
        self.vis_btn = ttk.Button(btn_frame, text="📊  Visualize", command=self.visualize_clusters, state="disabled")
        self.vis_btn.pack(side="left", padx=4)
        
        self.save_model_btn = ttk.Button(btn_frame, text="💾  Save Model", command=self.save_model, state="disabled")
        self.save_model_btn.pack(side="left", padx=4)
        
        self.clear_log_btn = ttk.Button(btn_frame, text="🗑️  Clear Log", command=self.clear_log)
        self.clear_log_btn.pack(side="left", padx=4)

        # ===== LOG SECTION =====
        log_label = ttk.Label(master, text="📝 Status Log:", style="Title.TLabel")
        log_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=(12, 6))
        
        log_frame = ttk.Frame(master, style="Card.TFrame")
        log_frame.grid(row=4, column=0, columnspan=2, sticky="ewns", padx=8, pady=8)
        
        self.log = tk.Text(log_frame, height=10, width=100, bg="#f9f9f9", fg="#333333", font=("Segoe UI", 9), relief="solid", borderwidth=1)
        self.log.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.pack(side="right", fill="y")
        self.log.config(yscrollcommand=scrollbar.set)
        
        self.progress = ttk.Progressbar(master, orient="horizontal", mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=4)
        self.progress.grid_remove()

        # ===== CLUSTER NAMES SECTION =====
        names_label = ttk.Label(master, text="✏️  Edit Cluster Names:", style="Title.TLabel")
        names_label.grid(row=6, column=0, columnspan=2, sticky="w", padx=8, pady=(12, 6))
        
        self.names_frame = ttk.Frame(master, style="Card.TFrame")
        self.names_frame.grid(row=7, column=0, columnspan=2, sticky="ewns", padx=8, pady=8)

        # ===== FOOTER =====
        footer_frame = ttk.Frame(master, style="TFrame")
        footer_frame.grid(row=8, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        
        # Left side: Ownership and copyright
        left_footer = ttk.Frame(footer_frame, style="TFrame")
        left_footer.pack(side="left", fill="x", expand=True)
        
        copyright_text = f"© 2026 {self.owner_name}  •  All rights reserved  •  v1.5"
        self.imprint_label = ttk.Label(left_footer, text=copyright_text, foreground="#999999", font=("Segoe UI", 7))
        self.imprint_label.pack(side="left")
        
        # Right side: Links and more info
        right_footer = ttk.Frame(footer_frame, style="TFrame")
        right_footer.pack(side="right")
        
        link_text = f"🔗 {self.owner_website}"
        link_label = ttk.Label(right_footer, text=link_text, foreground="#0066cc", font=("Segoe UI", 7), cursor="hand2")
        link_label.pack(side="right")
        link_label.bind("<Button-1>", lambda e: webbrowser.open(self.owner_website))

        # ===== CONFIGURE GRID WEIGHTS =====
        master.columnconfigure(0, weight=1)
        master.rowconfigure(4, weight=1)
        master.rowconfigure(7, weight=0)

        self.df = None
        self.labels = None
        self.cluster_names = {}
        self.top_keywords = {}
        self.X = None
        self.model = None
        self.vectorizer = None

    def _configure_styles(self):
        """Configure professional styling for ttk widgets"""
        style = ttk.Style()
        
        # Define colors
        bg_color = "#f5f5f5"
        card_bg = "#ffffff"
        accent_color = "#0066cc"
        text_color = "#333333"
        border_color = "#dddddd"
        
        # Configure frame styles
        style.configure("Card.TFrame", background=card_bg, relief="flat", borderwidth=0)
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=text_color)
        style.configure("TLabelframe", background=bg_color, foreground=text_color)
        
        # Header label style
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"), background=card_bg, foreground=text_color)
        style.configure("Title.TLabel", font=("Segoe UI", 12, "bold"), background=bg_color, foreground="#000000")
        style.configure("Section.TLabel", font=("Segoe UI", 10, "bold"), background=bg_color, foreground="#333333")
        
        # Button styles - keep text visible on hover
        style.configure("TButton", padding=6, font=("Segoe UI", 9), foreground=text_color)
        style.map("TButton", 
                  foreground=[("pressed", text_color), ("active", text_color), ("!active", text_color)],
                  background=[("pressed", "#e0e0e0"), ("active", "#f0f0f0")])
        
        # Entry styles
        style.configure("TEntry", padding=4, font=("Segoe UI", 10))
        
        # OptionMenu styles
        style.configure("TCombobox", padding=4, font=("Segoe UI", 10))
        
        # Progressbar style
        style.configure("TProgressbar", thickness=20)
        
        # Configure main window background
        self.master.configure(bg=bg_color)
    
    def _on_alg_change(self, *args):
        if self.alg_var.get() == "dbscan":
            self.k_entry.config(state="disabled")
        else:
            self.k_entry.config(state="normal")

    def log_msg(self, msg: str):
        """Log a message with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_msg = f"[{timestamp}] {msg}"
        self.log.insert(tk.END, formatted_msg + "\n")
        self.log.see(tk.END)
        self.master.update_idletasks()
    
    def clear_log(self):
        """Clear the log window"""
        self.log.delete("1.0", tk.END)
        self.log_msg("Log cleared.")

    def show_imprint(self):
        win = tk.Toplevel(self.master)
        win.title("About - Text Analyzer Pro - v1.5")
        win.transient(self.master)
        win.grab_set()

        text = """TEXT ANALYZER PRO v1.5
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DESCRIPTION:
A modern, user-friendly desktop application for clustering and analyzing text data 
from Excel workbooks. Extract meaningful patterns, assign human-readable cluster names,
and visualize results with minimal effort.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

KEY FEATURES:
✓ Multi-sheet Excel support (.xlsx, .xls)
✓ Flexible text column selection
✓ Multiple clustering algorithms (K-Means, DBSCAN, Agglomerative)
✓ Automatic keyword extraction and cluster naming
✓ Interactive cluster name editing
✓ 2D visualization (PCA, t-SNE)
✓ Model persistence (save & load)
✓ Professional, responsive UI

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TECHNOLOGY STACK:
• Python 3.8+
• pandas – Data manipulation
• scikit-learn – ML algorithms
• openpyxl – Excel I/O  
• matplotlib – Visualization
• seaborn – Advanced plots
• tkinter – GUI framework
• ttkthemes – Modern styling

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

AUTHOR & OWNERSHIP:

Name: %s
LinkedIn: %s
Website: %s
Support: %s

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

LICENSE & PRIVACY:

License: MIT License
See LICENSE file for full terms.

Privacy: 
All processing happens locally on your machine. 
No data is sent to external servers or services.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SUPPORT & FEEDBACK:
• Report issues on GitHub
• Request features on GitHub
• Sponsor development via Buy Me a Coffee
• Direct inquiries via LinkedIn

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 LOVE THIS TOOL?

If Text Analyzer Pro helped you analyze data like a genius,
saved you hours of manual work, or made your research flow smoother,
consider buying me a coffee! ☕

Your support fuels development of new features, improvements, and
keeps this tool free and maintained for everyone.

Every coffee brings us closer to v2.0! 🚀

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VERSION: 1.5
RELEASE DATE: February 2026
STATUS: Active Development

ENHANCEMENTS IN v1.5:
✨ Professional UI redesign with modern styling
✨ Multi-sheet Excel file support
✨ Input validation & error handling
✨ Timestamped logging with clear feedback
✨ Cluster statistics display (count, percentages)
✨ Clear log functionality
✨ Responsive grid layout
✨ Better window scaling & DPI awareness
✨ Improved About dialog with ownership info
✨ Enhanced button labels with emojis

FOUNDATION (v0.1):
• Core text clustering engine
• K-Means, DBSCAN, Agglomerative algorithms
• TF-IDF vectorization
• Excel file loading & column selection
• Automatic keyword extraction
• Cluster naming system
• 2D visualization (PCA, t-SNE)
• Model save/load with joblib
• Basic GUI interface



© 2026 %s. All rights reserved.
""" % (self.owner_name, self.owner_contact, self.owner_website, self.owner_bmc, self.owner_name)

        # Use a scrollable text widget so long About text is readable; buttons remain fixed below
        content_frame = ttk.Frame(win, padding=12)
        content_frame.pack(fill="both", expand=True)

        # Header: show the tool name prominently
        header = ttk.Label(content_frame, text="✨ Text Analyzer Pro — v1.5", font=("Segoe UI", 14, "bold"))
        header.pack(side="top", anchor="w", pady=(0, 8))

        text_widget = tk.Text(content_frame, wrap="word", state="normal", bg="#f9f9f9", fg="#333333", font=("Consolas", 9), height=20)
        text_widget.insert("1.0", text)
        text_widget.config(state="disabled")
        text_widget.pack(side="left", fill="both", expand=True)
        
        # Scroll to top to show ownership info
        text_widget.see("1.0")

        # Vertical scrollbar for the text
        vsb = ttk.Scrollbar(content_frame, orient="vertical", command=text_widget.yview)
        vsb.pack(side="right", fill="y")
        text_widget.configure(yscrollcommand=vsb.set)

        # Buttons frame fixed at the bottom so buttons are always visible
        btn_frame = ttk.Frame(win, padding=12)
        btn_frame.pack(side="bottom", fill="x")

        # Create a frame for ownership buttons
        contact_label = ttk.Label(btn_frame, text="🔗 Connect with Author:", font=("Segoe UI", 9, "bold"))
        contact_label.pack(side="left", padx=(0, 10))

        if self.owner_contact:
            ttk.Button(btn_frame, text="💼 LinkedIn", command=lambda: webbrowser.open(self.owner_contact)).pack(side="left", padx=3)

        if self.owner_website:
            ttk.Button(btn_frame, text="🌐 Website", command=lambda: webbrowser.open(self.owner_website)).pack(side="left", padx=3)

        if self.owner_bmc:
            ttk.Button(btn_frame, text="☕ Buy Me a Coffee", command=lambda: webbrowser.open(self.owner_bmc)).pack(side="left", padx=3)

        ttk.Button(btn_frame, text="❌ Close", command=win.destroy).pack(side="right", padx=3)

        # Make dialog wider by default and give a larger minimum size so the text fits
        win.minsize(1000, 700)
        # Center the dialog over the main window with a reasonable offset
        self.master.update_idletasks()
        x = self.master.winfo_rootx()
        y = self.master.winfo_rooty()
        w = self.master.winfo_width()
        h = self.master.winfo_height()
        # Default geometry: wide and tall
        win.geometry(f"1000x750+{x + max(10, w//12)}+{y + max(10, h//12)}")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not path:
            return
        self.current_file_path = path
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
        
        # Load sheet names
        try:
            import openpyxl
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet_names = workbook.sheetnames
            workbook.close()
            
            # Populate sheet dropdown
            menu = self.sheet_menu["menu"]
            menu.delete(0, "end")
            for sheet in sheet_names:
                menu.add_command(label=sheet, command=lambda value=sheet: self._load_sheet(value))
            
            # Auto-select first sheet
            if sheet_names:
                self.sheet_var.set(sheet_names[0])
                self._load_sheet(sheet_names[0])
            
            self.log_msg(f"✓ Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel sheets: {e}")
            self.log_msg(f"✗ Error reading sheets: {e}")
    
    def _load_sheet(self, sheet_name):
        """Load data from the selected sheet"""
        if not self.current_file_path:
            return
        try:
            df = load_excel(self.current_file_path, sheet_name=sheet_name)
            self.df = df
            cols = list(df.columns)
            menu = self.col_menu["menu"]
            menu.delete(0, "end")
            for c in cols:
                menu.add_command(label=c, command=lambda value=c: self.col_var.set(value))
            if cols:
                self.col_var.set(cols[0])
            file_size_kb = os.path.getsize(self.current_file_path) / 1024
            self.log_msg(f"✓ Loaded sheet '{sheet_name}': {len(df)} rows, {len(cols)} columns, {file_size_kb:.1f} KB")
            self.log_msg(f"  Columns: {', '.join(cols)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {e}")
            self.log_msg(f"✗ Error loading sheet: {e}")

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
        
        # Input validation
        try:
            n_clusters = int(self.k_entry.get() or 5)
            if n_clusters < 2:
                messagebox.showwarning("Invalid parameter", "n_clusters must be at least 2")
                return
            if n_clusters > len(self.df):
                messagebox.showwarning("Invalid parameter", f"n_clusters ({n_clusters}) cannot exceed data size ({len(self.df)})")
                return
            top_n = int(self.name_top_entry.get() or 3)
            if top_n < 1:
                messagebox.showwarning("Invalid parameter", "name top N must be at least 1")
                return
        except ValueError as e:
            messagebox.showerror("Invalid input", f"Please enter valid numbers for parameters: {e}")
            return

        try:
            self.run_btn.config(state="disabled")
            self.save_btn.config(state="disabled")
            self.vis_btn.config(state="disabled")
            self.save_model_btn.config(state="disabled")
            self.progress.grid()
            self.progress["value"] = 0

            self.log_msg("="*60)
            self.log_msg(f"Starting clustering (Algorithm: {self.alg_var.get()}, n_clusters: {n_clusters})")
            self.progress["value"] = 5
            self.master.update_idletasks()
            texts = coerce_text_column(self.df[col]).tolist()

            self.log_msg("Preprocessing texts...")
            self.progress["value"] = 10
            self.master.update_idletasks()
            processed = preprocess_texts(texts)

            self.log_msg("Vectorizing texts...")
            self.progress["value"] = 30
            self.master.update_idletasks()
            vectorizer, X = vectorize_texts(processed)
            self.X = X
            self.vectorizer = vectorizer
            self.log_msg(f"  Vectorizer created: {X.shape[0]} documents, {X.shape[1]} features")

            self.log_msg("Clustering texts...")
            self.progress["value"] = 70
            self.master.update_idletasks()
            alg = self.alg_var.get()
            model, labels = cluster_texts(X, algorithm=alg, n_clusters=n_clusters)
            self.model = model
            self.labels = labels
            self.df["cluster_label"] = labels
            
            # Show cluster statistics
            unique_labels = np.unique(labels)
            self.log_msg(f"  Clusters found: {len(unique_labels)}")
            for label in unique_labels:
                count = np.sum(labels == label)
                percentage = (count / len(labels)) * 100
                self.log_msg(f"    Cluster {label}: {count} items ({percentage:.1f}%)")

            # top keywords and names
            self.log_msg("Extracting top keywords...")
            self.progress["value"] = 90
            self.master.update_idletasks()
            self.top_keywords = get_top_keywords_per_cluster(vectorizer, X, labels, top_n=10)
            self.cluster_names = assign_cluster_names(self.top_keywords, name_top_n=top_n, joiner=self.joiner_entry.get())
            
            self.log_msg("Suggested cluster names:")
            for cid, name in self.cluster_names.items():
                self.log_msg(f"  {cid}: {name}")

            # populate editable names
            self.populate_name_entries()

            self.log_msg("✓ Clustering finished! Edit names below and click 'Save results'")
            self.progress["value"] = 100
            self.master.update_idletasks()
            # enable save btn
            self.save_btn.config(state="normal")
            self.vis_btn.config(state="normal")
            self.save_model_btn.config(state="normal")
        except Exception as e:
            self.log_msg(f"✗ Clustering error: {str(e)}")
            messagebox.showerror("Error during clustering", str(e))
        finally:
            self.run_btn.config(state="normal")
            self.progress.grid_remove()

    def populate_name_entries(self):
        # Clear frame
        for w in self.names_frame.winfo_children():
            w.destroy()

        sorted_ids = sorted(self.cluster_names.keys())
        self.name_entries = {}
        for i, cid in enumerate(sorted_ids):
            ttk.Label(self.names_frame, text=f"{cid}:").grid(row=i, column=0, sticky="e")
            ent = ttk.Entry(self.names_frame, width=40)
            ent.insert(0, self.cluster_names[cid])
            ent.grid(row=i, column=1, sticky="w", padx=6, pady=2)
            # show top keywords as label
            kw = ", ".join([t for t, s in self.top_keywords.get(cid, [])])
            ttk.Label(self.names_frame, text=kw).grid(row=i, column=2, sticky="w", padx=6)
            self.name_entries[cid] = ent

    def visualize_clusters(self):
        if self.X is None or self.labels is None:
            messagebox.showwarning("Nothing to visualize", "Run clustering first")
            return
        method = self.vis_var.get()
        self.log_msg(f"Generating {method.upper()} visualization...")
        try:
            visualize_embeddings(self.X, self.labels, method=method)
            self.log_msg(f"✓ {method.upper()} visualization displayed")
        except Exception as e:
            self.log_msg(f"✗ Visualization failed: {str(e)}")
            messagebox.showerror("Visualization failed", str(e))

    def save_with_names(self):
        if self.df is None or self.labels is None:
            messagebox.showwarning("Nothing to save", "Run clustering first")
            return
        # read edited names
        final_names = {}
        for cid, ent in self.name_entries.items():
            name = ent.get().strip()
            if not name:
                messagebox.showwarning("Invalid input", f"Cluster name for cluster {cid} cannot be empty")
                return
            final_names[cid] = name

        # apply to df
        self.df["cluster_name"] = [final_names.get(int(l), "") for l in self.labels]
        out = self.out_entry.get().strip()
        if not out:
            messagebox.showwarning("No output", "Provide an output filepath")
            return
        try:
            save_results_excel(self.df, out)
            self.log_msg(f"✓ Results saved to {out}")
            messagebox.showinfo("Saved", f"Saved results to {out}")
        except Exception as e:
            self.log_msg(f"✗ Save failed: {str(e)}")
            messagebox.showerror("Save failed", str(e))

    def save_model(self):
        if self.model is None or self.vectorizer is None:
            messagebox.showwarning("Nothing to save", "Run clustering first")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".joblib",
            filetypes=[("Joblib files", "*.joblib")],
            title="Save Clustering Model"
        )
        if not path:
            return

        try:
            joblib.dump(
                {
                    "model": self.model,
                    "vectorizer": self.vectorizer,
                    "cluster_names": self.cluster_names,
                    "top_keywords": self.top_keywords,
                },
                path,
            )
            self.log_msg(f"✓ Model saved to {path}")
            messagebox.showinfo("Model Saved", f"Saved model to {path}")
        except Exception as e:
            self.log_msg(f"✗ Model save failed: {str(e)}")
            messagebox.showerror("Save Failed", f"Failed to save model: {e}")


from ttkthemes import ThemedTk

def main():
    root = ThemedTk(theme="arc")
    # Set DPI awareness for proper scaling on Windows
    try:
        root.tk.call('tk', 'scaling', 2.0)
    except Exception:
        pass
    # Ensure the main window has an explicit title (some WMs require title set on root)
    root.title("Text Analyzer Pro - v1.5")
    root.wm_title("Text Analyzer Pro - v1.5")
    # Set initial window geometry for better sizing
    root.geometry("1100x800")
    root.minsize(950, 650)
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
