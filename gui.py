import os
import webbrowser
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import joblib
from datetime import datetime
import numpy as np
from PIL import ImageTk

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
from wordcloud_tool import (
    WordCloudConfig,
    export_term_stats,
    get_effective_stopwords,
    prepare_wordcloud_data,
    render_wordcloud,
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

        self.wordcloud_btn = ttk.Button(
            btn_frame,
            text="☁️  Generate Wordcloud",
            command=self.open_wordcloud_builder,
            state="disabled",
        )
        self.wordcloud_btn.pack(side="left", padx=4)

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
        self.wordcloud_builder = None

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
            self.wordcloud_btn.config(state="disabled")
    
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
                self.wordcloud_btn.config(state="normal")
            else:
                self.wordcloud_btn.config(state="disabled")
            file_size_kb = os.path.getsize(self.current_file_path) / 1024
            self.log_msg(f"✓ Loaded sheet '{sheet_name}': {len(df)} rows, {len(cols)} columns, {file_size_kb:.1f} KB")
            self.log_msg(f"  Columns: {', '.join(cols)}")
            if self.wordcloud_builder is not None:
                self.wordcloud_builder.refresh_from_app()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {e}")
            self.log_msg(f"✗ Error loading sheet: {e}")
            self.wordcloud_btn.config(state="disabled")

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

    def open_wordcloud_builder(self):
        if self.df is None:
            messagebox.showwarning("No file", "Please select an Excel file and sheet first")
            return

        existing_builder = self.wordcloud_builder
        if existing_builder is not None and existing_builder.is_alive:
            existing_builder.refresh_from_app()
            existing_builder.focus()
            return

        self.wordcloud_builder = WordCloudBuilderWindow(self)

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


class WordCloudBuilderWindow:
    PHRASE_OPTIONS = ("Unigrams", "Up to Bigrams", "Up to Trigrams")
    BACKGROUND_OPTIONS = ("white", "ivory", "whitesmoke", "mintcream", "black", "midnightblue")
    COLORMAP_OPTIONS = ("viridis", "plasma", "inferno", "magma", "cividis", "Set2", "tab10", "cubehelix")

    def __init__(self, app):
        self.app = app
        self.window = tk.Toplevel(app.master)
        self.window.title(f"{self.app.app_title} - Wordcloud Builder")
        self.window.geometry("1280x780")
        self.window.minsize(1120, 700)
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        self.custom_stopwords = set()
        self.current_stats_df = None
        self.current_image = None
        self.preview_photo = None
        self.is_rendering = False

        self.context_var = tk.StringVar(self.window, value="No active sheet")
        self.status_var = tk.StringVar(self.window, value="Configure the controls, then generate a preview.")
        self.stopword_count_var = tk.StringVar(self.window, value="Effective stopwords: 0")
        self.column_var = tk.StringVar(self.window)
        self.max_words_var = tk.StringVar(self.window, value="200")
        self.min_frequency_var = tk.StringVar(self.window, value="1")
        self.width_var = tk.StringVar(self.window, value="1200")
        self.height_var = tk.StringVar(self.window, value="700")
        self.phrase_mode_var = tk.StringVar(self.window, value=self.PHRASE_OPTIONS[0])
        self.use_builtin_stopwords_var = tk.BooleanVar(self.window, value=True)
        self.lowercase_var = tk.BooleanVar(self.window, value=True)
        self.exclude_numeric_var = tk.BooleanVar(self.window, value=True)
        self.background_var = tk.StringVar(self.window, value="white")
        self.colormap_var = tk.StringVar(self.window, value="viridis")
        self.stopword_entry_var = tk.StringVar(self.window)

        self.total_rows_var = tk.StringVar(self.window, value="0")
        self.usable_rows_var = tk.StringVar(self.window, value="0")
        self.unique_terms_var = tk.StringVar(self.window, value="0")
        self.term_occurrences_var = tk.StringVar(self.window, value="0")

        self._build_layout()
        self.refresh_from_app(reset_preview=False)
        self._reset_preview_state("Generate a wordcloud to preview it here.", clear_summary=True)
        self.update_stopword_count()
        self.window.after(150, self.generate_wordcloud_thread)

    @property
    def is_alive(self):
        try:
            return bool(self.window.winfo_exists())
        except tk.TclError:
            return False

    def focus(self):
        if not self.is_alive:
            return
        self.window.deiconify()
        self.window.lift()
        self.window.focus_force()

    def close(self):
        if self.app.wordcloud_builder is self:
            self.app.wordcloud_builder = None
        if self.is_alive:
            self.window.destroy()

    def refresh_from_app(self, reset_preview=True):
        if not self.is_alive:
            return

        columns = list(self.app.df.columns) if self.app.df is not None else []
        current_choice = self.column_var.get().strip()
        preferred_choice = self.app.col_var.get().strip()

        self.column_combo["values"] = columns
        if preferred_choice in columns:
            self.column_var.set(preferred_choice)
        elif current_choice in columns:
            self.column_var.set(current_choice)
        elif columns:
            self.column_var.set(columns[0])
        else:
            self.column_var.set("")

        file_name = os.path.basename(self.app.current_file_path) if self.app.current_file_path else "No file"
        sheet_name = self.app.sheet_var.get().strip() or "No sheet"
        column_name = self.column_var.get().strip() or "No column"
        self.context_var.set(f"File: {file_name}    Sheet: {sheet_name}    Column: {column_name}")

        if reset_preview:
            self._reset_preview_state("Wordcloud context changed. Click Generate Preview to refresh.", clear_summary=True)

        self.generate_btn.config(state="normal" if columns and not self.is_rendering else "disabled")

    def _build_layout(self):
        outer = ttk.Frame(self.window, padding=12, style="TFrame")
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        header_frame = ttk.Frame(outer, style="TFrame")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header_frame.columnconfigure(0, weight=1)

        ttk.Label(header_frame, text="☁️  Wordcloud Builder", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header_frame, textvariable=self.context_var, style="TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

        content_frame = ttk.Frame(outer, style="TFrame")
        content_frame.grid(row=1, column=0, sticky="nsew")
        content_frame.columnconfigure(0, weight=0)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)

        left_panel = ttk.Frame(content_frame, style="Card.TFrame", padding=12)
        left_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        right_panel = ttk.Frame(content_frame, style="Card.TFrame", padding=12)
        right_panel.grid(row=0, column=1, sticky="nsew")
        right_panel.columnconfigure(0, weight=1)
        right_panel.rowconfigure(1, weight=1)

        controls_frame = ttk.LabelFrame(left_panel, text="Controls", padding=10)
        controls_frame.grid(row=0, column=0, sticky="ew")
        controls_frame.columnconfigure(1, weight=1)

        ttk.Label(controls_frame, text="Source column:", style="Section.TLabel").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
        self.column_combo = ttk.Combobox(controls_frame, textvariable=self.column_var, state="readonly", width=26)
        self.column_combo.grid(row=0, column=1, sticky="ew", pady=4)
        self.column_combo.bind("<<ComboboxSelected>>", lambda _event: self.refresh_from_app(reset_preview=True))

        ttk.Label(controls_frame, text="Max words:", style="Section.TLabel").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(controls_frame, textvariable=self.max_words_var, width=12).grid(row=1, column=1, sticky="ew", pady=4)

        ttk.Label(controls_frame, text="Min frequency:", style="Section.TLabel").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(controls_frame, textvariable=self.min_frequency_var, width=12).grid(row=2, column=1, sticky="ew", pady=4)

        ttk.Label(controls_frame, text="Width:", style="Section.TLabel").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(controls_frame, textvariable=self.width_var, width=12).grid(row=3, column=1, sticky="ew", pady=4)

        ttk.Label(controls_frame, text="Height:", style="Section.TLabel").grid(row=4, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Entry(controls_frame, textvariable=self.height_var, width=12).grid(row=4, column=1, sticky="ew", pady=4)

        ttk.Label(controls_frame, text="Phrase mode:", style="Section.TLabel").grid(row=5, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Combobox(
            controls_frame,
            textvariable=self.phrase_mode_var,
            values=self.PHRASE_OPTIONS,
            state="readonly",
            width=20,
        ).grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Checkbutton(
            controls_frame,
            text="Use built-in stopwords",
            variable=self.use_builtin_stopwords_var,
            command=self.update_stopword_count,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 2))
        ttk.Checkbutton(
            controls_frame,
            text="Lowercase normalization",
            variable=self.lowercase_var,
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Checkbutton(
            controls_frame,
            text="Exclude numeric-only tokens",
            variable=self.exclude_numeric_var,
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=2)

        ttk.Label(controls_frame, text="Background:", style="Section.TLabel").grid(row=9, column=0, sticky="e", padx=(0, 8), pady=(8, 4))
        ttk.Combobox(
            controls_frame,
            textvariable=self.background_var,
            values=self.BACKGROUND_OPTIONS,
            state="readonly",
            width=20,
        ).grid(row=9, column=1, sticky="ew", pady=(8, 4))

        ttk.Label(controls_frame, text="Colormap:", style="Section.TLabel").grid(row=10, column=0, sticky="e", padx=(0, 8), pady=4)
        ttk.Combobox(
            controls_frame,
            textvariable=self.colormap_var,
            values=self.COLORMAP_OPTIONS,
            state="readonly",
            width=20,
        ).grid(row=10, column=1, sticky="ew", pady=4)

        self.generate_btn = ttk.Button(controls_frame, text="Generate Preview", command=self.generate_wordcloud_thread)
        self.generate_btn.grid(row=11, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        stopwords_frame = ttk.LabelFrame(left_panel, text="Custom Stopwords", padding=10)
        stopwords_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        stopwords_frame.columnconfigure(0, weight=1)

        entry_row = ttk.Frame(stopwords_frame, style="Card.TFrame")
        entry_row.grid(row=0, column=0, sticky="ew")
        entry_row.columnconfigure(0, weight=1)

        ttk.Entry(entry_row, textvariable=self.stopword_entry_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(entry_row, text="Add", command=self.add_custom_stopwords).grid(row=0, column=1, sticky="e")

        list_frame = ttk.Frame(stopwords_frame, style="Card.TFrame")
        list_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        list_frame.columnconfigure(0, weight=1)

        self.stopwords_listbox = tk.Listbox(list_frame, height=6, exportselection=False)
        self.stopwords_listbox.grid(row=0, column=0, sticky="ew")
        stopwords_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.stopwords_listbox.yview)
        stopwords_scrollbar.grid(row=0, column=1, sticky="ns")
        self.stopwords_listbox.config(yscrollcommand=stopwords_scrollbar.set)

        stopword_btn_row = ttk.Frame(stopwords_frame, style="Card.TFrame")
        stopword_btn_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(stopword_btn_row, text="Remove Selected", command=self.remove_selected_stopwords).pack(side="left")
        ttk.Button(stopword_btn_row, text="Clear All", command=self.clear_custom_stopwords).pack(side="left", padx=(8, 0))
        ttk.Label(stopwords_frame, textvariable=self.stopword_count_var, style="TLabel").grid(row=3, column=0, sticky="w", pady=(8, 0))

        stats_frame = ttk.LabelFrame(left_panel, text="Quick Stats", padding=10)
        stats_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        stats_frame.columnconfigure(1, weight=1)
        left_panel.rowconfigure(2, weight=1)

        ttk.Label(stats_frame, text="Total rows:", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Label(stats_frame, textvariable=self.total_rows_var, style="TLabel").grid(row=0, column=1, sticky="e", pady=2)
        ttk.Label(stats_frame, text="Usable rows:", style="Section.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(stats_frame, textvariable=self.usable_rows_var, style="TLabel").grid(row=1, column=1, sticky="e", pady=2)
        ttk.Label(stats_frame, text="Unique terms:", style="Section.TLabel").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Label(stats_frame, textvariable=self.unique_terms_var, style="TLabel").grid(row=2, column=1, sticky="e", pady=2)
        ttk.Label(stats_frame, text="Term occurrences:", style="Section.TLabel").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Label(stats_frame, textvariable=self.term_occurrences_var, style="TLabel").grid(row=3, column=1, sticky="e", pady=2)

        ttk.Label(stats_frame, text="Top Terms", style="Section.TLabel").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 6))
        tree_frame = ttk.Frame(stats_frame, style="Card.TFrame")
        tree_frame.grid(row=5, column=0, columnspan=2, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(5, weight=1)

        self.terms_tree = ttk.Treeview(tree_frame, columns=("term", "count", "share"), show="headings", height=10)
        self.terms_tree.heading("term", text="Term")
        self.terms_tree.heading("count", text="Count")
        self.terms_tree.heading("share", text="Share")
        self.terms_tree.column("term", width=180, anchor="w")
        self.terms_tree.column("count", width=70, anchor="center")
        self.terms_tree.column("share", width=80, anchor="center")
        self.terms_tree.grid(row=0, column=0, sticky="nsew")
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.terms_tree.yview)
        tree_scrollbar.grid(row=0, column=1, sticky="ns")
        self.terms_tree.configure(yscrollcommand=tree_scrollbar.set)

        ttk.Label(right_panel, text="Preview", style="Title.TLabel").grid(row=0, column=0, sticky="w")

        preview_frame = ttk.Frame(right_panel, style="Card.TFrame")
        preview_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 12))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.preview_label = tk.Label(
            preview_frame,
            text="Generate a wordcloud to preview it here.",
            bg="#ffffff",
            fg="#555555",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 11),
            wraplength=620,
            justify="center",
        )
        self.preview_label.grid(row=0, column=0, sticky="nsew")

        action_row = ttk.Frame(right_panel, style="Card.TFrame")
        action_row.grid(row=2, column=0, sticky="ew")
        self.save_png_btn = ttk.Button(action_row, text="Save PNG", command=self.save_png, state="disabled")
        self.save_png_btn.pack(side="left")
        self.export_terms_btn = ttk.Button(action_row, text="Export Terms", command=self.export_terms, state="disabled")
        self.export_terms_btn.pack(side="left", padx=(8, 0))

        ttk.Label(outer, textvariable=self.status_var, style="TLabel").grid(row=2, column=0, sticky="w", pady=(8, 0))

    def add_custom_stopwords(self):
        raw_value = self.stopword_entry_var.get().strip()
        if not raw_value:
            return

        additions = [
            item.strip().lower()
            for item in raw_value.replace("\n", ",").split(",")
            if item.strip()
        ]
        self.custom_stopwords.update(additions)
        self.stopword_entry_var.set("")
        self._refresh_stopword_listbox()
        self.update_stopword_count()

    def remove_selected_stopwords(self):
        selected_indices = self.stopwords_listbox.curselection()
        if not selected_indices:
            return

        selected_words = [self.stopwords_listbox.get(index) for index in selected_indices]
        for word in selected_words:
            self.custom_stopwords.discard(word)

        self._refresh_stopword_listbox()
        self.update_stopword_count()

    def clear_custom_stopwords(self):
        self.custom_stopwords.clear()
        self._refresh_stopword_listbox()
        self.update_stopword_count()

    def update_stopword_count(self):
        try:
            config = self._build_config(validate_only=True)
            effective_count = len(get_effective_stopwords(config))
        except Exception:
            effective_count = len(self.custom_stopwords)
        self.stopword_count_var.set(f"Effective stopwords: {effective_count}")

    def generate_wordcloud_thread(self):
        if self.is_rendering:
            return

        try:
            config = self._build_config()
            column = self.column_var.get().strip()
            texts = coerce_text_column(self.app.df[column]).tolist()
        except Exception as e:
            messagebox.showerror("Invalid Wordcloud Settings", str(e))
            return

        self.is_rendering = True
        self.generate_btn.config(state="disabled")
        self.save_png_btn.config(state="disabled")
        self.export_terms_btn.config(state="disabled")
        self.status_var.set("Generating preview...")
        self.app.log_msg(f"Generating wordcloud for sheet '{self.app.sheet_var.get()}' and column '{column}'...")

        worker = threading.Thread(
            target=self._render_worker,
            args=(column, texts, config),
            daemon=True,
        )
        worker.start()

    def _render_worker(self, column, texts, config):
        try:
            stats_df, summary = prepare_wordcloud_data(texts, config)
            if stats_df.empty:
                self.app.master.after(0, lambda: self._finish_empty_render(column, summary))
                return

            image = render_wordcloud(stats_df, config)
            self.app.master.after(0, lambda: self._finish_render(column, stats_df, summary, image))
        except Exception as e:
            self.app.master.after(0, lambda error=str(e): self._finish_render_error(error))

    def _finish_render(self, column, stats_df, summary, image):
        if not self.is_alive:
            return

        self.current_stats_df = stats_df
        self.current_image = image
        self._update_summary(summary)
        self._populate_terms_table(stats_df)
        self._update_preview_image(image)
        self.is_rendering = False
        self.generate_btn.config(state="normal")
        self.save_png_btn.config(state="normal")
        self.export_terms_btn.config(state="normal")
        self.status_var.set(f"Preview ready for column '{column}'.")
        self.app.log_msg(f"✓ Wordcloud ready: {len(stats_df)} filtered terms from column '{column}'")

    def _finish_empty_render(self, column, summary):
        if not self.is_alive:
            return

        self._update_summary(summary)
        self._populate_terms_table(None)
        self._reset_preview_state("No terms remained after applying the current filters.")
        self.is_rendering = False
        self.generate_btn.config(state="normal")
        self.status_var.set("No preview generated because no terms remained after filtering.")
        self.app.log_msg(f"✗ Wordcloud skipped for column '{column}': no terms remained after filtering")
        messagebox.showwarning(
            "No Terms Available",
            "The selected column is empty after applying the current filters. Adjust the controls and try again.",
        )

    def _finish_render_error(self, error_message):
        if not self.is_alive:
            return

        self.is_rendering = False
        self.generate_btn.config(state="normal")
        self.status_var.set("Preview failed. See the error details and try again.")
        self.app.log_msg(f"✗ Wordcloud generation failed: {error_message}")
        messagebox.showerror("Wordcloud Generation Failed", error_message)

    def save_png(self):
        if self.current_image is None:
            messagebox.showwarning("No Preview", "Generate a wordcloud preview first")
            return

        default_name = self._default_export_stem() + ".png"
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png")],
            initialfile=default_name,
            title="Save Wordcloud PNG",
        )
        if not path:
            return

        try:
            self.current_image.save(path, format="PNG")
            self.app.log_msg(f"✓ Wordcloud image saved to {path}")
            messagebox.showinfo("Saved", f"Saved wordcloud image to {path}")
        except Exception as e:
            self.app.log_msg(f"✗ Wordcloud image save failed: {str(e)}")
            messagebox.showerror("Save Failed", f"Failed to save wordcloud image: {e}")

    def export_terms(self):
        if self.current_stats_df is None or self.current_stats_df.empty:
            messagebox.showwarning("No Terms", "Generate a wordcloud preview first")
            return

        default_name = self._default_export_stem() + "_terms.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="Export Term Statistics",
        )
        if not path:
            return

        try:
            saved_path = export_term_stats(self.current_stats_df, path)
            self.app.log_msg(f"✓ Wordcloud terms exported to {saved_path}")
            messagebox.showinfo("Exported", f"Exported wordcloud terms to {saved_path}")
        except Exception as e:
            self.app.log_msg(f"✗ Term export failed: {str(e)}")
            messagebox.showerror("Export Failed", f"Failed to export term statistics: {e}")

    def _build_config(self, validate_only=False):
        if self.app.df is None:
            raise ValueError("Load a sheet before opening the wordcloud builder.")

        column = self.column_var.get().strip()
        if not column:
            raise ValueError("Select a source column for the wordcloud.")
        if column not in self.app.df.columns:
            raise ValueError(f"Column '{column}' is no longer available in the active sheet.")

        config = WordCloudConfig(
            max_words=int(self.max_words_var.get().strip() or "200"),
            min_frequency=int(self.min_frequency_var.get().strip() or "1"),
            width=int(self.width_var.get().strip() or "1200"),
            height=int(self.height_var.get().strip() or "700"),
            phrase_mode=self.phrase_mode_var.get().strip() or self.PHRASE_OPTIONS[0],
            use_builtin_stopwords=bool(self.use_builtin_stopwords_var.get()),
            lowercase=bool(self.lowercase_var.get()),
            exclude_numeric=bool(self.exclude_numeric_var.get()),
            background_color=self.background_var.get().strip() or "white",
            colormap=self.colormap_var.get().strip() or "viridis",
            custom_stopwords=set(self.custom_stopwords),
        )

        if not validate_only:
            self.context_var.set(
                f"File: {os.path.basename(self.app.current_file_path)}    "
                f"Sheet: {self.app.sheet_var.get().strip() or 'No sheet'}    "
                f"Column: {column}"
            )
        return config

    def _refresh_stopword_listbox(self):
        self.stopwords_listbox.delete(0, tk.END)
        for word in sorted(self.custom_stopwords):
            self.stopwords_listbox.insert(tk.END, word)

    def _update_summary(self, summary):
        self.total_rows_var.set(str(summary.get("total_rows", 0)))
        self.usable_rows_var.set(str(summary.get("usable_rows", 0)))
        self.unique_terms_var.set(str(summary.get("unique_terms", 0)))
        self.term_occurrences_var.set(str(summary.get("kept_term_occurrences", 0)))

    def _populate_terms_table(self, stats_df):
        for item_id in self.terms_tree.get_children():
            self.terms_tree.delete(item_id)

        if stats_df is None or stats_df.empty:
            return

        for _index, row in stats_df.head(10).iterrows():
            self.terms_tree.insert("", "end", values=(row["term"], int(row["count"]), f"{row['share']:.1%}"))

    def _update_preview_image(self, image):
        preview_image = image.copy()
        preview_image.thumbnail((760, 560))
        self.preview_photo = ImageTk.PhotoImage(preview_image)
        self.preview_label.config(image=self.preview_photo, text="", bg="#ffffff")

    def _reset_preview_state(self, message, clear_summary=False):
        self.current_stats_df = None
        self.current_image = None
        self.preview_photo = None
        self.preview_label.config(image="", text=message, bg="#ffffff")
        self.save_png_btn.config(state="disabled")
        self.export_terms_btn.config(state="disabled")
        self._populate_terms_table(None)
        if clear_summary:
            self._update_summary({})

    def _default_export_stem(self):
        base_name = os.path.splitext(os.path.basename(self.app.current_file_path or "wordcloud"))[0]
        sheet_name = self._slugify(self.app.sheet_var.get().strip() or "sheet")
        column_name = self._slugify(self.column_var.get().strip() or "column")
        return f"{base_name}_{sheet_name}_{column_name}_wordcloud"

    def _slugify(self, value):
        characters = [char.lower() if char.isalnum() else "_" for char in value]
        slug = "".join(characters).strip("_")
        while "__" in slug:
            slug = slug.replace("__", "_")
        return slug or "item"


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
