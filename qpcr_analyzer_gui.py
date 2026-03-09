import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import json

class QPCRAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("qPCR ddCT Analysis Framework")
        self.root.geometry("1000x900")
        
        # Make the layout expand nicely
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Use a more colorful theme
        style = ttk.Style(self.root)
        try:
            style.theme_use('radiance') # 'radiance' is a colorful Ubuntu-like theme
        except tk.TclError:
            style.theme_use('clam')
        
        # Create a padded main container
        self.main_container = ttk.Frame(self.root, padding="15 15 15 15")
        self.main_container.pack(fill="both", expand=True)
        
        self.df = None
        self.omissions = set() # Store (target, sample) tuples to omit
        
        self.setup_ui()
        
    def setup_ui(self):
        # Frame 0: Config
        frame_config = ttk.Frame(self.main_container)
        frame_config.pack(fill="x", pady=(0, 10))
        ttk.Button(frame_config, text="Load Config Template", command=self.load_config).pack(side="left", padx=(0, 5))
        ttk.Button(frame_config, text="Save Config Template", command=self.save_config).pack(side="left", padx=5)

        ttk.Label(frame_config, text="App Theme:").pack(side="left", padx=(20, 5))
        self.app_theme_var = tk.StringVar(value="radiance")
        self.cb_app_theme = ttk.Combobox(frame_config, textvariable=self.app_theme_var, state="readonly", width=12)
        self.cb_app_theme['values'] = ["arc", "radiance", "equilux", "breeze", "plastik", "clearlooks"]
        self.cb_app_theme.pack(side="left")
        self.cb_app_theme.bind("<<ComboboxSelected>>", self.change_app_theme)

        ttk.Label(frame_config, text="Plot Style:").pack(side="left", padx=(15, 5))
        self.plot_style_var = tk.StringVar(value="ggplot")
        self.cb_plot_style = ttk.Combobox(frame_config, textvariable=self.plot_style_var, state="readonly", width=15)
        self.cb_plot_style['values'] = ["ggplot", "seaborn", "fivethirtyeight", "classic", "bmh", "default"]
        self.cb_plot_style.pack(side="left")

        # Frame 1: File Selection
        frame_file = ttk.LabelFrame(self.main_container, text="Step 1: Data Input")
        frame_file.pack(fill="x", pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(frame_file, textvariable=self.file_path_var, state="readonly", width=70).pack(side="left", padx=5, pady=5)
        ttk.Button(frame_file, text="Load Excel Data", command=self.load_file).pack(side="left", padx=5, pady=5)
        
        # Frame 2: Column Mapping
        frame_mapping = ttk.LabelFrame(self.main_container, text="Step 2: Column Mapping")
        frame_mapping.pack(fill="x", pady=5)
        
        ttk.Label(frame_mapping, text="Sample Name Col:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.col_sample_var = tk.StringVar()
        self.cb_sample = ttk.Combobox(frame_mapping, textvariable=self.col_sample_var, state="readonly", width=20)
        self.cb_sample.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_mapping, text="Target Name Col:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.col_target_var = tk.StringVar()
        self.cb_target = ttk.Combobox(frame_mapping, textvariable=self.col_target_var, state="readonly", width=20)
        self.cb_target.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(frame_mapping, text="CT Value Col:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        self.col_ct_var = tk.StringVar()
        self.cb_ct = ttk.Combobox(frame_mapping, textvariable=self.col_ct_var, state="readonly", width=15)
        self.cb_ct.grid(row=0, column=5, padx=5, pady=5)
        
        # Grouping optionally by Custom Task
        ttk.Label(frame_mapping, text="Custom Task Col (Opt):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.col_task_var = tk.StringVar()
        self.cb_task = ttk.Combobox(frame_mapping, textvariable=self.col_task_var, state="readonly", width=20)
        self.cb_task.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(frame_mapping, text="Apply Mapping", command=self.apply_mapping).grid(row=1, column=4, columnspan=2, pady=5)
        
        # Frame 3: Control Selection
        frame_controls = ttk.LabelFrame(self.main_container, text="Step 3: Analyze Controls")
        frame_controls.pack(fill="both", expand=True, pady=5)
        
        # Bulk Setter
        frame_global = ttk.Frame(frame_controls)
        frame_global.pack(fill="x", padx=5, pady=5)
        ttk.Label(frame_global, text="Bulk Set All - Target Ctrl:").pack(side="left")
        self.bulk_tgt_var = tk.StringVar()
        self.cb_bulk_tgt = ttk.Combobox(frame_global, textvariable=self.bulk_tgt_var, state="readonly", width=15)
        self.cb_bulk_tgt.pack(side="left", padx=2)
        
        ttk.Label(frame_global, text="Sample Ctrl:").pack(side="left")
        self.bulk_smp_var = tk.StringVar()
        self.cb_bulk_smp = ttk.Combobox(frame_global, textvariable=self.bulk_smp_var, state="readonly", width=15)
        self.cb_bulk_smp.pack(side="left", padx=2)
        
        ttk.Button(frame_global, text="Apply to All", command=self.apply_bulk_controls).pack(side="left", padx=5)

        # Scrollable area for per-target
        self.canvas = tk.Canvas(frame_controls)
        self.scrollbar = ttk.Scrollbar(frame_controls, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.target_control_vars = {}
        self.sample_control_vars = {}
        
        # Frame 4: Action Output
        frame_actions = ttk.Frame(self.main_container)
        frame_actions.pack(fill="x", pady=10)
        
        ttk.Button(frame_actions, text="Select Omissions", command=self.open_omissions_window, width=20).pack(side="left", padx=(0, 5))
        ttk.Button(frame_actions, text="Analyze & Generate Output", command=self.run_analysis, width=40).pack(side="left", padx=5)
        
        # Info text
        self.text_info = tk.Text(self.main_container, height=12, font=("Consolas", 10), bg="#f5f5f5", relief="flat")
        self.text_info.pack(fill="both", expand=True, pady=(5, 0))
        self.log("Waiting for data...")

    def change_app_theme(self, event=None):
        theme = self.app_theme_var.get()
        style = ttk.Style(self.root)
        try:
            if hasattr(self.root, 'set_theme'):
                self.root.set_theme(theme)
            else:
                style.theme_use(theme)
        except tk.TclError:
            pass

    def log(self, msg):
        self.text_info.insert("end", msg + "\n")
        self.text_info.see("end")

    def save_config(self):
        config_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if not config_path:
            return
            
        data = {
            "mappings": {
                "sample_col": self.col_sample_var.get(),
                "target_col": self.col_target_var.get(),
                "ct_col": self.col_ct_var.get(),
                "task_col": self.col_task_var.get()
            },
            "controls": {
                "targets": {tg: var.get() for tg, var in self.target_control_vars.items()},
                "samples": {tg: var.get() for tg, var in self.sample_control_vars.items()}
            },
            "omissions": list(self.omissions)
        }
        
        try:
            with open(config_path, 'w') as f:
                json.dump(data, f, indent=4)
            self.log(f"Config saved to {config_path}")
            messagebox.showinfo("Success", "Configuration template saved!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {e}")

    def load_config(self):
        config_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not config_path:
            return
            
        try:
            with open(config_path, 'r') as f:
                data = json.load(f)
                
            mappings = data.get("mappings", {})
            self.col_sample_var.set(mappings.get("sample_col", ""))
            self.col_target_var.set(mappings.get("target_col", ""))
            self.col_ct_var.set(mappings.get("ct_col", ""))
            self.col_task_var.set(mappings.get("task_col", ""))
            self.omissions = set(tuple(x) for x in data.get("omissions", []))
            
            # Re-apply mappings to trigger the target grid generation
            if self.df is not None:
                self.apply_mapping()
                
                # Now that the grid is built, populate the saved controls
                controls = data.get("controls", {})
                tgt_ctrls = controls.get("targets", {})
                smp_ctrls = controls.get("samples", {})
                
                for tg, var in self.target_control_vars.items():
                    if tg in tgt_ctrls:
                        var.set(tgt_ctrls[tg])
                
                for tg, var in self.sample_control_vars.items():
                    if tg in smp_ctrls:
                        var.set(smp_ctrls[tg])
                        
            self.log(f"Config loaded from {config_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {e}")

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.file_path_var.set(filepath)
            try:
                self.df = pd.read_excel(filepath)
                cols = [""] + self.df.columns.tolist()
                
                self.cb_sample['values'] = cols
                self.cb_target['values'] = cols
                self.cb_ct['values'] = cols
                self.cb_task['values'] = cols
                
                # Auto-detect common names
                col_upper = [c.upper() for c in cols]
                if "SAMPLE NAME" in col_upper: self.cb_sample.set(cols[col_upper.index("SAMPLE NAME")])
                elif "SAMPLE" in col_upper: self.cb_sample.set(cols[col_upper.index("SAMPLE")])
                    
                if "TARGET NAME" in col_upper: self.cb_target.set(cols[col_upper.index("TARGET NAME")])
                elif "TARGET" in col_upper: self.cb_target.set(cols[col_upper.index("TARGET")])
                    
                if "CT" in col_upper: self.cb_ct.set(cols[col_upper.index("CT")])
                
                self.log(f"Successfully loaded file with columns: {self.df.columns.tolist()}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not read excel file: {e}")
                
    def apply_mapping(self):
        if self.df is None: return
        
        scol = self.col_sample_var.get()
        tcol = self.col_target_var.get()
        ctcol = self.col_ct_var.get()
        
        if not scol or not tcol or not ctcol:
            messagebox.showerror("Error", "Please map Sample, Target, and CT columns.")
            return
            
        try:
            samples = sorted(self.df[scol].dropna().astype(str).unique())
            targets = sorted(self.df[tcol].dropna().astype(str).unique())
            
            self.cb_bulk_tgt['values'] = targets
            self.cb_bulk_smp['values'] = samples
            
            # Clear scrollable frame
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
                
            self.target_control_vars = {}
            self.sample_control_vars = {}
            
            # Add header
            ttk.Label(self.scrollable_frame, text="Target", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
            ttk.Label(self.scrollable_frame, text="Target Control (Ref Gene)", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=2, sticky="w")
            ttk.Label(self.scrollable_frame, text="Sample Control (Ref Sample)", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=2, sticky="w")

            for i, tg in enumerate(targets, start=1):
                ttk.Label(self.scrollable_frame, text=tg).grid(row=i, column=0, padx=5, pady=2, sticky="w")
                
                tgt_var = tk.StringVar()
                cb_tgt = ttk.Combobox(self.scrollable_frame, textvariable=tgt_var, state="readonly", values=targets, width=20)
                cb_tgt.grid(row=i, column=1, padx=5, pady=2)
                self.target_control_vars[tg] = tgt_var
                
                smp_var = tk.StringVar()
                cb_smp = ttk.Combobox(self.scrollable_frame, textvariable=smp_var, state="readonly", values=samples, width=20)
                cb_smp.grid(row=i, column=2, padx=5, pady=2)
                self.sample_control_vars[tg] = smp_var
            
            self.log(f"Found {len(samples)} unique samples and {len(targets)} unique targets.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            
    def apply_bulk_controls(self):
        bt = self.bulk_tgt_var.get()
        bs = self.bulk_smp_var.get()
        
        for tg, var in self.target_control_vars.items():
            if bt: var.set(bt)
        for tg, var in self.sample_control_vars.items():
            if bs: var.set(bs)

    def open_omissions_window(self):
        if self.df is None:
            messagebox.showerror("Error", "Please load data and apply mapping first.")
            return
            
        scol = self.col_sample_var.get()
        tcol = self.col_target_var.get()
        taskcol = self.col_task_var.get()
        if not scol or not tcol:
            messagebox.showerror("Error", "Please apply mapping first.")
            return
            
        # Get unique combinations of Task (if mapped) and Sample
        if taskcol:
            df_cleaned = self.df.dropna(subset=[scol, tcol, taskcol])
            # Create a label for columns combining Task and Sample
            tasks = sorted(df_cleaned[taskcol].astype(str).unique())
            samples = sorted(df_cleaned[scol].astype(str).unique())
            col_labels = [f"{t} - {s}" for t in tasks for s in samples]
            # Mapping from col index to (task, sample)
            col_mapping = [(t, s) for t in tasks for s in samples]
        else:
            df_cleaned = self.df.dropna(subset=[scol, tcol])
            samples = sorted(df_cleaned[scol].astype(str).unique())
            col_labels = samples
            col_mapping = [(None, s) for s in samples]
            
        targets = sorted(df_cleaned[tcol].astype(str).unique())
        
        win = tk.Toplevel(self.root)
        win.title("Select Target-Sample Analysis Omissions")
        win.geometry("1000x600")
        
        canvas = tk.Canvas(win)
        scrollbar = ttk.Scrollbar(win, orient="horizontal", command=canvas.xview)
        vscrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(xscrollcommand=scrollbar.set, yscrollcommand=vscrollbar.set)
        
        btn_frame = ttk.Frame(win, padding="10")
        btn_frame.pack(side="bottom", fill="x")
        ttk.Button(btn_frame, text="Confirm & Close", command=win.destroy, width=20).pack()
        
        vscrollbar.pack(side="right", fill="y")
        scrollbar.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)
        
        ttk.Label(scrollable_frame, text="Check boxes to OMIT the combination from analysis", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=len(col_labels)+1, pady=10, sticky="w")
        
        # Header Row
        for j, label in enumerate(col_labels):
            ttk.Label(scrollable_frame, text=label, font=("Arial", 9, "bold")).grid(row=1, column=j+1, padx=2, sticky="w")
            
        for i, tgt in enumerate(targets):
            ttk.Label(scrollable_frame, text=tgt, font=("Arial", 9, "bold")).grid(row=i+2, column=0, padx=5, sticky="e")
            for j, (tsk, smp) in enumerate(col_mapping):
                # The tuple stored in omissions is either (Target, Sample) or (Target, Sample, Task) depending on task presence
                om_tuple = (tgt, smp, tsk) if taskcol else (tgt, smp)
                
                var = tk.BooleanVar(value=om_tuple in self.omissions)
                cb = ttk.Checkbutton(scrollable_frame, variable=var)
                cb.grid(row=i+2, column=j+1, padx=2)
                
                def make_cmd(tup=om_tuple, v=var):
                    return lambda: self.omissions.add(tup) if v.get() else self.omissions.discard(tup)
                cb.config(command=make_cmd())

    def run_analysis(self):
        if self.df is None: return
        
        scol = self.col_sample_var.get()
        tcol = self.col_target_var.get()
        ctcol = self.col_ct_var.get()
        taskcol = self.col_task_var.get()
        
        has_controls = any(v.get() for v in self.target_control_vars.values())
        if not has_controls:
            messagebox.showerror("Error", "Please set controls for at least one target.")
            return
            
        try:
            self.log("Starting analysis...")
            df_work = self.df.copy()
            
            # 1. Clean data: Treat 'Undetermined' as NaN, drop rows with NaN CT if any
            df_work[ctcol] = pd.to_numeric(df_work[ctcol], errors='coerce')
            df_work = df_work.dropna(subset=[scol, tcol, ctcol])
            
            # Apply manual Omissions
            if self.omissions:
                if taskcol:
                    mask = ~df_work.apply(lambda row: (str(row[tcol]), str(row[scol]), str(row[taskcol])) in self.omissions, axis=1)
                else:
                    mask = ~df_work.apply(lambda row: (str(row[tcol]), str(row[scol])) in self.omissions, axis=1)
                    
                dropped = len(df_work) - sum(mask)
                df_work = df_work[mask]
                if dropped > 0:
                    self.log(f"Manually omitted {dropped} data points matching selected combinations.")
            
            # Setup columns to group by
            group_cols = [scol, tcol]
            if taskcol:
                group_cols.append(taskcol)
            
            # 1.5 Auto Outlier Removal for small replicates
            # Using Modified Z-score based on Median Absolute Deviation (MAD)
            # which handles n=3 groups much better than IQR
            
            count_col = df_work.groupby(group_cols)[ctcol].transform('count')
            median_col = df_work.groupby(group_cols)[ctcol].transform('median')
            
            # Calculate absolute deviation from median
            abs_dev = (df_work[ctcol] - median_col).abs()
            
            # Calculate MAD for the group
            mad_col = abs_dev.groupby(df_work[group_cols].apply(tuple, axis=1)).transform('median')
            
            # modified z-score = 0.6745 * (x - median) / MAD
            # if MAD is 0 (e.g. all replicates are identical), set z-score to 0
            mod_z = np.where(mad_col == 0, 0, 0.6745 * abs_dev / mad_col)
            
            # Alternatively, for biological triplicates, a simple distance threshold vs median works well
            # if the max distance to median is > 1.5 CTs for example.
            # But let's use a dynamic approach: if n=3, the "odd one out" will pull the mean but not the median
            # We flag points that are > 1.0 CT away from the median AND are the furthest point.
            
            def flag_triplicate_outliers(group):
                if len(group) < 3:
                    return pd.Series(True, index=group.index)
                
                med = group[ctcol].median()
                dists = (group[ctcol] - med).abs()
                
                # If the max deviation is > 1.0 CT from the median, drop the furthest one
                # You can adjust 1.0 based on typical qPCR variance. 
                # e.g., for 13.9, 14.4, 31.4 -> median is 14.4. 31.4 is 17 units away > 1.0.
                if dists.max() > 1.0:
                    # Keep everything EXCEPT the single furthest point
                    is_max = dists == dists.max()
                    # if there's a tie for furthest, we drop the first to be safe, or neither. 
                    # Usually, outliers are uniquely far.
                    return ~is_max
                
                return pd.Series(True, index=group.index)
                
            # Apply the custom triplicate logic
            valid_mask = df_work.groupby(group_cols, group_keys=False).apply(flag_triplicate_outliers)
            
            # Alignment fix in case index gets messed up (apply can sometimes return out of order)
            valid_mask = valid_mask.loc[df_work.index]
            
            df_work_clean = df_work[valid_mask].copy()
            
            # Store dropped info for logging later if desired
            dropped_count = len(df_work) - len(df_work_clean)
            if dropped_count > 0:
                self.log(f"Auto-removed {dropped_count} outlier replicate(s).")
            
            # 2. Mean & SD of CT replicates
            agg_df = df_work_clean.groupby(group_cols)[ctcol].agg(
                CT_Mean='mean',
                CT_SD='std'
            ).reset_index()
            # If standard deviation is NaN (e.g. only 1 replicate), fill with 0
            agg_df['CT_SD'] = agg_df['CT_SD'].fillna(0)
            
            results_list = []
            
            for tg, var in self.target_control_vars.items():
                c_tg = var.get()
                c_smp = self.sample_control_vars[tg].get()
                
                if not c_tg or not c_smp:
                    continue
                    
                tg_df = agg_df[agg_df[tcol] == tg].copy()
                if tg_df.empty: continue
                
                # 3. Calculate dCT
                ctrl_tgt_df = agg_df[agg_df[tcol] == c_tg].copy()
                ctrl_tgt_df = ctrl_tgt_df.rename(columns={'CT_Mean': 'CT_Mean_CtrlTgt', 'CT_SD': 'CT_SD_CtrlTgt'})
                
                merge_cols = [scol]
                if taskcol: merge_cols.append(taskcol)
                
                tg_df = tg_df.merge(ctrl_tgt_df[merge_cols + ['CT_Mean_CtrlTgt', 'CT_SD_CtrlTgt']], on=merge_cols, how='left')
                
                tg_df['dCT'] = tg_df['CT_Mean'] - tg_df['CT_Mean_CtrlTgt']
                tg_df['dCT_SD'] = np.sqrt(tg_df['CT_SD']**2 + tg_df['CT_SD_CtrlTgt']**2)
                
                # 4. Calculate ddCT
                ctrl_smp_mask = (tg_df[scol] == c_smp)
                ctrl_smp_df = tg_df[ctrl_smp_mask].copy()
                ctrl_smp_df = ctrl_smp_df.rename(columns={'dCT': 'dCT_CtrlSmp', 'dCT_SD': 'dCT_SD_CtrlSmp'})
                
                merge_cols_smp = [tcol]
                if taskcol: merge_cols_smp.append(taskcol)
                
                tg_df = tg_df.merge(ctrl_smp_df[merge_cols_smp + ['dCT_CtrlSmp', 'dCT_SD_CtrlSmp']], on=merge_cols_smp, how='left')
                
                tg_df['ddCT'] = tg_df['dCT'] - tg_df['dCT_CtrlSmp']
                tg_df['ddCT_SD'] = np.sqrt(tg_df['dCT_SD']**2 + tg_df['dCT_SD_CtrlSmp']**2)
                
                # 5. Calculate RQ
                tg_df['RQ'] = 2 ** (-tg_df['ddCT'])
                tg_df['RQ_min'] = 2 ** (-(tg_df['ddCT'] + tg_df['ddCT_SD']))
                tg_df['RQ_max'] = 2 ** (-(tg_df['ddCT'] - tg_df['ddCT_SD']))
                
                tg_df['RQ_error_down'] = tg_df['RQ'] - tg_df['RQ_min']
                tg_df['RQ_error_up'] = tg_df['RQ_max'] - tg_df['RQ']
                
                results_list.append(tg_df)
                
            if not results_list:
                messagebox.showerror("Error", "No valid targets selected for analysis.")
                return
                
            result_df = pd.concat(results_list, ignore_index=True)
            
            # Filter out target controls so they aren't plotted
            to_keep = []
            for _, row in result_df.iterrows():
                tg = row[tcol]
                c_tg = self.target_control_vars[tg].get()
                to_keep.append(tg != c_tg)
                
            result_df = result_df[to_keep]
            
            # Save TSV
            output_dir = os.path.dirname(self.file_path_var.get())
            base_name = os.path.splitext(os.path.basename(self.file_path_var.get()))[0]
            out_tsv = os.path.join(output_dir, f"{base_name}_analyzed.tsv")
            result_df.to_csv(out_tsv, sep='\t', index=False)
            self.log(f"Saved analysis TSV to: {out_tsv}")
            
            # Generate Plots
            self.generate_plots(result_df, scol, tcol, taskcol, output_dir, base_name)
            
            messagebox.showinfo("Success", "Analysis complete! Results and plots saved.")
            
        except Exception as e:
            self.log(f"Analysis Error: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Analysis failed: {str(e)}")

    def generate_plots(self, df, scol, tcol, taskcol, output_dir, base_name):
        # Exclude the reference sample from the plot usually (RQ = 1, Error = 0)
        to_keep = []
        for _, row in df.iterrows():
            tg = row[tcol]
            c_smp = self.sample_control_vars[tg].get()
            to_keep.append(row[scol] != c_smp)
            
        plot_df = df[to_keep].copy()
        
        if len(plot_df) == 0:
            self.log("Warning: No data to plot after removing control sample/target!")
            return
            
        if taskcol:
            plot_df['Plot_Group'] = plot_df[taskcol].astype(str) + "\n" + plot_df[scol].astype(str)
            group_col = 'Plot_Group'
        else:
            group_col = scol
            
        filename = f"{base_name}_RQ_Combined_Barchart.png"
        title = "RQ Barchart (All Tasks)" if taskcol else "RQ Barchart"
        
        groups = []
        if taskcol:
            # Sort by task, then sample
            sorted_df = plot_df[['Plot_Group', taskcol, scol]].drop_duplicates().sort_values([taskcol, scol])
            groups = sorted_df['Plot_Group'].tolist()
        else:
            groups = sorted(plot_df[group_col].unique())
            
        targets = sorted(plot_df[tcol].unique())
        
        x = np.arange(len(groups))
        width = 0.8 / len(targets)
        
        # Apply select plot style
        plot_style = self.plot_style_var.get()
        try:
            # seaborn-v0_8 in newer matplotlib but 'seaborn' historically
            if plot_style == 'seaborn' and 'seaborn-v0_8' in plt.style.available:
                plt.style.use('seaborn-v0_8')
            else:
                plt.style.use(plot_style)
        except Exception:
            plt.style.use('ggplot')
        
        # Make figure wider if there are many groups
        fig_width = max(10, len(groups) * len(targets) * 0.4)
        fig, ax = plt.subplots(figsize=(fig_width, 6))
        
        for i, target in enumerate(targets):
            target_df = plot_df[plot_df[tcol] == target].set_index(group_col)
            
            # Align data to groups
            y_vals = []
            yerr_lower = []
            yerr_upper = []
            
            for g in groups:
                if g in target_df.index:
                    row = target_df.loc[g]
                    y_vals.append(row['RQ'])
                    yerr_lower.append(row['RQ_error_down'])
                    yerr_upper.append(row['RQ_error_up'])
                else:
                    y_vals.append(0)
                    yerr_lower.append(0)
                    yerr_upper.append(0)
                    
            offset = (i - len(targets)/2) * width + width/2
            ax.bar(x + offset, y_vals, width, label=target, 
                   yerr=[yerr_lower, yerr_upper], capsize=5)
        
        ax.set_ylabel(r'Relative Quantification ($2^{-\Delta\Delta C_T}$)')
        # ax.set_title(title) # Removed as per user request
        ax.set_xticks(x)
        ax.set_xticklabels(groups, rotation=45 if not taskcol else 0, ha='center')
        ax.legend(title=tcol, bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # Baseline RQ=1
        ax.axhline(y=1.0, color='r', linestyle='--', alpha=0.3)
        
        plt.tight_layout()
        out_png = os.path.join(output_dir, filename)
        plt.savefig(out_png, dpi=300, bbox_inches='tight')
        plt.close()
        self.log(f"Saved combined barchart to: {out_png}")

if __name__ == "__main__":
    try:
        root = ThemedTk(theme="radiance")
    except Exception:
        root = tk.Tk()
        ttk.Style().theme_use('clam')
        
    app = QPCRAnalyzerApp(root)
    root.mainloop()
