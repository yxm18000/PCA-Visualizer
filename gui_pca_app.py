import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import mplcursors
import sys
import ctypes

# --- GUIアプリケーションのクラス定義 ---
class PcaApp:
    def __init__(self, root):
        # --- ダークモード用カラーテーマ ---
        self.BG_COLOR = "#2E2E2E"
        self.TEXT_COLOR = "#EAEAEA"
        self.FRAME_COLOR = "#3C3C3C"
        self.BUTTON_COLOR = "#555555"
        self.BUTTON_ACTIVE_COLOR = "#6A6A6A"
        self.LABEL_BG_COLOR = "#383838"

        self.root = root
        self.root.title("PCA Visualizer")
        self.root.geometry("1440x810")
        self.root.configure(bg=self.BG_COLOR)

        self.file_path = None
        self.pca = None
        self.pc_df = None
        self.df = None
        self.numerical_df_columns = None
        self.style_widgets = {}

        # ★★★ 変更点: 色の選択肢に 'white' と 'black' を追加 ★★★
        self.matplotlib_colors = ['blue', 'red', 'green', 'purple', 'orange', 'cyan', 'magenta', 'brown', 'gold', 'teal', 'white', 'black']
        self.matplotlib_markers = ['o', 's', '^', 'D', 'v', '*', 'p', 'X', '+', 'H']

        self.set_dark_title_bar()
        self.setup_ui()
        self.apply_dark_theme_to_tabs()

    def setup_ui(self):
        # --- 1. トップコントロールフレーム ---
        top_control_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        top_control_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        button_style = {
            "font": ("Arial", 10), "bg": self.BUTTON_COLOR, "fg": self.TEXT_COLOR,
            "activebackground": self.BUTTON_ACTIVE_COLOR, "activeforeground": self.TEXT_COLOR,
            "relief": tk.RAISED, "bd": 2, "padx": 10, "pady": 5
        }

        select_button = tk.Button(top_control_frame, text="1. CSVファイルを選択", command=self.select_file, **button_style)
        select_button.pack(side=tk.LEFT, padx=5)

        self.path_label = tk.Label(
            top_control_frame, text="ファイルが選択されていません", relief=tk.SUNKEN, bd=2,
            padx=5, anchor='w', bg=self.LABEL_BG_COLOR, fg=self.TEXT_COLOR, font=("Arial", 9)
        )
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        run_button = tk.Button(top_control_frame, text="2. 分析実行", command=self.run_analysis, **button_style)
        run_button.pack(side=tk.LEFT, padx=5)
        
        self.export_button = tk.Button(top_control_frame, text="3. 結果をエクスポート", command=self.export_results, **button_style, state=tk.DISABLED)
        self.export_button.pack(side=tk.LEFT, padx=5)


        # --- 2. メインコンテンツフレーム ---
        main_content_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        main_content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # --- 3. 左側: スタイル設定フレーム ---
        style_outer_frame = tk.LabelFrame(
            main_content_frame, text="スタイル設定", bg=self.BG_COLOR, fg=self.TEXT_COLOR,
            padx=10, pady=10, font=("Arial", 10, "bold")
        )
        style_outer_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        style_canvas = tk.Canvas(style_outer_frame, bg=self.FRAME_COLOR, highlightthickness=0, width=300)
        scrollbar = ttk.Scrollbar(style_outer_frame, orient="vertical", command=style_canvas.yview)
        self.style_inner_frame = tk.Frame(style_canvas, bg=self.FRAME_COLOR)
        self.style_inner_frame.bind("<Configure>", lambda e: style_canvas.configure(scrollregion=style_canvas.bbox("all")))
        style_canvas.create_window((0, 0), window=self.style_inner_frame, anchor="nw")
        style_canvas.configure(yscrollcommand=scrollbar.set)
        style_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.initial_style_message = tk.Label(self.style_inner_frame, text="CSVファイルを読み込むと\n設定項目が表示されます。",
                                              bg=self.FRAME_COLOR, fg=self.TEXT_COLOR, justify=tk.CENTER)
        self.initial_style_message.pack(pady=20, padx=10)

        # --- 4. 右側: グラフ表示用Notebook ---
        self.notebook = ttk.Notebook(main_content_frame)
        self.notebook.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        initial_tab = tk.Frame(self.notebook, bg=self.FRAME_COLOR)
        self.notebook.add(initial_tab, text="Plots")
        tk.Label(initial_tab, text="分析を実行するとここにグラフが表示されます。", bg=self.FRAME_COLOR, fg=self.TEXT_COLOR, font=("Arial", 12)).pack(expand=True)


    def apply_dark_theme_to_tabs(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure("TNotebook", background=self.BG_COLOR, borderwidth=0)
        style.configure("TNotebook.Tab", background=self.BUTTON_COLOR, foreground=self.TEXT_COLOR, padding=[10, 5], font=("Arial", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", self.FRAME_COLOR), ("!selected", self.BUTTON_COLOR)],
                  foreground=[("selected", self.TEXT_COLOR), ("!selected", self.TEXT_COLOR)])
        style.configure("TFrame", background=self.FRAME_COLOR)

    def set_dark_title_bar(self):
        if sys.platform == "win32":
            try:
                DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                value = 2
                hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(ctypes.c_int(value)), ctypes.sizeof(ctypes.c_int))
                self.root.withdraw()
                self.root.deiconify()
            except Exception as e:
                print(f"タイトルバーのダークモード設定に失敗しました: {e}")

    def select_file(self):
        path = filedialog.askopenfilename(title="CSVファイルを選択してください", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self.file_path = path.strip(' "')
            display_name = self.file_path.split('/')[-1]
            self.path_label.config(text=display_name)
            self._setup_style_ui()
            self.export_button.config(state=tk.DISABLED)

    def run_analysis(self):
        if not self.file_path:
            messagebox.showwarning("警告", "先にCSVファイルを選択してください。")
            return
        
        self.export_button.config(state=tk.DISABLED)

        try:
            self.df = pd.read_csv(self.file_path)
            if self.df.empty:
                messagebox.showerror("データエラー", "CSVファイルが空です。")
                return

            numerical_df = self.df.select_dtypes(include=['number'])
            if numerical_df.empty or len(numerical_df.columns) < 2:
                messagebox.showerror("データエラー", "分析には少なくとも2つ以上の数値列が必要です。")
                return

            self.numerical_df_columns = numerical_df.columns
            scaler = StandardScaler()
            scaled_data = scaler.fit_transform(numerical_df)
            
            self.pca = PCA(n_components=min(10, len(self.numerical_df_columns)))
            principal_components = self.pca.fit_transform(scaled_data)
            
            pc_cols = [f'PC{i+1}' for i in range(self.pca.n_components_)]
            self.pc_df = pd.DataFrame(data=principal_components, columns=pc_cols, index=numerical_df.index)
            if 'category' in self.df.columns:
                self.pc_df['category'] = self.df.loc[numerical_df.index, 'category'].astype(str)

            for i in reversed(range(self.notebook.index('end'))):
                self.notebook.forget(i)

            current_styles = self._get_current_styles()
            self._draw_plot_on_tab("PCA Scatter Plot", self._create_scatter_plot, style_map=current_styles)
            self._draw_plot_on_tab("Explained Variance", self._create_explained_variance_plot)
            self._draw_plot_on_tab("Loadings Plot", self._create_loadings_plot)
            
            messagebox.showinfo("成功", "分析と描画が完了しました。")
            self.export_button.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("エラー", f"分析中に予期せぬエラーが発生しました:\n{e}")
            import traceback
            traceback.print_exc()

    def export_results(self):
        if self.pc_df is None or self.df is None:
            messagebox.showwarning("エクスポート不可", "先に分析を実行してください。")
            return

        try:
            id_col_name = self.df.columns[0]
            columns_to_join = [self.df[[id_col_name]]]
            if 'category' in self.df.columns:
                columns_to_join.append(self.df[['category']])

            export_df = self.pc_df.join(columns_to_join)
            
            final_columns = [id_col_name]
            if 'category' in self.df.columns:
                final_columns.append('category')
            final_columns.extend([col for col in self.pc_df.columns if col.startswith('PC')])
            export_df = export_df.reindex(columns=final_columns)

            original_filename = self.file_path.split('/')[-1].rsplit('.', 1)[0]
            default_savename = f"{original_filename}_pca_scores.csv"

            filepath = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=default_savename,
                title="主成分スコアをCSV形式で保存"
            )

            if filepath:
                export_df.to_csv(filepath, index=False)
                messagebox.showinfo("成功", f"分析結果を以下のファイルに保存しました:\n{filepath}")

        except Exception as e:
            messagebox.showerror("エクスポートエラー", f"ファイルのエクスポート中にエラーが発生しました:\n{e}")

    def _draw_plot_on_tab(self, tab_title, plot_function, **kwargs):
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=tab_title)
        fig = plot_function(**kwargs)
        if not fig: return
        canvas = FigureCanvasTkAgg(fig, master=tab_frame)
        canvas.draw()
        toolbar = NavigationToolbar2Tk(canvas, tab_frame)
        toolbar.config(background=self.FRAME_COLOR)
        toolbar._message_label.config(background=self.FRAME_COLOR, foreground=self.TEXT_COLOR)
        for button in toolbar.winfo_children():
            button.config(background=self.BUTTON_COLOR)
        toolbar.update()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    def _create_scatter_plot(self, style_map=None):
        plt.style.use('default')
        matplotlib.rcParams.update({'font.family': 'serif', 'font.size': 12, 'axes.linewidth': 1.5})
        fig, ax = plt.subplots(figsize=(10, 7), facecolor='white')
        ax.set_facecolor('white')

        has_category = 'category' in self.pc_df.columns
        scatter_artists = []
        
        if has_category and style_map:
            default_style = {'color': 'black', 'marker': 'x'}
            for category_name, group_df in self.pc_df.groupby('category'):
                style = style_map.get(str(category_name), default_style)
                # ★★★ 変更点: 黒背景で見えなくなるのを防ぐため、黒(black)の点には白い縁取りを追加 ★★★
                edge_color = 'white' if style['color'] == 'black' else 'k'
                scatter = ax.scatter(group_df['PC1'], group_df['PC2'], c=style['color'], marker=style['marker'], label=category_name,
                                     alpha=0.7, s=50, edgecolors=edge_color, linewidths=0.5)
                scatter.set_gid(category_name)
                scatter_artists.append(scatter)
            legend = ax.legend(title='Category', frameon=True, facecolor='white', edgecolor='black', labelcolor='black')
            if legend.get_title(): legend.get_title().set_color('black')
        else:
            scatter = ax.scatter(self.pc_df['PC1'], self.pc_df['PC2'], alpha=0.7, c='blue', edgecolors='k', linewidths=0.5)
            scatter_artists.append(scatter)

        ax.set_title('Principal Component Analysis (2D Scatter Plot)', color='black', fontweight='bold')
        ax.set_xlabel(f'Principal Component 1 ({self.pca.explained_variance_ratio_[0]:.2%})', color='black')
        ax.set_ylabel(f'Principal Component 2 ({self.pca.explained_variance_ratio_[1]:.2%})', color='black')
        
        for spine in ax.spines.values(): spine.set_color('black')
        ax.tick_params(axis='x', colors='black'); ax.tick_params(axis='y', colors='black')
        ax.grid(True, color='#CCCCCC', linestyle='--', linewidth=0.5)
        ax.axhline(0, color='grey', lw=0.8, ls='--'); ax.axvline(0, color='grey', lw=0.8, ls='--')

        cursor = mplcursors.cursor(scatter_artists, hover=True)
        id_column_name = self.df.columns[0]
        @cursor.connect("add")
        def on_add(sel):
            point_index = sel.index
            if has_category:
                category_gid = sel.artist.get_gid()
                original_index = self.pc_df[self.pc_df['category'] == category_gid].index[point_index]
            else:
                original_index = self.pc_df.index[point_index]
            
            data_id = self.df.loc[original_index, id_column_name]
            text_lines = [f"{id_column_name}: {data_id}", f"Index: {original_index}"]
            if has_category: text_lines.append(f"Category: {self.pc_df.loc[original_index, 'category']}")
            
            sel.annotation.set_text('\n'.join(text_lines))
            sel.annotation.get_bbox_patch().set(facecolor='#4C566A', alpha=0.95, edgecolor="#E5E9F0")
            sel.annotation.arrow_patch.set(arrowstyle="->", facecolor=self.TEXT_COLOR, alpha=0.7)
            sel.annotation.set_color(self.TEXT_COLOR)
        
        fig.tight_layout(pad=2.0)
        return fig

    def _create_explained_variance_plot(self):
        plt.style.use('default')
        fig, ax1 = plt.subplots(figsize=(10, 7), facecolor='white')
        ax1.set_facecolor('white')
        variance_ratio = self.pca.explained_variance_ratio_
        cum_variance_ratio = np.cumsum(variance_ratio)
        pc_labels = [f'PC{i+1}' for i in range(len(variance_ratio))]
        ax1.bar(pc_labels, variance_ratio, alpha=0.7, color='steelblue', label='Explained Variance Ratio')
        ax1.set_xlabel('Principal Components', color='black', fontweight='bold')
        ax1.set_ylabel('Explained Variance Ratio', color='black')
        ax1.tick_params(axis='y', labelcolor='black'); ax1.tick_params(axis='x', colors='black', rotation=45)
        ax2 = ax1.twinx()
        ax2.plot(pc_labels, cum_variance_ratio, color='firebrick', marker='o', linestyle='-', label='Cumulative Explained Variance')
        ax2.set_ylabel('Cumulative Explained Variance', color='black')
        ax2.tick_params(axis='y', labelcolor='black'); ax2.set_ylim(0, 1.05)
        ax1.set_title('Explained Variance by Principal Components', color='black', fontweight='bold')
        fig.legend(loc="upper right", bbox_to_anchor=(0.9, 0.9), bbox_transform=ax1.transAxes)
        for spine in ax1.spines.values(): spine.set_color('black')
        for spine in ax2.spines.values(): spine.set_color('black')
        fig.tight_layout()
        return fig
        
    def _create_loadings_plot(self):
        plt.style.use('default')
        fig, ax = plt.subplots(figsize=(8, 8), facecolor='white')
        ax.set_facecolor('white')
        loadings = self.pca.components_.T[:, :2]
        for i, var_name in enumerate(self.numerical_df_columns):
            ax.arrow(0, 0, loadings[i, 0], loadings[i, 1], head_width=0.03, head_length=0.03, fc='darkred', ec='darkred', alpha=0.7)
            ax.text(loadings[i, 0] * 1.15, loadings[i, 1] * 1.15, var_name, color='black', ha='center', va='center',
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="none", alpha=0.7))
        ax.set_xlabel('Principal Component 1 Loadings', color='black'); ax.set_ylabel('Principal Component 2 Loadings', color='black')
        ax.set_title('Loadings Plot', color='black', fontweight='bold')
        ax.set_xlim(-1.1, 1.1); ax.set_ylim(-1.1, 1.1)
        ax.axhline(0, color='grey', ls='--', lw=0.8); ax.axvline(0, color='grey', ls='--', lw=0.8)
        ax.grid(True, color='#CCCCCC', ls='--', lw=0.5)
        for spine in ax.spines.values(): spine.set_color('black')
        ax.tick_params(axis='x', colors='black'); ax.tick_params(axis='y', colors='black')
        circle = plt.Circle((0, 0), 1, color='gray', fill=False, ls='--', alpha=0.7)
        ax.add_artist(circle)
        fig.tight_layout()
        return fig

    # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
    # ★★★ 変更点: OptionMenuの項目色を変更するロジックを追加 ★★★
    # ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
    def _setup_style_ui(self):
        for widget in self.style_inner_frame.winfo_children(): widget.destroy()
        self.style_widgets = {}
        if not self.file_path: return
        try:
            df = pd.read_csv(self.file_path)
            if 'category' not in df.columns:
                tk.Label(self.style_inner_frame, text="'category' 列が見つかりません。", bg=self.FRAME_COLOR, fg=self.TEXT_COLOR).pack(pady=20)
                return
            categories = df['category'].unique()
            header_frame = tk.Frame(self.style_inner_frame, bg=self.FRAME_COLOR)
            header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
            self.style_inner_frame.grid_columnconfigure(0, weight=1)
            tk.Label(header_frame, text="カテゴリ名", font=("Arial", 9, "bold"), bg=self.FRAME_COLOR, fg=self.TEXT_COLOR).pack(side=tk.LEFT, padx=5)
            tk.Label(header_frame, text="マーカー", font=("Arial", 9, "bold"), bg=self.FRAME_COLOR, fg=self.TEXT_COLOR).pack(side=tk.RIGHT, padx=10)
            tk.Label(header_frame, text="色", font=("Arial", 9, "bold"), bg=self.FRAME_COLOR, fg=self.TEXT_COLOR).pack(side=tk.RIGHT, padx=10)
            
            for i, category in enumerate(categories):
                cat_str = str(category)
                row_frame = tk.Frame(self.style_inner_frame, bg=self.FRAME_COLOR)
                row_frame.grid(row=i + 1, column=0, sticky="ew", pady=2)
                
                category_label = tk.Label(row_frame, text=cat_str, bg=self.FRAME_COLOR, fg=self.TEXT_COLOR, width=15, anchor='w', wraplength=100)
                category_label.pack(side=tk.LEFT, padx=5)
                
                # --- マーカー設定 ---
                marker_var = tk.StringVar(value=self.matplotlib_markers[i % len(self.matplotlib_markers)])
                marker_menu = tk.OptionMenu(row_frame, marker_var, *self.matplotlib_markers)
                marker_menu.config(width=4, bg=self.BUTTON_COLOR, fg=self.TEXT_COLOR, activebackground=self.BUTTON_ACTIVE_COLOR, relief=tk.RAISED, direction="below")
                marker_menu["menu"].config(bg=self.BUTTON_COLOR, fg=self.TEXT_COLOR)
                marker_menu.pack(side=tk.RIGHT, padx=(0, 5))
                
                # --- 色設定 ---
                color_var = tk.StringVar(value=self.matplotlib_colors[i % len(self.matplotlib_colors)])
                # ★★★ 変更点 1: OptionMenuを一旦変数に格納する ★★★
                color_menu = tk.OptionMenu(row_frame, color_var, *self.matplotlib_colors)
                color_menu.config(width=6, bg=self.BUTTON_COLOR, fg=self.TEXT_COLOR, activebackground=self.BUTTON_ACTIVE_COLOR, relief=tk.RAISED, direction="below")
                
                # ★★★ 変更点 2: ドロップダウンメニューの各項目の色を設定する ★★★
                color_menu_dropdown = color_menu["menu"]
                color_menu_dropdown.config(bg=self.BUTTON_COLOR, fg=self.TEXT_COLOR)
                # メニューの各項目に対してループ処理
                for color_name in self.matplotlib_colors:
                    # 'command'を使って、選択されたときに色を更新するラムダ関数を各項目に設定
                    # 'foreground'で各項目の文字色を設定
                    color_menu_dropdown.entryconfigure(
                        color_name, 
                        command=tk._setit(color_var, color_name), # OptionMenuの標準的なコマンド
                        foreground=color_name
                    )

                color_menu.pack(side=tk.RIGHT, padx=5)

                initial_color = color_var.get()
                category_label.config(fg=initial_color)

                color_var.trace_add("write", 
                                  lambda name, index, mode, var=color_var, label=category_label: 
                                      label.config(fg=var.get()))
                
                self.style_widgets[cat_str] = {'color_var': color_var, 'marker_var': marker_var}

        except Exception as e:
            messagebox.showerror("ファイル読み込みエラー", f"ファイルの読み込み中にエラーが発生しました:\n{e}")
            self.file_path = None
            self.path_label.config(text="ファイルが選択されていません")


    def _get_current_styles(self):
        style_map = {}
        for category, widgets in self.style_widgets.items():
            style_map[category] = {'color': widgets['color_var'].get(), 'marker': widgets['marker_var'].get()}
        return style_map

if __name__ == "__main__":
    root = tk.Tk()
    app = PcaApp(root)
    root.mainloop()
