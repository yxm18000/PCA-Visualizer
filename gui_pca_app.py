import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
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
        self.root.geometry("900x700")
        self.root.configure(bg=self.BG_COLOR)

        self.file_path = None
        self.canvas_widget = None
        self.toolbar = None

        # --- Windowsのタイトルバーをダークモードに ---
        self.set_dark_title_bar()

        # --- UI要素の作成 ---
        control_frame = tk.Frame(root, padx=10, pady=10, bg=self.BG_COLOR)
        control_frame.pack(fill=tk.X)

        button_style = {
            "font": ("Arial", 10), "bg": self.BUTTON_COLOR, "fg": self.TEXT_COLOR,
            "activebackground": self.BUTTON_ACTIVE_COLOR, "activeforeground": self.TEXT_COLOR,
            "relief": tk.RAISED, "bd": 2, "padx": 10
        }

        select_button = tk.Button(control_frame, text="1. CSVファイルを選択", command=self.select_file, **button_style)
        select_button.pack(side=tk.LEFT, padx=5)

        self.path_label = tk.Label(
            control_frame, text="ファイルが選択されていません", relief=tk.SUNKEN, bd=2,
            padx=5, anchor='w', bg=self.LABEL_BG_COLOR, fg=self.TEXT_COLOR, font=("Arial", 9)
        )
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        run_button = tk.Button(control_frame, text="2. 分析実行", command=self.run_analysis, **button_style)
        run_button.pack(side=tk.LEFT, padx=5)

        self.graph_frame = tk.Frame(root, bg=self.FRAME_COLOR, bd=1, relief=tk.SUNKEN)
        self.graph_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def set_dark_title_bar(self):
        """Windows 10/11のタイトルバーをダークモードにする"""
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
        path = filedialog.askopenfilename(
            title="CSVファイルを選択してください",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.file_path = path.strip(' "')
            display_name = self.file_path.split('/')[-1]
            self.path_label.config(text=display_name)
            print(f"ファイルを選択しました: {self.file_path}")

    def run_analysis(self):
        if not self.file_path:
            messagebox.showwarning("警告", "先にCSVファイルを選択してください。")
            return

        if self.canvas_widget:
            self.canvas_widget.get_tk_widget().destroy()
        if self.toolbar:
            self.toolbar.destroy()

        try:
            fig = self._create_pca_figure_for_publication()
            if not fig:
                return

            self.canvas_widget = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.canvas_widget.draw()
            
            self.toolbar = NavigationToolbar2Tk(self.canvas_widget, self.graph_frame)
            self.toolbar.config(background=self.FRAME_COLOR)
            self.toolbar._message_label.config(background=self.FRAME_COLOR, foreground=self.TEXT_COLOR)
            for button in self.toolbar.winfo_children():
                button.config(background=self.BUTTON_COLOR)
            self.toolbar.update()
            
            self.canvas_widget.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            messagebox.showinfo("成功", "分析と描画が完了しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"分析中に予期せぬエラーが発生しました:\n{e}")
            print(f"エラー: {e}")

    def _create_pca_figure_for_publication(self):
        df = pd.read_csv(self.file_path)
        numerical_df = df.select_dtypes(include=['number'])
        has_category = 'category' in df.columns

        if numerical_df.empty or len(numerical_df) < 2:
            messagebox.showerror("データエラー", "分析可能な数値データが不足しています。")
            return None

        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(numerical_df)
        pca = PCA(n_components=2)
        principal_components = pca.fit_transform(scaled_data)

        pc_df = pd.DataFrame(data=principal_components, columns=['PC1', 'PC2'], index=numerical_df.index)
        if has_category:
            pc_df['category'] = df.loc[numerical_df.index, 'category']

        plt.style.use('default')
        matplotlib.rcParams.update({
            'font.family': 'serif', 'font.size': 12, 'axes.linewidth': 1.5,
        })

        fig, ax = plt.subplots(figsize=(10, 7), facecolor='white')
        ax.set_facecolor('white')

        scatter_artists = []
        
        if has_category:
            style_map = {
                'on':      {'color': 'blue',   'marker': 'o', 'label': 'On'},
                'off':     {'color': 'red',    'marker': 's', 'label': 'Off'},
                'unknown': {'color': 'green',  'marker': '^', 'label': 'Unknown'}
            }
            default_style = {'color': 'black', 'marker': 'x', 'label': 'Other'}

            for category_name, group_df in pc_df.groupby('category'):
                style = style_map.get(str(category_name).lower(), default_style)
                scatter = ax.scatter(group_df['PC1'], group_df['PC2'], c=style['color'], 
                                     marker=style['marker'], label=style['label'],
                                     alpha=0.7, s=50, edgecolors='k', linewidths=0.5)
                scatter.set_gid(category_name)
                scatter_artists.append(scatter)
            
            legend = ax.legend(title='Category', frameon=True, facecolor='white', edgecolor='black',
                               labelcolor='black')
            if legend.get_title() is not None:
                legend.get_title().set_color('black')
        else:
            scatter = ax.scatter(pc_df['PC1'], pc_df['PC2'], alpha=0.7, c='blue',
                                 edgecolors='k', linewidths=0.5)
            scatter_artists.append(scatter)

        ax.set_title('Principal Component Analysis (2D Visualization)', color='black', fontweight='bold')
        ax.set_xlabel('Principal Component 1', color='black')
        ax.set_ylabel('Principal Component 2', color='black')
        
        ax.spines['bottom'].set_color('black')
        ax.spines['top'].set_color('black')
        ax.spines['left'].set_color('black')
        ax.spines['right'].set_color('black')
        ax.tick_params(axis='x', colors='black')
        ax.tick_params(axis='y', colors='black')

        ax.grid(True, color='#CCCCCC', linestyle='--', linewidth=0.5)
        ax.axhline(0, color='grey', linewidth=0.8, linestyle='--')
        ax.axvline(0, color='grey', linewidth=0.8, linestyle='--')

        cursor = mplcursors.cursor(scatter_artists, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            point_index = sel.index
            if has_category:
                category_name = sel.artist.get_gid()
                original_index = pc_df[pc_df['category'] == category_name].index[point_index]
                sel.annotation.set_text(f'Index: {original_index}\nCategory: {category_name}')
            else:
                original_index = pc_df.index[point_index]
                sel.annotation.set_text(f'Index: {original_index}')
            
            sel.annotation.get_bbox_patch().set(facecolor='#4C566A', alpha=0.95, edgecolor="#E5E9F0")
            sel.annotation.arrow_patch.set(arrowstyle="->", facecolor=self.TEXT_COLOR, alpha=0.7)
            sel.annotation.set_color(self.TEXT_COLOR)

        fig.tight_layout(pad=2.0)
        return fig

if __name__ == "__main__":
    root = tk.Tk()
    app = PcaApp(root)
    root.mainloop()
