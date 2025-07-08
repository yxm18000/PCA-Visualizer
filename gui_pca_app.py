# (前の回答と同じコードをここにコピー＆ペースト)
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import mplcursors

# --- GUIアプリケーションのクラス定義 ---
class PcaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PCA Visualizer")
        self.root.geometry("900x700")

        self.file_path = None
        self.canvas_widget = None
        self.toolbar = None

        # --- UI要素の作成 ---
        control_frame = tk.Frame(root, padx=10, pady=10)
        control_frame.pack(fill=tk.X)

        select_button = tk.Button(control_frame, text="1. CSVファイルを選択", command=self.select_file)
        select_button.pack(side=tk.LEFT, padx=5)

        self.path_label = tk.Label(control_frame, text="ファイルが選択されていません", relief=tk.SUNKEN, padx=5, anchor='w')
        self.path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        run_button = tk.Button(control_frame, text="2. 分析実行", command=self.run_analysis)
        run_button.pack(side=tk.LEFT, padx=5)

        self.graph_frame = tk.Frame(root)
        self.graph_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

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
            fig = self._create_pca_figure()
            if not fig:
                return

            self.canvas_widget = FigureCanvasTkAgg(fig, master=self.graph_frame)
            self.canvas_widget.draw()
            
            self.toolbar = NavigationToolbar2Tk(self.canvas_widget, self.graph_frame)
            self.toolbar.update()
            
            self.canvas_widget.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            messagebox.showinfo("成功", "分析と描画が完了しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"分析中に予期せぬエラーが発生しました:\n{e}")
            print(f"エラー: {e}")

    def _create_pca_figure(self):
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

        fig, ax = plt.subplots(figsize=(10, 7))
        scatter_artists = []

        if has_category:
            style_map = {
                'on':      {'color': 'blue',   'marker': 'o', 'label': 'On'},
                'off':     {'color': 'red',    'marker': 'x', 'label': 'Off'},
                'unknown': {'color': 'gray',   'marker': 's', 'label': 'Unknown'}
            }
            default_style = {'color': 'purple', 'marker': '^', 'label': 'Other'}

            for category_name, group_df in pc_df.groupby('category'):
                style = style_map.get(str(category_name).lower(), default_style)
                scatter = ax.scatter(group_df['PC1'], group_df['PC2'], c=style['color'], 
                                     marker=style['marker'], label=style['label'], alpha=0.7)
                scatter.set_gid(category_name)
                scatter_artists.append(scatter)
            ax.legend(title='Category')
        else:
            scatter = ax.scatter(pc_df['PC1'], pc_df['PC2'], alpha=0.7)
            scatter_artists.append(scatter)

        ax.set_title('Principal Component Analysis (2D Visualization)')
        ax.set_xlabel('Principal Component 1')
        ax.set_ylabel('Principal Component 2')
        ax.grid(True)
        ax.axhline(0, color='grey', linewidth=0.8)
        ax.axvline(0, color='grey', linewidth=0.8)

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
            
            sel.annotation.set_bbox(dict(boxstyle="round,pad=0.3", fc="lightblue", ec="black", lw=0.5, alpha=0.9))
            sel.annotation.arrow_patch.set(arrowstyle="->", fc="black", alpha=0.5)

        return fig

if __name__ == "__main__":
    root = tk.Tk()
    app = PcaApp(root)
    root.mainloop()