import tkinter as tk
from tkinter import filedialog, messagebox
import warnings

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
from matplotlib.lines import Line2D

warnings.filterwarnings("ignore")


def block_to_time(block):
    mins = (block - 1) * 15
    return f"{mins // 60:02d}:{mins % 60:02d}"


def load_zone(file, sheet, col_demand, skiprows):
    raw = pd.read_excel(file, sheet_name=sheet, header=None)
    df = raw.iloc[skiprows:, [1, col_demand]].copy()
    df.columns = ["time", "demand_mw"]
    df["demand_mw"] = pd.to_numeric(df["demand_mw"], errors="coerce")
    df = df.dropna(subset=["demand_mw"]).head(96).reset_index(drop=True)
    df["block"] = range(1, len(df) + 1)
    return df


class DemandApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MP Power Management - Zone Demand Analysis Tool")
        self.root.geometry("820x640")
        self.root.resizable(True, True)
        self.root.configure(bg="#F5F5F5")

        self.cz_file = tk.StringVar()
        self.ez_file = tk.StringVar()
        self.wz_file = tk.StringVar()

        self.cz = self.ez = self.wz = self.fig = None

        self.build_ui()

    def build_ui(self):
        title_frame = tk.Frame(self.root, bg="#0F6E56", pady=12)
        title_frame.pack(fill="x")

        tk.Label(
            title_frame,
            text="MP Power Management",
            font=("Segoe UI", 16, "bold"),
            bg="#0F6E56",
            fg="white",
        ).pack()
        tk.Label(
            title_frame,
            text="Zone-wise Day Ahead Demand Analysis Tool",
            font=("Segoe UI", 10),
            bg="#0F6E56",
            fg="#9FE1CB",
        ).pack()

        main = tk.Frame(self.root, bg="#F5F5F5", padx=20, pady=15)
        main.pack(fill="both", expand=True)

        file_frame = tk.LabelFrame(
            main,
            text="  Upload Demand Files  ",
            font=("Segoe UI", 9, "bold"),
            bg="#F5F5F5",
            fg="#333",
            padx=12,
            pady=10,
        )
        file_frame.pack(fill="x", pady=(0, 12))

        zones = [
            ("Central Zone (CZ)", self.cz_file, "#1D9E75"),
            ("East Zone (EZ)", self.ez_file, "#534AB7"),
            ("West Zone (WZ)", self.wz_file, "#D85A30"),
        ]

        for label, var, color in zones:
            row = tk.Frame(file_frame, bg="#F5F5F5", pady=4)
            row.pack(fill="x")

            tk.Label(
                row,
                text=label,
                width=20,
                anchor="w",
                font=("Segoe UI", 10, "bold"),
                fg=color,
                bg="#F5F5F5",
            ).pack(side="left")

            tk.Entry(
                row,
                textvariable=var,
                font=("Segoe UI", 9),
                width=46,
                relief="solid",
                bd=1,
                state="readonly",
            ).pack(side="left", padx=6)

            tk.Button(
                row,
                text="Browse",
                font=("Segoe UI", 9),
                bg=color,
                fg="white",
                relief="flat",
                padx=10,
                cursor="hand2",
                command=lambda v=var: self.browse_file(v),
            ).pack(side="left")

        btn_frame = tk.Frame(main, bg="#F5F5F5", pady=5)
        btn_frame.pack()

        tk.Button(
            btn_frame,
            text="  Analyse & Show Chart  ",
            font=("Segoe UI", 11, "bold"),
            bg="#0F6E56",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.run_analysis,
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame,
            text="  Save Chart as PNG  ",
            font=("Segoe UI", 11, "bold"),
            bg="#534AB7",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.save_chart,
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame,
            text="  Clear  ",
            font=("Segoe UI", 11),
            bg="#888",
            fg="white",
            relief="flat",
            padx=14,
            pady=8,
            cursor="hand2",
            command=self.clear_all,
        ).pack(side="left", padx=8)

        summary_frame = tk.LabelFrame(
            main,
            text="  Demand Summary  ",
            font=("Segoe UI", 9, "bold"),
            bg="#F5F5F5",
            fg="#333",
            padx=10,
            pady=8,
        )
        summary_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.summary_text = tk.Text(
            summary_frame,
            font=("Courier New", 9),
            height=10,
            relief="flat",
            bg="#1E1E1E",
            fg="#D4D4D4",
            insertbackground="white",
            padx=10,
            pady=8,
        )
        self.summary_text.pack(fill="both", expand=True)

        self.status = tk.StringVar(
            value="Ready - Upload your 3 zone files and click Analyse"
        )
        tk.Label(
            self.root,
            textvariable=self.status,
            font=("Segoe UI", 9),
            bg="#0F6E56",
            fg="white",
            anchor="w",
            padx=12,
            pady=4,
        ).pack(fill="x", side="bottom")

    def browse_file(self, var):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if path:
            var.set(path)

    def validate(self):
        for name, var in [
            ("Central Zone", self.cz_file),
            ("East Zone", self.ez_file),
            ("West Zone", self.wz_file),
        ]:
            if not var.get():
                messagebox.showwarning("Missing File", f"Please select the {name} file.")
                return False
        return True

    def load_data(self):
        try:
            self.cz = load_zone(self.cz_file.get(), "Day Ahead", 3, 4)
            self.ez = load_zone(self.ez_file.get(), "Sheet1 (2)", 5, 4)
            self.wz = load_zone(self.wz_file.get(), "WZ", 7, 1)
            return True
        except Exception as e:
            messagebox.showerror("Error loading file", str(e))
            return False

    def show_summary(self):
        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", "end")

        lines = []
        lines.append("  " + "=" * 54)

        for df, name in [
            (self.cz, "Central Zone"),
            (self.ez, "East Zone"),
            (self.wz, "West Zone"),
        ]:
            pk_blk = int(df["demand_mw"].idxmax()) + 1
            min_blk = int(df["demand_mw"].idxmin()) + 1
            avg_val = df["demand_mw"].mean()
            avg_blk = int((df["demand_mw"] - avg_val).abs().idxmin()) + 1

            lines.append(f"\n  {name}")
            lines.append(f"  {'-' * 44}")
            lines.append(
                f"  Peak : {df['demand_mw'].max():>7,.0f} MW"
                f"  |  Block {pk_blk:>2}  ({block_to_time(pk_blk)})"
            )
            lines.append(
                f"  Min  : {df['demand_mw'].min():>7,.0f} MW"
                f"  |  Block {min_blk:>2}  ({block_to_time(min_blk)})"
            )
            lines.append(
                f"  Avg  : {avg_val:>7,.0f} MW"
                f"  |  Block {avg_blk:>2}  ({block_to_time(avg_blk)})"
            )

        total = self.cz["demand_mw"] + self.ez["demand_mw"] + self.wz["demand_mw"]
        lines.append(f"\n  {'-' * 44}")
        lines.append("  All 3 Zones Combined")
        lines.append(
            f"  Peak : {total.max():>7,.0f} MW"
            f"  |  Block {int(total.idxmax()) + 1}"
            f"  ({block_to_time(int(total.idxmax()) + 1)})"
        )
        lines.append(f"  Avg  : {total.mean():>7,.0f} MW")
        lines.append("\n  " + "=" * 54)

        self.summary_text.insert("end", "\n".join(lines))
        self.summary_text.config(state="disabled")

    def build_chart(self):
        colors = {"CZ": "#1D9E75", "EZ": "#534AB7", "WZ": "#D85A30"}
        light_colors = {"CZ": "#A7D8C9", "EZ": "#B4B0E3", "WZ": "#EDB39F"}
        zone_frames = [
            (self.cz, "CZ", "Central Zone"),
            (self.ez, "EZ", "East Zone"),
            (self.wz, "WZ", "West Zone"),
        ]
        x_ticks = list(range(4, 97, 4))
        x_tick_labels = [block_to_time(i) for i in x_ticks]

        self.fig, axes = plt.subplots(
            2,
            1,
            figsize=(15.5, 9.8),
            gridspec_kw={"height_ratios": [1.45, 1.0]},
            constrained_layout=True,
        )
        self.fig.patch.set_facecolor("#F6F7FB")
        self.fig.suptitle(
            "MP Power Management - Zone Demand Analysis",
            fontsize=19,
            fontweight="bold",
            y=0.99,
        )

        ax1 = axes[0]
        ax1.set_facecolor("white")

        y_min = min(df["demand_mw"].min() for df, _, _ in zone_frames)
        y_max = max(df["demand_mw"].max() for df, _, _ in zone_frames)
        y_range = y_max - y_min
        peak_offsets = {
            "CZ": (10, -42),
            "EZ": (-68, -20),
            "WZ": (10, 16),
        }

        for df, key, label in zone_frames:
            ax1.plot(
                df["block"],
                df["demand_mw"],
                color=colors[key],
                linewidth=2.8,
                label=label,
                solid_capstyle="round",
                zorder=3,
            )
            ax1.fill_between(
                df["block"],
                df["demand_mw"],
                y_min - y_range * 0.08,
                color=colors[key],
                alpha=0.05,
                zorder=1,
            )

            pk = int(df["demand_mw"].idxmax())
            pk_val = df["demand_mw"].iloc[pk]
            ax1.scatter(
                pk + 1,
                pk_val,
                color=colors[key],
                s=135,
                edgecolor="white",
                linewidth=1.6,
                zorder=5,
            )

            dx, dy = peak_offsets[key]
            ax1.annotate(
                f"{label}\nBlock {pk + 1}  |  {block_to_time(pk + 1)}\n{pk_val:,.0f} MW",
                xy=(pk + 1, pk_val),
                xytext=(dx, dy),
                textcoords="offset points",
                fontsize=9.2,
                color=colors[key],
                fontweight="bold",
                ha="left" if dx > 0 else "right",
                va="bottom" if dy > 0 else "top",
                arrowprops=dict(
                    arrowstyle="-|>",
                    color=colors[key],
                    lw=1.4,
                    shrinkA=6,
                    shrinkB=6,
                    mutation_scale=12,
                ),
                bbox=dict(
                    boxstyle="round,pad=0.45",
                    fc="white",
                    ec=colors[key],
                    alpha=0.96,
                    linewidth=1.4,
                ),
            )

        ax1.set_xticks(x_ticks)
        ax1.set_xticklabels(x_tick_labels, rotation=40, ha="right", fontsize=9)
        ax1.set_ylabel("Demand (MW)", fontsize=11, fontweight="bold")
        ax1.set_title(
            "Zone-wise Demand Curve - 96 Blocks (15 min each)",
            fontsize=13,
            pad=10,
            fontweight="bold",
        )
        ax1.legend(
            loc="upper right",
            fontsize=10.5,
            frameon=True,
            facecolor="white",
            edgecolor="#D8DCE6",
        )
        ax1.yaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, _: f"{int(x):,}")
        )
        ax1.grid(axis="y", alpha=0.24, linestyle="--", linewidth=0.9)
        ax1.grid(axis="x", alpha=0.08, linestyle="--", linewidth=0.8)
        ax1.set_xlim(1, 96)
        ax1.set_ylim(y_min - y_range * 0.10, y_max + y_range * 0.24)
        ax1.tick_params(axis="y", labelsize=10)
        ax1.tick_params(axis="x", colors="#444")
        ax1.spines["top"].set_visible(False)
        ax1.spines["right"].set_visible(False)
        ax1.spines["left"].set_color("#C8CDD8")
        ax1.spines["bottom"].set_color("#C8CDD8")

        ymin_plot = ax1.get_ylim()[0]
        ax1.axvspan(25, 40, alpha=0.08, color="#D8BE8A", zorder=0)
        ax1.axvspan(64, 80, alpha=0.08, color="#E8C0AE", zorder=0)
        ax1.text(
            26,
            ymin_plot + y_range * 0.05,
            "Morning peak window",
            fontsize=8.5,
            color="#9A6820",
            fontweight="bold",
        )
        ax1.text(
            65,
            ymin_plot + y_range * 0.05,
            "Evening peak window",
            fontsize=8.5,
            color="#C0633D",
            fontweight="bold",
        )

        ax2 = axes[1]
        ax2.set_facecolor("white")

        zones = ["Central Zone", "East Zone", "West Zone"]
        avg_vals = [
            self.cz["demand_mw"].mean(),
            self.ez["demand_mw"].mean(),
            self.wz["demand_mw"].mean(),
        ]
        peak_vals = [
            self.cz["demand_mw"].max(),
            self.ez["demand_mw"].max(),
            self.wz["demand_mw"].max(),
        ]
        x = list(range(len(zones)))
        bar_width = 0.34

        for i, (avg, peak, key) in enumerate(zip(avg_vals, peak_vals, ["CZ", "EZ", "WZ"])):
            ax2.bar(
                i - bar_width / 1.7,
                avg,
                width=bar_width,
                color=light_colors[key],
                edgecolor="white",
                linewidth=1.0,
                zorder=3,
            )
            ax2.bar(
                i + bar_width / 1.7,
                peak,
                width=bar_width,
                color=colors[key],
                edgecolor="white",
                linewidth=1.0,
                zorder=3,
            )

            ax2.text(
                i - bar_width / 1.7,
                avg + 55,
                f"{avg:,.0f}",
                ha="center",
                va="bottom",
                fontsize=9.5,
                color=colors[key],
                fontweight="bold",
            )
            ax2.text(
                i + bar_width / 1.7,
                peak + 55,
                f"{peak:,.0f}",
                ha="center",
                va="bottom",
                fontsize=9.5,
                color=colors[key],
                fontweight="bold",
            )

        ax2.set_xticks(x)
        ax2.set_xticklabels(zones, fontsize=11, fontweight="bold")
        ax2.set_ylabel("Demand (MW)", fontsize=11, fontweight="bold")
        ax2.set_title(
            "Average vs Peak Demand by Zone",
            fontsize=13,
            pad=10,
            fontweight="bold",
        )
        ax2.legend(
            handles=[
                Line2D([0], [0], color="#7D8AA6", lw=10, alpha=0.35, label="Average Demand"),
                Line2D([0], [0], color="#7D8AA6", lw=10, alpha=1.0, label="Peak Demand"),
            ],
            fontsize=10,
            frameon=True,
            facecolor="white",
            edgecolor="#D8DCE6",
            loc="upper left",
        )
        ax2.yaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, _: f"{int(x):,}")
        )
        ax2.grid(axis="y", alpha=0.24, linestyle="--", linewidth=0.9)
        ax2.tick_params(axis="y", labelsize=10)
        ax2.tick_params(axis="x", colors="#444")
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.spines["left"].set_color("#C8CDD8")
        ax2.spines["bottom"].set_color("#C8CDD8")
        ax2.set_ylim(0, max(peak_vals) * 1.22)

    def run_analysis(self):
        if not self.validate():
            return
        self.status.set("Loading data...")
        self.root.update()

        if not self.load_data():
            self.status.set("Error loading files.")
            return

        self.show_summary()
        self.build_chart()
        self.fig.show()
        self.status.set("Analysis complete - Chart displayed. Use Save PNG to export.")

    def save_chart(self):
        if self.fig is None:
            messagebox.showwarning("No Chart", "Please run analysis first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile="mp_demand_analysis.png",
            filetypes=[("PNG Image", "*.png")],
        )
        if path:
            self.fig.savefig(
                path,
                dpi=220,
                bbox_inches="tight",
                facecolor="#F6F7FB",
            )
            messagebox.showinfo("Saved!", f"Chart saved to:\n{path}")
            self.status.set(f"Chart saved: {path}")

    def clear_all(self):
        self.cz_file.set("")
        self.ez_file.set("")
        self.wz_file.set("")
        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.config(state="disabled")
        self.cz = self.ez = self.wz = self.fig = None
        self.status.set("Cleared - Ready for new analysis")


if __name__ == "__main__":
    root = tk.Tk()
    app = DemandApp(root)
    root.mainloop()
