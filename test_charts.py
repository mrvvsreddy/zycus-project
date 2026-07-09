from pptx_tools import chart_hbar, chart_grouped_vbar
import os

os.makedirs("debug_charts", exist_ok=True)

names = ["Titan", "UniSan"]
pct_vals = [41.7, 0.0]
bar_colors = ["#FF4D6A", "#00E696"]

hbar_buf = chart_hbar(names, pct_vals, bar_colors, xlabel="% Overdue Tasks")
with open("debug_charts/hbar.png", "wb") as f:
    f.write(hbar_buf.read())

blocker_vals = [5, 0]
crit_vals = [0, 0]
vbar_buf = chart_grouped_vbar(
    names,
    {"Blockers": blocker_vals, "Critical Blocked": crit_vals},
    {"Blockers": "#00D2FF", "Critical Blocked": "#FF4D6A"}
)
with open("debug_charts/vbar.png", "wb") as f:
    f.write(vbar_buf.read())
