"""Figures for the ASTra launch article. Numbers come from benchmarks/results/0.4.1/REPORT.md."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle, FancyArrowPatch

plt.rcParams["font.family"] = "Liberation Sans"
INK, MUTED, LINE = "#1b2430", "#5b6675", "#d9dee5"
BLUE, BLUE_L = "#2457d6", "#dbe6ff"
GREEN, GREEN_L = "#17805d", "#d8f1e7"
AMBER, AMBER_L = "#b86a00", "#fdebcc"
RED, RED_L = "#b3261e", "#fbdedb"
BG = "#ffffff"
import os
OUT = os.path.dirname(os.path.abspath(__file__)) + "/"


def save(fig, name):
    fig.savefig(OUT + name, dpi=200, facecolor=BG, bbox_inches="tight", pad_inches=0.25)
    plt.close(fig)


def box(ax, x, y, w, h, fc, ec, text="", fs=12, color=INK, weight="bold", r=0.04, lw=1.6):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}", fc=fc, ec=ec, lw=lw))
    if text:
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, color=color, fontweight=weight)


def arrow(ax, x1, y1, x2, y2, color=MUTED, lw=2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=16, color=color, lw=lw))


# 1. Hero: the five verbs --------------------------------------------------
fig, ax = plt.subplots(figsize=(11.5, 5.2))
ax.set_xlim(0, 11.5); ax.set_ylim(0, 5.2); ax.axis("off")
ax.text(0.1, 4.9, "An AI agent works on a Word file through ASTra, not through XML", fontsize=17, fontweight="bold", color=INK)
ax.text(0.1, 4.5, "Runs locally. Nothing leaves your machine.", fontsize=12, color=MUTED)
box(ax, 0.1, 1.5, 1.85, 1.5, "#f3f5f8", LINE, "AI agent\n(Claude, Cursor,\nany MCP client)", 10.5)
box(ax, 9.65, 1.5, 1.75, 1.5, "#f3f5f8", LINE, "report.docx\nplain Word file", 10.5)
box(ax, 2.55, 0.75, 6.5, 3.35, BLUE_L, BLUE, "", lw=2)
ax.text(5.8, 3.8, "ASTra", ha="center", fontsize=14, fontweight="bold", color=BLUE)
verbs = [("Create", "Markdown in,\nreal .docx out"), ("Read", "compact outline\nwith stable IDs"), ("Edit", "typed edits in\none atomic batch"),
         ("Check", "schema gate on\nevery save"), ("Undo", "git history,\npreview, revert")]
w = 1.12
for i, (v, d) in enumerate(verbs):
    x = 2.7 + i * 1.22
    box(ax, x, 1.0, w, 2.55, "white", BLUE, "", lw=1.2)
    ax.text(x + w / 2, 3.2, v, ha="center", fontsize=12.5, fontweight="bold", color=INK)
    ax.text(x + w / 2, 2.15, d, ha="center", va="center", fontsize=8, color=MUTED, linespacing=1.5)
arrow(ax, 2.0, 2.25, 2.5, 2.25); arrow(ax, 9.1, 2.25, 9.6, 2.25)
ax.text(5.8, 0.3, "19 MCP tools  \u00b7  CLI  \u00b7  TypeScript library", ha="center", fontsize=11, color=MUTED)
save(fig, "01-overview.png")

# 2. Reading cost -----------------------------------------------------------
kinds = [("tables, wide", 6.38), ("tables", 6.25), ("mixed", 4.30), ("links + bookmarks", 2.84), ("tracked changes", 2.74),
         ("comments", 2.64), ("prose", 2.52), ("lists", 2.22)]
kinds_sorted = sorted(kinds, key=lambda k: k[1])
fig, ax = plt.subplots(figsize=(10, 5.2))
ys = range(len(kinds_sorted))
ax.barh(list(ys), [k[1] for k in kinds_sorted], color=BLUE, height=0.62, zorder=3)
ax.set_yticks(list(ys)); ax.set_yticklabels([k[0] for k in kinds_sorted], fontsize=12, color=INK)
for y, k in zip(ys, kinds_sorted):
    ax.text(k[1] - 0.12, y, f"{k[1]:.1f}x", va="center", ha="right", fontsize=12, fontweight="bold", color="white")
ax.axvline(2.89, color=AMBER, lw=2, ls="--", zorder=4)
ax.text(3.0, -0.85, "median of all 40 documents: 2.9x", color=AMBER, fontsize=11.5, fontweight="bold", va="center")
ax.set_ylim(-1.2, len(kinds_sorted) - 0.5); ax.set_xlim(0, 7.6); ax.set_xlabel("how many times fewer tokens than word/document.xml", fontsize=11, color=MUTED)
ax.grid(axis="x", color=LINE, zorder=0); ax.set_axisbelow(True)
for s in ("top", "right", "left"): ax.spines[s].set_visible(False)
ax.tick_params(left=False, colors=MUTED)
ax.set_title("What the agent reads: outline vs raw XML", loc="left", fontsize=16, fontweight="bold", color=INK, pad=14)
save(fig, "02-reading-cost.png")

# 3. Writing cost (create) ---------------------------------------------------
fig, ax = plt.subplots(figsize=(10, 3.4))
ax.set_xlim(0, 10); ax.set_ylim(0, 3.4); ax.axis("off")
ax.text(0.0, 3.05, "What the agent writes to create a one-page report", fontsize=16, fontweight="bold", color=INK)
ax.add_patch(Rectangle((0, 1.75), 4333 / 4333 * 8.4, 0.5, fc=LINE, ec="none"))
ax.add_patch(Rectangle((0, 0.85), 273 / 4333 * 8.4, 0.5, fc=BLUE, ec="none"))
ax.text(0, 2.38, "the XML it becomes: 4,333 tokens", fontsize=11.5, color=MUTED)
ax.text(0, 1.48, "Markdown the agent writes: 273 tokens", fontsize=11.5, color=BLUE, fontweight="bold")
ax.text(8.6, 1.98, "", fontsize=1)
ax.text(0, 0.3, "16x fewer on that page; 15.9x to 17.8x across the two creation documents (o200k_base tokenizer, a proxy)", fontsize=10.5, color=MUTED)
save(fig, "03-writing-cost.png")

# 4. Edit locality -----------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(11, 4.4))
for ax in axes:
    ax.set_xlim(0, 10); ax.set_ylim(1.2, 6); ax.axis("off")
def grid(ax, n, hi, cols, color, light):
    size, gap = 0.62, 0.16
    for i in range(n):
        r, c = divmod(i, cols)
        x = 0.4 + c * (size + gap); y = 4.35 - r * (size + gap)
        on = i in hi
        ax.add_patch(FancyBboxPatch((x, y), size, size, boxstyle="round,pad=0,rounding_size=0.07", fc=color if on else light, ec=color if on else LINE, lw=1.2))
ax = axes[0]
ax.text(0.4, 5.75, "ASTra: change one date", fontsize=14, fontweight="bold", color=INK)
ax.text(0.4, 5.4, "median over 1,092 benchmark edits", fontsize=10, color=MUTED)
grid(ax, 17, {0}, 9, BLUE, "#eef2f8")
ax.text(0.4, 2.6, "1 of 17 package parts changed", fontsize=12.5, fontweight="bold", color=BLUE)
ax.text(0.4, 2.1, "Demo report: 1 of 19 paragraphs changed,", fontsize=10.5, color=MUTED)
ax.text(0.4, 1.7, "the bold \"34%\" next to it untouched.", fontsize=10.5, color=MUTED)
ax = axes[1]
ax.text(0.4, 5.75, "Plain parse and re-serialise", fontsize=14, fontweight="bold", color=INK)
ax.text(0.4, 5.4, "no product, no edit, median of 40 documents", fontsize=10, color=MUTED)
grid(ax, 17, set(range(8)), 9, AMBER, "#eef2f8")
ax.text(0.4, 2.6, "8 of 17 package parts rewritten", fontsize=12.5, fontweight="bold", color=AMBER)
ax.text(0.4, 2.1, "Same content, different bytes.", fontsize=10.5, color=MUTED)
fig.suptitle("A small edit should be a small change", x=0.06, ha="left", fontsize=17, fontweight="bold", color=INK, y=1.03)
save(fig, "04-edit-locality.png")

# 5. The gate ------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 4.4))
ax.set_xlim(0, 11); ax.set_ylim(0, 4.4); ax.axis("off")
ax.text(0.1, 4.05, "The gate sits at the save: an edit that adds errors never reaches the disk", fontsize=16, fontweight="bold", color=INK)
box(ax, 0.1, 1.8, 2.0, 1.2, "#f3f5f8", LINE, "edit batch\ncolor: \"banana\"", 11)
box(ax, 2.9, 1.8, 2.0, 1.2, "#f3f5f8", LINE, "session\n(in memory)", 11)
box(ax, 5.7, 1.8, 2.1, 1.2, BLUE_L, BLUE, "schema check\nOpen XML SDK", 11, color=BLUE)
box(ax, 8.6, 2.55, 2.3, 0.95, RED_L, RED, "save refused\n7 new errors + xPaths", 10.5, color=RED)
box(ax, 8.6, 1.0, 2.3, 0.95, GREEN_L, GREEN, "saved + new\ngit commit", 10.5, color=GREEN)
arrow(ax, 2.15, 2.4, 2.85, 2.4); arrow(ax, 4.95, 2.4, 5.65, 2.4)
arrow(ax, 7.85, 2.7, 8.55, 3.0, color=RED); arrow(ax, 7.85, 2.1, 8.55, 1.5, color=GREEN)
ax.text(8.4, 3.5, "adds errors", fontsize=9.5, color=RED, ha="right")
ax.text(7.95, 1.1, "no new errors", fontsize=9.5, color=GREEN, ha="center")
ax.text(0.1, 0.35, "Default mode refuses any save that adds errors. Strict mode refuses any error at all.", fontsize=11, color=MUTED)
save(fig, "05-gate.png")

# 6. History timeline ----------------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 3.9))
ax.set_xlim(0, 11); ax.set_ylim(0, 3.9); ax.axis("off")
ax.text(0.1, 3.55, "Every save is a commit. Undo is a new commit, never a rewrite.", fontsize=16, fontweight="bold", color=INK)
ax.plot([0.9, 9.6], [2.0, 2.0], color=LINE, lw=4, zorder=1)
pts = [(1.0, "8d60e2e4", "Initial state"), (3.9, "974c6a19", "Create the Q3 board\nreport from Markdown"), (6.8, "3ce87606", "Move the board meeting\nto November 18"), (9.6, "revert", "Back to the\nOctober 14 text")]
for i, (x, h, m) in enumerate(pts):
    c = BLUE if i < 3 else GREEN
    ax.scatter([x], [2.0], s=330, color=c, zorder=3, edgecolor="white", linewidth=2)
    ax.text(x, 2.55, h, ha="center", fontsize=11, fontweight="bold", color=c, family="Noto Sans Mono")
    ax.text(x, 1.15, m, ha="center", va="top", fontsize=10.5, color=INK, linespacing=1.35)
ax.text(8.2, 1.45, "preview_revert\nshows the diff first", ha="center", fontsize=9, color=MUTED)
save(fig, "06-history.png")

# 7. Latency --------------------------------------------------------------------
pages = [1, 10, 50, 100, 200]
save_ms = [316, 332, 500, 712, 1120]; open_ms = [103, 118, 211, 313, 549]; edit_ms = [116, 129, 204, 309, 524]
fig, ax = plt.subplots(figsize=(10, 4.8))
for ys_, lab, col in ((save_ms, "save (includes schema check)", BLUE), (open_ms, "open", GREEN), (edit_ms, "one edit", AMBER)):
    ax.plot(pages, ys_, marker="o", lw=2.5, color=col, label=lab, zorder=3)
    ax.text(pages[-1] + 8, ys_[-1] + {"save (includes schema check)": 0, "open": 45, "one edit": -45}[lab], f"{ys_[-1]} ms", color=col, fontsize=11, fontweight="bold", va="center")
ax.set_xscale("log"); ax.set_xticks(pages); ax.set_xticklabels([str(p) for p in pages], fontsize=11)
ax.set_xlim(0.8, 420); ax.set_ylim(0, 1300)
ax.set_xlabel("document length (pages)", fontsize=11, color=MUTED); ax.set_ylabel("median milliseconds", fontsize=11, color=MUTED)
ax.grid(color=LINE, zorder=0); ax.set_axisbelow(True)
for s in ("top", "right"): ax.spines[s].set_visible(False)
ax.legend(frameon=False, fontsize=11, loc="upper left")
ax.set_title("A 100-page document saves in 0.7 seconds", loc="left", fontsize=16, fontweight="bold", color=INK, pad=14)
ax.text(0.8, -260, "15 timed runs per size, one laptop (13th-gen i7, WSL2), published release 0.4.1", fontsize=9.5, color=MUTED)
save(fig, "07-latency.png")

# 8. Scoreboard -------------------------------------------------------------------
rows = [
    ("Edits touch only what you asked", "goal: no other paragraph changes", "1,047 of 1,047", True),
    ("Produced files open in independent readers", "goal: all 5 readers pass (none use ASTra code)", "1,132 of 1,132", True),
    ("Creating from Markdown keeps every block", "goal: at least 98% of blocks", "18 of 18", True),
    ("A failed batch changes nothing", "goal: 100% of rejected batches", "108 of 108", True),
    ("Saving a 100-page document is fast", "goal: median under 5 seconds", "0.71 s", True),
    ("Outline is much smaller than the XML", "goal: at least 3x fewer tokens", "2.9x", False),
    ("The gate catches injected file faults", "goal: at least 95%, no false alarms", "83.5%", False),
    ("Revert restores the file exactly", "goal: every part, byte for byte", "whole file 2 of 300", False),
]
fig, ax = plt.subplots(figsize=(11, 7))
ax.set_xlim(0, 11); ax.set_ylim(0, 7); ax.axis("off")
ax.text(0.0, 6.65, "Benchmark of release 0.4.1: goals written before measuring", fontsize=16, fontweight="bold", color=INK)
ax.text(0.0, 6.28, "5 of 8 reached. The 3 misses are published with the rest.", fontsize=11.5, color=MUTED)
for i, (what, goal, got, ok) in enumerate(rows):
    y = 5.45 - i * 0.67
    fc, ec, col = (GREEN_L, GREEN, GREEN) if ok else (AMBER_L, AMBER, AMBER)
    ax.add_patch(FancyBboxPatch((0, y - 0.3), 11, 0.62, boxstyle="round,pad=0,rounding_size=0.06", fc="#f7f8fa", ec="none"))
    ax.add_patch(FancyBboxPatch((0.1, y - 0.14), 1.0, 0.34, boxstyle="round,pad=0,rounding_size=0.08", fc=fc, ec=ec, lw=1))
    ax.text(0.6, y + 0.03, "reached" if ok else "missed", ha="center", va="center", fontsize=9.5, fontweight="bold", color=col)
    ax.text(1.3, y + 0.12, what, va="center", fontsize=11.5, fontweight="bold", color=INK)
    ax.text(1.3, y - 0.15, goal, va="center", fontsize=9.5, color=MUTED)
    ax.text(10.9, y, got, va="center", ha="right", fontsize=13, fontweight="bold", color=col)
save(fig, "08-scoreboard.png")
print("done")
