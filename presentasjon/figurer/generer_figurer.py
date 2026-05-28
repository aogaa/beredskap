"""Generer fire visualiseringer for beredskapsplan-dokumentet."""
import sys
sys.stdout.reconfigure(encoding="utf-8")

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Circle
import networkx as nx
import numpy as np

# Designprofil — solid, alvor, seriøsitet
COLOR_PRIMARY = "#1A2B47"      # Mørk marineblå
COLOR_ACCENT = "#7B2D2F"       # Mørk burgunder — høy prioritet
COLOR_SECONDARY = "#2D4F3D"    # Mørk skogsgrønn — etablerte ressurser
COLOR_NEUTRAL = "#4A5568"      # Grå
COLOR_BG = "#F5F5F0"           # Lys beige
COLOR_GRID = "#D0D0CB"

# Felles font-konfigurasjon
plt.rcParams["font.family"] = "serif"
plt.rcParams["font.serif"] = ["Cambria", "Georgia", "DejaVu Serif"]
plt.rcParams["axes.edgecolor"] = COLOR_PRIMARY
plt.rcParams["axes.labelcolor"] = COLOR_PRIMARY
plt.rcParams["xtick.color"] = COLOR_PRIMARY
plt.rcParams["ytick.color"] = COLOR_PRIMARY
plt.rcParams["axes.titlecolor"] = COLOR_PRIMARY
plt.rcParams["font.weight"] = "regular"


# =====================================================================
# FIGUR 1 — Prioriteringsmatrise (viktighet × kompleksitet)
# Med visuell jitter for tiltak med identiske verdier
# =====================================================================
def figur_prioriteringsmatrise():
    # Tiltak med visuell offset for å unngå overlapp
    # (kode, navn, viktighet, kompleksitet, offset_x, offset_y, høy_prioritet)
    tiltak_plot = [
        ("T1", "Lokal møteplass",                  5, 3.5,  0.0,  0.0,  False),
        ("T2", "Beredskapsfrivillige",             5, 3,    0.0,  0.0,  False),
        ("T3", "Register sårbare beboere",         5, 4,    0.0,  0.0,  False),
        ("T8", "Koordinering spontanfrivillige",   5, 2,    0.0,  0.0,  True),
        # T4-T7 alle på (4, 2) — spred i 2x2 mini-grid
        ("T4", "Beredskapskurs befolkning",        4, 2,   -0.12,  0.13, False),
        ("T5", "Førstehjelp befolkning",           4, 2,    0.12,  0.13, False),
        ("T6", "Beredskapsvenner og nabolag",      4, 2,   -0.12, -0.13, False),
        ("T7", "Kompetansekart",                   4, 2,    0.12, -0.13, False),
        # T9, T10
        ("T9", "Ungdom og skole",                  3, 3,    0.0,   0.0,  False),
        ("T10","Flerspråklig beredskap",           3, 2,    0.0,   0.0,  False),
    ]

    fig, ax = plt.subplots(figsize=(12, 8.5), facecolor="white")
    ax.set_facecolor(COLOR_BG)

    # Kvadrant-bakgrunner
    ax.axhspan(4.5, 5.6, xmin=0, xmax=0.46,
               facecolor=COLOR_SECONDARY, alpha=0.10)
    ax.axhspan(4.5, 5.6, xmin=0.46, xmax=1,
               facecolor=COLOR_PRIMARY, alpha=0.08)
    ax.axhspan(0.5, 4.5, xmin=0, xmax=0.46,
               facecolor=COLOR_NEUTRAL, alpha=0.05)

    # Akselinjer
    ax.axhline(y=4.5, color=COLOR_NEUTRAL, linestyle="--", linewidth=0.8, alpha=0.6)
    ax.axvline(x=2.7, color=COLOR_NEUTRAL, linestyle="--", linewidth=0.8, alpha=0.6)

    # Plot hvert tiltak
    for kode, navn, vikt, kompl, ox, oy, hp in tiltak_plot:
        color = COLOR_ACCENT if hp else COLOR_PRIMARY
        size = 1100 if hp else 850
        edgewidth = 2.5 if hp else 1.5

        x = kompl + ox
        y = vikt + oy

        ax.scatter(x, y, s=size, c=color,
                   edgecolors="white", linewidth=edgewidth,
                   zorder=3)
        ax.text(x, y, kode, color="white",
                ha="center", va="center", fontsize=10,
                fontweight="bold", zorder=4)

    # Tekstforklaring til høyre for grafen — alle tiltakene med navn
    legend_y = 5.45
    ax.text(5.6, legend_y, "Tiltak:", fontsize=11, fontweight="bold",
            color=COLOR_PRIMARY, ha="left")
    tiltak_navn_liste = [
        ("T1", "Lokal møteplass"),
        ("T2", "Beredskapsfrivillige"),
        ("T3", "Register sårbare beboere"),
        ("T4", "Beredskapskurs befolkning"),
        ("T5", "Førstehjelp befolkning"),
        ("T6", "Beredskapsvenner og nabolag"),
        ("T7", "Kompetansekart"),
        ("T8", "Koordinering spontanfrivillige"),
        ("T9", "Ungdom og skole"),
        ("T10", "Flerspråklig beredskap"),
    ]
    for i, (kode, navn) in enumerate(tiltak_navn_liste):
        y_pos = 5.15 - i * 0.42
        color = COLOR_ACCENT if kode == "T8" else COLOR_PRIMARY
        ax.text(5.6, y_pos, kode + " —",
                fontsize=10, color=color, fontweight="bold", ha="left")
        ax.text(6.05, y_pos, navn,
                fontsize=10, color=COLOR_PRIMARY, ha="left")

    # Kvadrant-etiketter
    ax.text(1.5, 5.45, "STRATEGISKE RASKE GEVINSTER",
            fontsize=11, color=COLOR_SECONDARY, fontweight="bold",
            ha="center", style="italic")
    ax.text(3.85, 5.45, "STRATEGISK GRUNNLAG",
            fontsize=11, color=COLOR_PRIMARY, fontweight="bold",
            ha="center", style="italic")
    ax.text(1.5, 0.75, "BREDDE-TILTAK",
            fontsize=11, color=COLOR_PRIMARY, fontweight="bold",
            ha="center", style="italic")
    ax.text(3.85, 0.75, "LANGSIKTIGE SATSINGER",
            fontsize=11, color=COLOR_PRIMARY, fontweight="bold",
            ha="center", style="italic")

    # Akser
    ax.set_xlim(0.3, 5.2)
    ax.set_ylim(0.3, 5.7)
    ax.set_xlabel("Kompleksitet  →  (1 = enkel, 5 = svært krevende)",
                  fontsize=12.5, color=COLOR_PRIMARY, labelpad=10, fontweight="bold")
    ax.set_ylabel("Viktighet  →  (1 = lav effekt, 5 = avgjørende)",
                  fontsize=12.5, color=COLOR_PRIMARY, labelpad=10, fontweight="bold")
    ax.set_xticks([1, 2, 3, 4, 5])
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.tick_params(axis="both", labelsize=11)
    ax.grid(True, color=COLOR_GRID, linewidth=0.6, alpha=0.7)
    ax.set_axisbelow(True)

    ax.set_title("Prioriteringsmatrise — alle ti tiltak",
                 fontsize=15, color=COLOR_PRIMARY,
                 fontweight="bold", pad=15)

    # Legend nederst
    high_prio_patch = mpatches.Patch(color=COLOR_ACCENT,
                                     label="Høy prioritet (T8 — ofte glemt)")
    std_patch = mpatches.Patch(color=COLOR_PRIMARY,
                               label="Ordinært tiltak")
    leg = ax.legend(handles=[high_prio_patch, std_patch],
              loc="lower left", framealpha=0.97,
              fontsize=10.5, edgecolor=COLOR_PRIMARY)
    for txt in leg.get_texts():
        txt.set_color(COLOR_PRIMARY)

    for spine in ax.spines.values():
        spine.set_color(COLOR_PRIMARY)
        spine.set_linewidth(1.3)

    plt.tight_layout()
    plt.savefig("C:/codex/beredskap/presentasjon/figurer/01_prioriteringsmatrise.png",
                dpi=200, bbox_inches="tight", facecolor="white")
    plt.close()
    print("OK Figur 1: Prioriteringsmatrise")


# =====================================================================
# FIGUR 2 — Avhengighetsgraf (forenklet — kun primære avhengigheter)
# =====================================================================
def figur_avhengighetsgraf():
    fig, ax = plt.subplots(figsize=(13, 9.5), facecolor="white")
    ax.set_facecolor("white")

    G = nx.DiGraph()

    # Tre-lag struktur — T1 i sentrum, så ytre lag
    pos = {
        "T1": (0, 0),         # SENTRUM: møteplass
        # Indre ring — direkte avhengig av T1 / driver T1
        "T2": (-1.9, 1.2),
        "T3": (-1.9, -1.2),
        "T7": (1.9, 1.2),
        "T8": (1.9, -1.2),
        # Mellomring
        "T6": (-3.6, 0),
        "T4": (0, 2.8),
        "T5": (3.6, 0),
        # Ytre ring
        "T9": (-2.2, -2.8),
        "T10": (2.2, -2.8),
    }

    labels = {
        "T1": "T1\nMøteplass",
        "T2": "T2\nBeredskaps-\nfrivillige",
        "T3": "T3\nRegister\nsårbare",
        "T7": "T7\nKompetanse-\nkart",
        "T8": "T8\nSpontan-\nfrivillige",
        "T6": "T6\nBeredskaps-\nvenner",
        "T4": "T4\nBeredskaps-\nkurs",
        "T5": "T5\nFørste-\nhjelp",
        "T9": "T9\nUngdom\nog skole",
        "T10": "T10\nFlerspråklig",
    }

    # Forenklet — bare hovedavhengigheter
    edges = [
        # Indre ring → T1 (alle aktiveres fra møteplassen)
        ("T2", "T1"), ("T3", "T1"), ("T7", "T1"), ("T8", "T1"),
        # T8 og T3 går via T2
        ("T8", "T2"), ("T3", "T2"),
        # T7 forsyner T2 og T3
        ("T7", "T2"), ("T7", "T3"), ("T7", "T8"),
        # Mellomring kobler til indre
        ("T6", "T2"), ("T6", "T3"), ("T6", "T5"),
        ("T4", "T2"), ("T4", "T6"),
        ("T5", "T1"), ("T5", "T4"),
        # Ytre ring
        ("T9", "T4"), ("T9", "T2"),
        ("T10", "T7"), ("T10", "T6"),
    ]
    G.add_edges_from(edges)

    # Nodefarger og størrelser
    node_colors = []
    node_sizes = []
    for n in pos.keys():
        if n == "T8":
            node_colors.append(COLOR_ACCENT)
            node_sizes.append(4800)
        elif n == "T1":
            node_colors.append(COLOR_PRIMARY)
            node_sizes.append(5800)
        elif n in ("T2", "T3", "T7"):
            node_colors.append(COLOR_PRIMARY)
            node_sizes.append(4500)
        elif n in ("T4", "T5", "T6"):
            node_colors.append(COLOR_NEUTRAL)
            node_sizes.append(4000)
        else:  # T9, T10
            node_colors.append(COLOR_NEUTRAL)
            node_sizes.append(3500)

    # Tegn kanter — mørkere og mer synlige
    nx.draw_networkx_edges(G, pos, ax=ax,
                           edge_color=COLOR_PRIMARY,
                           arrows=True, arrowsize=15,
                           arrowstyle="-|>",
                           connectionstyle="arc3,rad=0.10",
                           width=1.5, alpha=0.55,
                           node_size=node_sizes)

    # Tegn noder
    nx.draw_networkx_nodes(G, pos, ax=ax,
                           node_color=node_colors,
                           node_size=node_sizes,
                           edgecolors="white",
                           linewidths=2.5)

    # Tegn etiketter
    for node, (x, y) in pos.items():
        ax.text(x, y, labels[node],
                ha="center", va="center",
                fontsize=9, color="white",
                fontweight="bold")

    # Tittel
    ax.set_title("Avhengighetsgraf — hvordan de ti tiltakene henger sammen",
                 fontsize=15, color=COLOR_PRIMARY,
                 fontweight="bold", pad=15)

    # Forklarende tekst
    ax.text(0, -4.0,
            "Piler peker fra tiltaket mot det det er avhengig av.\n"
            "T1 (blå, sentrum) er grunninfrastrukturen. "
            "T8 (burgunder) har høy prioritet.",
            fontsize=11, color=COLOR_PRIMARY,
            ha="center", style="italic")

    ax.set_xlim(-5.0, 5.0)
    ax.set_ylim(-4.7, 3.6)
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    plt.savefig("C:/codex/beredskap/presentasjon/figurer/02_avhengighetsgraf.png",
                dpi=200, bbox_inches="tight", facecolor="white")
    plt.close()
    print("OK Figur 2: Avhengighetsgraf")


# =====================================================================
# FIGUR 3 — Tidslinje første år (Gantt-lignende)
# =====================================================================
def figur_tidslinje():
    aktiviteter = [
        ("Dialog med bydel og Oslo kommune", 1, 3, "forankring",
         "Møter med beredskapsansvarlig"),
        ("T1 Nivå 1 — møteplass etablert", 1, 3, "forankring",
         "Adresse og rolle kommunisert"),
        ("T3 Spor 1 — register startes opp", 1, 3, "forankring",
         "Registeransvarlig, samtykkemal"),
        ("T7 — kompetansekartlegging startes", 2, 4, "forankring",
         "Kompetanseansvarlig, første samtaler"),
        ("T2 — rekruttering beredskapsfrivillige", 3, 5, "oppbygging",
         "6-8 frivillige per møteplass"),
        ("T2 — grunnkurs (12-16 t)", 5, 3, "oppbygging",
         "I samarbeid med Oslo kommune"),
        ("T6 Nivå 1 — beredskapsvennkvelder", 4, 8, "oppbygging",
         "4-6 kvelder, borettslag"),
        ("T8 — mottakssjef opplært", 5, 2, "oppbygging",
         "Modul 7 i T2-grunnkurset"),
        ("T1 Nivå 2 — utstyrt møteplass", 6, 4, "oppbygging",
         "Beredskapslager etablert"),
        ("T4 — beredskapskurs befolkning", 8, 4, "aktivering",
         "4-6 kurs gjennomført"),
        ("T5 — førstehjelpskurs (HLR + AED)", 8, 4, "aktivering",
         "Minimum 4 kurs"),
        ("T9 — pilot ungdomsskole", 8, 5, "aktivering",
         "Innsats for andre"),
        ("T10 — språkpakke + Bydelsmødre", 9, 3, "aktivering",
         "3-5 språk identifisert"),
        ("Felles øvelse + evaluering", 11, 2, "aktivering",
         "T1/T2/T3/T8 testet"),
    ]

    fase_colors = {
        "forankring": COLOR_SECONDARY,
        "oppbygging": COLOR_PRIMARY,
        "aktivering": COLOR_ACCENT,
    }

    fig, ax = plt.subplots(figsize=(14, 8), facecolor="white")
    ax.set_facecolor("white")

    for i, (navn, start, dur, fase, beskr) in enumerate(aktiviteter):
        y = len(aktiviteter) - i - 1
        c = fase_colors[fase]
        ax.barh(y, dur, left=start, height=0.55,
                color=c, alpha=0.92,
                edgecolor="white", linewidth=1.5)
        if dur >= 3:
            ax.text(start + dur/2, y, beskr,
                    color="white", fontsize=9.5,
                    ha="center", va="center", style="italic",
                    fontweight="bold")
        ax.text(0.5, y, navn,
                color=COLOR_PRIMARY, fontsize=10.5,
                ha="right", va="center", fontweight="bold")

    måneder = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]
    ax.set_xticks([m + 0.5 for m in range(1, 13)])
    ax.set_xticklabels(måneder, fontsize=11, fontweight="bold")
    ax.set_xlim(-7, 13.2)
    ax.set_xlabel("Måned i første pilotår", fontsize=12.5,
                  color=COLOR_PRIMARY, labelpad=10, fontweight="bold")

    ax.axvline(x=3, color=COLOR_PRIMARY, linestyle="-", linewidth=1.2, alpha=0.5)
    ax.axvline(x=8, color=COLOR_PRIMARY, linestyle="-", linewidth=1.2, alpha=0.5)

    ax.text(2, len(aktiviteter) + 0.2, "Fase A\nForankring",
            ha="center", fontsize=11, fontweight="bold",
            color=COLOR_SECONDARY)
    ax.text(5.5, len(aktiviteter) + 0.2, "Fase B\nOppbygging",
            ha="center", fontsize=11, fontweight="bold",
            color=COLOR_PRIMARY)
    ax.text(10.5, len(aktiviteter) + 0.2, "Fase C\nAktivering",
            ha="center", fontsize=11, fontweight="bold",
            color=COLOR_ACCENT)

    ax.set_yticks([])
    ax.set_ylim(-0.8, len(aktiviteter) + 0.9)
    ax.grid(True, axis="x", color=COLOR_GRID, linewidth=0.6, alpha=0.6)
    ax.set_axisbelow(True)

    for side in ("top", "left", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(COLOR_PRIMARY)
    ax.spines["bottom"].set_linewidth(1.3)

    ax.set_title("Tidslinje for første pilotår — sekvens og fasing av tiltak",
                 fontsize=15, color=COLOR_PRIMARY,
                 fontweight="bold", pad=20, x=0.35)

    plt.tight_layout()
    plt.savefig("C:/codex/beredskap/presentasjon/figurer/03_tidslinje.png",
                dpi=200, bbox_inches="tight", facecolor="white")
    plt.close()
    print("OK Figur 3: Tidslinje")


# =====================================================================
# FIGUR 4 — Risiko-tiltak-matrise (med matplotlib Circle, ikke font-symboler)
# =====================================================================
def figur_risiko_matrise():
    risikoer = [
        "Bortfall av\nvannforsyning",
        "Strømbortfall /\nrasjonering",
        "Cyberangrep\nmot ekom",
        "Pandemi",
    ]

    tiltak_koder = ["T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10"]

    # 2 = sterk effekt, 1 = indirekte effekt
    matrise = np.array([
        # T1  T2  T3  T4  T5  T6  T7  T8  T9  T10
        [  2,  2,  2,  2,  1,  2,  2,  2,  1,  2],  # Vannforsyning
        [  2,  2,  2,  2,  1,  2,  1,  2,  1,  1],  # Strømbortfall
        [  2,  2,  2,  1,  1,  2,  2,  2,  1,  2],  # Cyberangrep ekom
        [  1,  2,  2,  1,  2,  2,  2,  2,  1,  2],  # Pandemi
    ])

    fig, ax = plt.subplots(figsize=(13, 6), facecolor="white")
    ax.set_facecolor(COLOR_BG)

    n_rows, n_cols = matrise.shape

    # Rist-bakgrunn
    for i in range(n_rows):
        for j in range(n_cols):
            v = matrise[i, j]
            # Hver celle som rektangel
            rect_color = "white"
            ax.add_patch(mpatches.Rectangle(
                (j - 0.5, i - 0.5), 1, 1,
                facecolor=rect_color,
                edgecolor=COLOR_GRID,
                linewidth=1.0))

            # Sirkel i hver celle som symboliserer effekt
            if v == 2:
                # Stor solid sirkel — sterk effekt
                ax.add_patch(Circle((j, i), 0.28,
                                    facecolor=COLOR_PRIMARY,
                                    edgecolor=COLOR_PRIMARY,
                                    linewidth=1.0,
                                    zorder=3))
            elif v == 1:
                # Liten åpen sirkel — indirekte effekt
                ax.add_patch(Circle((j, i), 0.18,
                                    facecolor="white",
                                    edgecolor=COLOR_PRIMARY,
                                    linewidth=2,
                                    zorder=3))

    # Akser
    ax.set_xticks(np.arange(n_cols))
    ax.set_yticks(np.arange(n_rows))
    ax.set_xticklabels(tiltak_koder, fontsize=13, color=COLOR_PRIMARY,
                       fontweight="bold")
    ax.set_yticklabels(risikoer, fontsize=12, color=COLOR_PRIMARY,
                       fontweight="bold")
    ax.set_xlim(-0.55, n_cols - 0.45)
    ax.set_ylim(-0.55, n_rows - 0.45)
    ax.invert_yaxis()
    ax.set_aspect("equal")

    # Skjul spines
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(axis="both", which="major", length=0)

    # Tittel
    ax.set_title("Risiko-tiltak-matrise — KRB Oslo 2025 toppscenarioer mot de ti tiltakene",
                 fontsize=14, color=COLOR_PRIMARY,
                 fontweight="bold", pad=15)

    # Forklaring til høyre (egen aksiform med små sirkler)
    # Stor sirkel = sterk effekt
    legend_x = n_cols + 0.4
    ax.add_patch(Circle((legend_x, 0.5), 0.22,
                        facecolor=COLOR_PRIMARY,
                        edgecolor=COLOR_PRIMARY,
                        linewidth=1.0,
                        clip_on=False, zorder=10))
    ax.text(legend_x + 0.45, 0.5, "Sterk effekt",
            fontsize=11, color=COLOR_PRIMARY,
            va="center", ha="left")

    # Liten sirkel = indirekte
    ax.add_patch(Circle((legend_x, 1.4), 0.14,
                        facecolor="white",
                        edgecolor=COLOR_PRIMARY,
                        linewidth=2,
                        clip_on=False, zorder=10))
    ax.text(legend_x + 0.45, 1.4, "Indirekte effekt",
            fontsize=11.5, color=COLOR_PRIMARY,
            va="center", ha="left", fontweight="bold")

    plt.tight_layout()
    plt.savefig("C:/codex/beredskap/presentasjon/figurer/04_risiko_matrise.png",
                dpi=200, bbox_inches="tight", facecolor="white")
    plt.close()
    print("OK Figur 4: Risiko-tiltak-matrise")


if __name__ == "__main__":
    figur_prioriteringsmatrise()
    figur_avhengighetsgraf()
    figur_tidslinje()
    figur_risiko_matrise()
    print("\nAlle fire figurer generert i C:/codex/beredskap/presentasjon/figurer/")
