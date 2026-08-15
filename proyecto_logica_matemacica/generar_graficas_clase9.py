# -*- coding: utf-8 -*-
"""Genera las graficas del Deber de la Clase 9.1 (Analisis de Funciones)."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os

NAVY = '#1B365D'
TEAL = '#0D9488'
RED = '#EF4444'
GREY = '#94A3B8'

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'graficas_clase9')
os.makedirs(OUT, exist_ok=True)


def base_axes(ax, xlim, ylim, xlabel='x', ylabel='y'):
    ax.axhline(0, color=GREY, lw=1.2, zorder=1)
    ax.axvline(0, color=GREY, lw=1.2, zorder=1)
    ax.grid(True, ls=':', lw=0.7, color='#CBD5E1', zorder=0)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_xlabel(xlabel, fontsize=10, color='#334155')
    ax.set_ylabel(ylabel, fontsize=10, color='#334155')
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(colors='#64748B', labelsize=9)


def frange(a, b, step):
    out, v = [], a
    while v <= b + 1e-9:
        out.append(round(v, 6))
        v += step
    return out


# ---------- PARTE A: f(x) = x^2 - 2x - 3 ----------
fig, ax = plt.subplots(figsize=(6.2, 4.4), dpi=160)
xs = frange(-3, 5, 0.02)
ys = [x * x - 2 * x - 3 for x in xs]
base_axes(ax, (-3, 5), (-6, 8))
ax.plot(xs, ys, color=NAVY, lw=2.6, zorder=3, label='f(x) = x² − 2x − 3')
ax.plot([-1, 3], [0, 0], 'o', color=RED, ms=8, zorder=5)
ax.annotate('(-1, 0)', (-1, 0), textcoords='offset points', xytext=(-38, 10),
            fontsize=9.5, color='#991B1B', fontweight='bold')
ax.annotate('(3, 0)', (3, 0), textcoords='offset points', xytext=(8, 10),
            fontsize=9.5, color='#991B1B', fontweight='bold')
ax.plot([1], [-4], 'o', color=TEAL, ms=9, zorder=5)
ax.annotate('V(1, −4)  mínimo', (1, -4), textcoords='offset points', xytext=(12, -16),
            fontsize=9.5, color='#0F766E', fontweight='bold')
ax.plot([0], [-3], 'o', color='#64748B', ms=6, zorder=5)
ax.annotate('(0, −3)', (0, -3), textcoords='offset points', xytext=(-48, -4),
            fontsize=8.5, color='#475569')
ax.axhline(-4, color=TEAL, ls='--', lw=1, alpha=0.5, zorder=2)
ax.annotate('Rango: [−4, ∞)', (-2.8, -4), textcoords='offset points', xytext=(0, 6),
            fontsize=8.5, color='#0F766E')
ax.axvline(1, color=TEAL, ls=':', lw=1, alpha=0.6, zorder=2)
ax.annotate('decrece', (-0.6, 5.2), fontsize=9, color='#475569', ha='center')
ax.annotate('crece', (2.7, 5.2), fontsize=9, color='#475569', ha='center')
ax.legend(loc='upper center', fontsize=9.5, frameon=False)
ax.set_title('Parte A — f(x) = x² − 2x − 3', fontsize=11.5, color=NAVY,
             fontweight='bold', pad=10)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'parte_A.png'), bbox_inches='tight', facecolor='white')
plt.close(fig)

# ---------- PARTE B: f(x) = -3x + 9 ----------
fig, ax = plt.subplots(figsize=(6.2, 4.4), dpi=160)
xs = frange(-1, 6, 0.05)
ys = [-3 * x + 9 for x in xs]
base_axes(ax, (-1, 6), (-9, 13))
ax.plot(xs, ys, color=NAVY, lw=2.6, zorder=3, label='f(x) = −3x + 9')
ax.plot([3], [0], 'o', color=RED, ms=8, zorder=5)
ax.annotate('Raíz (3, 0)', (3, 0), textcoords='offset points', xytext=(10, 10),
            fontsize=9.5, color='#991B1B', fontweight='bold')
ax.plot([0], [9], 'o', color='#64748B', ms=6, zorder=5)
ax.annotate('(0, 9)', (0, 9), textcoords='offset points', xytext=(8, 4),
            fontsize=9, color='#475569')
ax.annotate('m = −3 < 0  →  decreciente', (4.2, 6.5), fontsize=9.5, color='#0F766E',
            fontweight='bold', ha='center')
ax.set_title('Parte B — f(x) = −3x + 9', fontsize=11.5, color=NAVY,
             fontweight='bold', pad=10)
ax.legend(loc='lower left', fontsize=9.5, frameon=False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'parte_B.png'), bbox_inches='tight', facecolor='white')
plt.close(fig)

# ---------- PARTE C: h(t) = -t^2 + 6t ----------
fig, ax = plt.subplots(figsize=(6.2, 4.4), dpi=160)
ts = frange(-0.4, 6.6, 0.02)
hs = [-t * t + 6 * t for t in ts]
base_axes(ax, (-0.9, 7.2), (-4.5, 11.5), xlabel='t (segundos)', ylabel='h (metros)')
ax.plot(ts, hs, color=NAVY, lw=2.6, zorder=3, label='h(t) = −t² + 6t')
ts_real = frange(0, 6, 0.02)
ax.fill_between(ts_real, [0] * len(ts_real), [-t * t + 6 * t for t in ts_real],
                color=TEAL, alpha=0.10, zorder=2)
ax.plot([0, 6], [0, 0], 'o', color=RED, ms=8, zorder=5)
bbox_w = dict(boxstyle='round,pad=0.3', fc='white', ec='none', alpha=0.92)
ax.annotate('t = 0 s\n(despegue)', (0, 0), textcoords='offset points', xytext=(-30, -44),
            fontsize=9, color='#991B1B', fontweight='bold', ha='center', bbox=bbox_w,
            zorder=8)
ax.annotate('t = 6 s\n(vuelve al suelo)', (6, 0), textcoords='offset points',
            xytext=(26, -44), fontsize=9, color='#991B1B', fontweight='bold',
            ha='center', bbox=bbox_w, zorder=8)
ax.plot([3], [9], 'o', color=TEAL, ms=9, zorder=5)
ax.annotate('V(3, 9)  altura máxima', (3, 9), textcoords='offset points', xytext=(0, 14),
            fontsize=9.5, color='#0F766E', fontweight='bold', ha='center')
ax.axvline(3, color=TEAL, ls=':', lw=1, alpha=0.6, zorder=2)
ax.annotate('sube', (1.5, 2.0), fontsize=9, color='#475569', ha='center')
ax.annotate('baja', (4.5, 2.0), fontsize=9, color='#475569', ha='center')
ax.set_title('Parte C — Altura del dron:  h(t) = −t² + 6t', fontsize=11.5, color=NAVY,
             fontweight='bold', pad=10)
ax.legend(loc='upper right', fontsize=9.5, frameon=False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'parte_C.png'), bbox_inches='tight', facecolor='white')
plt.close(fig)

# ---------- PARTE D: U(x) = -2500x^2 + 23550x - 20000 ----------
fig, ax = plt.subplots(figsize=(6.8, 4.6), dpi=160)
xs = frange(-0.3, 9.3, 0.02)
us = [-2500 * x * x + 23550 * x - 20000 for x in xs]
base_axes(ax, (-0.3, 9.4), (-25000, 46000),
          xlabel='x  (contenedores de 25 t exportados al mes)',
          ylabel='U(x)  (utilidad neta en USD)')
ax.plot(xs, us, color=NAVY, lw=2.6, zorder=3, label='U(x) = −2500x² + 23550x − 20000')

xs_pos = frange(0.943822, 8.476178, 0.01)
ax.fill_between(xs_pos, [0] * len(xs_pos),
                [-2500 * x * x + 23550 * x - 20000 for x in xs_pos],
                color=TEAL, alpha=0.10, zorder=2)

ax.plot([0.9438, 8.4762], [0, 0], 'o', color=RED, ms=8, zorder=5)
ax.annotate('x ≈ 0.94', (0.9438, 0), textcoords='offset points', xytext=(-8, -26),
            fontsize=9, color='#991B1B', fontweight='bold', ha='center')
ax.annotate('x ≈ 8.48', (8.4762, 0), textcoords='offset points', xytext=(4, -26),
            fontsize=9, color='#991B1B', fontweight='bold', ha='center')

ax.plot([4.71], [35460.25], 'o', color=TEAL, ms=9, zorder=6)
ax.annotate('V(4.71 ; $35,460)', (4.71, 35460.25), textcoords='offset points',
            xytext=(0, 12), fontsize=9.5, color='#0F766E', fontweight='bold', ha='center')

ints = list(range(1, 9))
ax.plot(ints, [-2500 * x * x + 23550 * x - 20000 for x in ints], 's',
        color='#1E40AF', ms=5.5, zorder=5, label='Dominio real: contenedores enteros')
ax.plot([5], [35250], 's', color='#F59E0B', ms=10, zorder=7,
        label='Óptimo real: 5 contenedores → $35,250')

ax.axvline(4.71, color=TEAL, ls=':', lw=1, alpha=0.6, zorder=2)
ax.annotate('crece', (2.6, 40500), fontsize=9, color='#475569', ha='center')
ax.annotate('decrece', (7.0, 40500), fontsize=9, color='#475569', ha='center')
ax.set_title('Parte D — Utilidad mensual de la exportación de cacao',
             fontsize=11.5, color=NAVY, fontweight='bold', pad=10)
ax.legend(loc='lower center', fontsize=8.5, frameon=False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'parte_D.png'), bbox_inches='tight', facecolor='white')
plt.close(fig)

print('Graficas generadas en:', OUT)
for f in sorted(os.listdir(OUT)):
    print('  -', f)
