"""
ANÀLISI ESTADÍSTIC V2H — v3
Estructura cohesionada amb degradació i FV→VE integrats.
Payback principal = amb degradació (valor real). Matriu de decisió = amb degradació.
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from pathlib import Path
import pandas as pd, numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from scipy import stats
import warnings; warnings.filterwarnings('ignore')

_DIR = Path(__file__).parent
CARPETA = _DIR / "graficsAE"
os.makedirs(CARPETA, exist_ok=True)

def llegir(f): return pd.read_csv(f, sep=';', decimal=',', encoding='utf-8-sig')
df      = llegir(_DIR / "Minitab_Factorial.csv")
df_pis  = llegir(_DIR / "Minitab_PIS.csv")
df_casa = llegir(_DIR / "Minitab_CASA.csv")
df_dec  = llegir(_DIR / "Minitab_Decisio_Payback.csv")

LLINDAR = 8
VIDA_UTIL = 10

def recomanacio(row):
    pb_dif = row['PB_Difer_Deg_anys']
    pb_uni = row['Payback_Unidir_anys']
    estalvi_net = row['Estalvi_Net_EUR']
    if pb_dif < LLINDAR:
        return 'Bidireccional'
    elif pb_uni < LLINDAR:
        return 'Unidireccional'
    elif estalvi_net > 0:
        return 'Unidir. (retorn lent)'
    else:
        return 'No rendible'

df_dec['Recomanacio'] = df_dec.apply(recomanacio, axis=1)

pv = df_pis[df_pis['Carrega'] != 'Publica'].copy()
cv = df_casa[df_casa['Carrega'] != 'Publica'].copy()
dv = df[df['Carrega'] != 'Publica'].copy()

plt.rcParams.update({
    'font.family': 'sans-serif', 'font.sans-serif': ['Arial', 'DejaVu Sans'],
    'font.size': 10, 'axes.titlesize': 12, 'axes.labelsize': 10,
    'figure.dpi': 180, 'savefig.dpi': 180,
    'axes.grid': True, 'grid.alpha': 0.2,
    'figure.constrained_layout.use': True,
})
C_PIS = '#3B8BD4'; C_CASA = '#D85A30'; C_FV = '#5DCA5D'
C_BIDIR = '#2E75B6'; C_UNIDIR = '#5DCA5D'; C_DIFER = '#EF9F27'
C_DEG = '#9B59B6'

def guardar(nom):
    plt.savefig(os.path.join(CARPETA, nom), bbox_inches='tight', pad_inches=0.15)
    plt.close()
    print(f"  ✓ {nom}")

def mt(s):
    """Subíndex ₂ via mathtext (Arial no té U+2082 i surt com a requadre)."""
    return s.replace('₂', '$_2$')

def anova_factors(df_in, factors, variable):
    d = df_in[factors + [variable]].dropna().copy()
    # El filtre del sentinella 999 només s'aplica a paybacks (no a estalvis,
    # que poden superar legítimament els 900 €)
    if 'PB' in variable or 'Payback' in variable:
        d = d[d[variable] < 900]
    if len(d) < 4:
        return None
    grand = d[variable].mean()
    ss_tot = ((d[variable] - grand) ** 2).sum()
    if ss_tot < 1e-9:
        return None  # variable constant (p. ex. FV→VE al PIS): ANOVA degenerada
    rows = []; ss_res = ss_tot
    for fac in factors:
        grps = d.groupby(fac)[variable]
        ss_b = sum(len(g) * (g.mean() - grand) ** 2 for _, g in grps)
        df_b = grps.ngroups - 1
        rows.append({'Font': fac, 'SS': ss_b, 'df': df_b})
        ss_res -= ss_b
    df_res = len(d) - 1 - sum(r['df'] for r in rows)
    if df_res <= 0:
        df_res = 1
    ms_res = ss_res / df_res
    for r in rows:
        r['MS'] = r['SS'] / r['df'] if r['df'] > 0 else 0
        r['F'] = r['MS'] / ms_res if ms_res > 0 else 0
        r['p'] = 1 - stats.f.cdf(r['F'], r['df'], df_res) if r['F'] > 0 else 1
        r['eta2'] = r['SS'] / ss_tot if ss_tot > 0 else 0
    rows.append({'Font': 'Residual', 'SS': ss_res, 'df': df_res, 'MS': ms_res,
                 'F': np.nan, 'p': np.nan, 'eta2': ss_res / ss_tot})
    rows.append({'Font': 'Total', 'SS': ss_tot, 'df': len(d) - 1, 'MS': np.nan,
                 'F': np.nan, 'p': np.nan, 'eta2': 1.0})
    return pd.DataFrame(rows)

def p_str(p):
    if np.isnan(p): return '—'
    if p < 0.001: return '<0,001***'
    if p < 0.01:  return f'{p:.3f}**'.replace('.', ',')
    if p < 0.05:  return f'{p:.3f}*'.replace('.', ',')
    return f'{p:.3f}'.replace('.', ',')

print("=" * 60); print("A · ANOVA + η²"); print("=" * 60)

kpis = {
    'Estalvi_V2H_EUR':      'Estalvi V2H (€)',
    'Estalvi_Total_EUR':     'Estalvi total (€)',
    'Estalvi_FV_VE_EUR':     'Estalvi FV→VE (€/any)',
    'Autoconsum_pct':        'Autoconsum (%)',
    'EFC_V2H':               'EFC V2H (cicles/any)',
    'Cost_Degradacio_EUR':   'Cost degradació (€/any)',
    'Carg_FV_VE_kWh':        'Càrrega FV→VE (kWh/any)',
    'tCO2_Anual':            'Emissions CO₂ (tCO₂/any)',
}

all_anova = []
for var, nom in kpis.items():
    t = anova_factors(df_pis, ['Mobilitat', 'Carrega'], var)
    if t is not None:
        t['Variable'] = nom; t['Cas'] = 'PIS'; all_anova.append(t)
    t = anova_factors(df_casa, ['FV', 'Mobilitat', 'Carrega'], var)
    if t is not None:
        t['Variable'] = nom; t['Cas'] = 'CASA'; all_anova.append(t)

df_anova = pd.concat(all_anova, ignore_index=True)
df_anova['p_str'] = df_anova['p'].apply(p_str)
df_anova['eta2_pct'] = (df_anova['eta2'] * 100).round(1)
df_anova.to_csv(CARPETA / "T1_ANOVA_Complet.csv", sep=';', decimal=',',
                index=False, encoding='utf-8-sig')
print("  ✓ T1_ANOVA_Complet.csv")

# CSV resum complet dels 27 escenaris
df[['Habitatge', 'FV', 'Mobilitat', 'Carrega',
    'Estalvi_V2H_EUR', 'Estalvi_FV_VE_EUR', 'Estalvi_Total_EUR', 'Estalvi_Net_EUR',
    'Autoconsum_pct', 'EFC_V2H', 'Cost_Degradacio_EUR', 'Carg_FV_VE_kWh',
    'tCO2_Anual',
    'PB_Bidir_Deg_anys', 'Payback_Unidir_anys', 'PB_Difer_Deg_anys'
]].to_csv(CARPETA / "T0_Resum_27_Escenaris.csv", sep=';', decimal=',',
          index=False, encoding='utf-8-sig')
print("  ✓ T0_Resum_27_Escenaris.csv")

for cas, df_cas, facs in [("PIS", df_pis, ['Mobilitat', 'Carrega']),
                           ("CASA", df_casa, ['FV', 'Mobilitat', 'Carrega'])]:
    mat = {}
    for var, nom in kpis.items():
        t = anova_factors(df_cas, facs, var)
        if t is None:
            continue
        for _, r in t.iterrows():
            if r['Font'] in facs:
                mat.setdefault(nom, {})[r['Font']] = r['eta2']
    if not mat:
        continue
    eta_df = pd.DataFrame(mat).T
    eta_df = eta_df[facs]
    fig, ax = plt.subplots(figsize=(max(5, 2.8 * len(facs)),
                                     max(3.5, 0.8 * len(eta_df))))
    im = ax.imshow(eta_df.values, cmap='YlOrRd', vmin=0, vmax=1, aspect='auto')
    ax.set_xticks(range(len(eta_df.columns)))
    ax.set_xticklabels(eta_df.columns, fontsize=11)
    ax.set_yticks(range(len(eta_df.index)))
    ax.set_yticklabels([mt(x) for x in eta_df.index], fontsize=10)
    for i in range(len(eta_df.index)):
        for j in range(len(eta_df.columns)):
            v = eta_df.values[i, j]
            if not np.isnan(v):
                ax.text(j, i, f'{v:.1%}', ha='center', va='center',
                        fontsize=11, fontweight='bold',
                        color='white' if v > 0.45 else 'black')
    cb = plt.colorbar(im, ax=ax, shrink=0.8)
    cb.set_label('η²', fontsize=11)
    ax.set_title(f'CAS {cas} — η² (proporció de variància explicada per factor)',
                 fontsize=13, pad=12)
    guardar(f"A_{cas}_Eta2_Heatmap.png")

print("\n" + "=" * 60); print("B · Efectes principals"); print("=" * 60)

# Només variables no redundants (cost deg. ≡ EFC, estalvi FV→VE ≡ FV→VE kWh,
# estalvi net ≈ estalvi V2H: ρ = 1,00 entre parelles — vegeu secció H).
# Al PIS s'ometen les variables FV→VE (constants a zero).
vars_ep_cas = {
    'PIS':  [('Estalvi_V2H_EUR', 'Estalvi V2H (€)'),
             ('Autoconsum_pct',  'Autoconsum (%)'),
             ('EFC_V2H',         'EFC (cicles/any)'),
             ('tCO2_Anual',      'CO$_2$ (tCO$_2$)')],
    'CASA': [('Estalvi_V2H_EUR', 'Estalvi V2H (€)'),
             ('Autoconsum_pct',  'Autoconsum (%)'),
             ('EFC_V2H',         'EFC (cicles/any)'),
             ('Carg_FV_VE_kWh',  'FV→VE (kWh)'),
             ('tCO2_Anual',      'CO$_2$ (tCO$_2$)')],
}

for cas, df_cas, facs in [("PIS", df_pis, ['Mobilitat', 'Carrega']),
                           ("CASA", df_casa, ['FV', 'Mobilitat', 'Carrega'])]:
    vars_ep = [v for v, _ in vars_ep_cas[cas]]
    noms_ep = [n for _, n in vars_ep_cas[cas]]
    nf = len(facs); nv = len(vars_ep)
    fig, axes = plt.subplots(nv, nf, figsize=(4.2 * nf, 2.5 * nv), squeeze=False)
    fig.suptitle(f'CAS {cas} — Efectes principals dels factors', fontsize=14)
    fig.set_constrained_layout_pads(hspace=0.12, wspace=0.05)
    for i, (var, nom) in enumerate(zip(vars_ep, noms_ep)):
        sub = df_cas[df_cas[var] < 900]
        for j, fac in enumerate(facs):
            ax = axes[i][j]
            t = anova_factors(df_cas, facs, var)
            eta_txt = ''
            if t is not None:
                row = t[t['Font'] == fac]
                if len(row):
                    eta_txt = f'η²={row.iloc[0]["eta2"]:.1%}'
            means = sub.groupby(fac)[var].mean()
            col = {'Carrega': C_BIDIR, 'FV': C_CASA, 'Mobilitat': C_PIS}.get(fac, C_BIDIR)
            ax.plot(range(len(means)), means.values, 'o-', color=col, markersize=7, lw=2)
            ax.set_xticks(range(len(means)))
            ax.set_xticklabels(means.index, fontsize=8)
            ax.set_title(f'{fac}  ({eta_txt})', fontsize=9, pad=4)
            if j == 0:
                ax.set_ylabel(nom, fontsize=9)
    guardar(f"B_{cas}_Efectes_Principals.png")

print("\n" + "=" * 60); print("C · Heatmaps KPI"); print("=" * 60)

# Un únic panell 2×3 per tipologia amb els indicadors essencials i no redundants.
# Per a paybacks, els valors sentinella (999, càrrega pública) es mostren com '>40'.
PB_CAP = 40


def do_heatmaps(df_in, title, fname, hm_vars, hm_noms, hm_fmts, invert_cmap=None):
    """invert_cmap: set of indices where lower=better (costs, paybacks)."""
    if invert_cmap is None:
        invert_cmap = set()
    fig, axes = plt.subplots(2, 3, figsize=(15, 7.5))
    fig.suptitle(title, fontsize=14)
    for idx, (v, n, f) in enumerate(zip(hm_vars, hm_noms, hm_fmts)):
        ax = axes.flat[idx]
        piv = df_in.pivot_table(index='Mobilitat', columns='Carrega',
                                values=v, aggfunc='mean')
        is_pb = 'PB' in v or 'Payback' in v
        vals = piv.values.copy()
        if is_pb:
            vals = np.clip(vals, None, PB_CAP)
        cmap = 'RdYlGn_r' if idx in invert_cmap else 'RdYlGn'
        im = ax.imshow(vals, cmap=cmap, aspect='auto')
        ax.set_xticks(range(len(piv.columns)))
        ax.set_xticklabels(piv.columns, fontsize=9)
        ax.set_yticks(range(len(piv.index)))
        ax.set_yticklabels(piv.index, fontsize=9)
        for i_ in range(len(piv.index)):
            for j_ in range(len(piv.columns)):
                val = piv.values[i_, j_]
                rng = vals.max() - vals.min()
                if rng == 0:
                    rng = 1
                c = 'white' if abs(min(val, vals.max()) - vals.mean()) > 0.35 * rng else 'black'
                lbl = f'>{PB_CAP}' if (is_pb and val >= PB_CAP) else f'{val:{f}}'
                ax.text(j_, i_, lbl, ha='center', va='center',
                        fontsize=10, fontweight='bold', color=c)
        ax.set_title(n, fontsize=10, pad=6)
    guardar(fname)


hm_pis_vars = ['Estalvi_V2H_EUR', 'Estalvi_Net_EUR', 'Autoconsum_pct',
               'EFC_V2H', 'tCO2_Anual', 'PB_Difer_Deg_anys']
hm_pis_noms = ['Estalvi V2H (€)', 'Estalvi net (€)', 'Autoconsum (%)',
               'EFC (cicles/any)', 'CO$_2$ (tCO$_2$)', 'PB difer. deg. (anys)']
hm_pis_fmts = ['.1f', '.1f', '.1f', '.1f', '.3f', '.1f']

hm_cfv_vars = ['Estalvi_V2H_EUR', 'Estalvi_Net_EUR', 'Autoconsum_pct',
               'Carg_FV_VE_kWh', 'tCO2_Anual', 'PB_Difer_Deg_anys']
hm_cfv_noms = ['Estalvi V2H (€)', 'Estalvi net (€)', 'Autoconsum (%)',
               'FV→VE (kWh)', 'CO$_2$ (tCO$_2$)', 'PB difer. deg. (anys)']
hm_cfv_fmts = ['.1f', '.1f', '.1f', '.1f', '.3f', '.1f']

do_heatmaps(df_pis, 'CAS PIS — Indicadors clau (Mobilitat × Càrrega)',
            'C_PIS_HM.png', hm_pis_vars, hm_pis_noms, hm_pis_fmts,
            invert_cmap={4, 5})  # CO₂ i payback: lower=better

do_heatmaps(df_casa[df_casa['FV'] == 'Si'],
            'CAS CASA amb FV — Indicadors clau (Mobilitat × Càrrega)',
            'C_CASA_FV_HM.png', hm_cfv_vars, hm_cfv_noms, hm_cfv_fmts,
            invert_cmap={4, 5})

do_heatmaps(df_casa[df_casa['FV'] == 'No'],
            'CAS CASA sense FV — Indicadors clau (Mobilitat × Càrrega)',
            'C_CASA_noFV_HM.png', hm_pis_vars, hm_pis_noms, hm_pis_fmts,
            invert_cmap={4, 5})

print("\n" + "=" * 60); print("D · Interaccions"); print("=" * 60)

inter_vars = ['Estalvi_V2H_EUR', 'Estalvi_Net_EUR', 'Estalvi_FV_VE_EUR',
              'Cost_Degradacio_EUR', 'Autoconsum_pct', 'PB_Difer_Deg_anys']
inter_noms = ['Estalvi V2H (€)', 'Estalvi net (€)', 'Estalvi FV→VE (€)',
              'Cost degradació (€/any)', 'Autoconsum (%)', 'PB diferencial (anys)']

casa_mod = df_casa[df_casa['Mobilitat'] == 'Moderada']
fig, axes = plt.subplots(2, 3, figsize=(15, 9))
fig.suptitle('CAS CASA — Interacció FV × Càrrega (mobilitat moderada)', fontsize=14)
for idx, (v, n) in enumerate(zip(inter_vars, inter_noms)):
    ax = axes.flat[idx]
    for fv_val, ls, col, lbl in [('Si', 'o-', C_FV, 'Amb FV'),
                                  ('No', 's--', C_CASA, 'Sense FV')]:
        sub = casa_mod[(casa_mod['FV'] == fv_val) & (casa_mod[v] < 900)]
        if len(sub) == 0:
            continue
        means = sub.groupby('Carrega')[v].mean()
        ax.plot(range(len(means)), means.values, ls, label=lbl, lw=2,
                markersize=7, color=col)
    ax.set_xticks(range(len(means)))
    ax.set_xticklabels(means.index, fontsize=9)
    ax.set_title(n, fontsize=10, pad=6)
    ax.legend(fontsize=8, loc='best')
guardar("D_CASA_Interaccio_FVxCarrega.png")

#    Payback principal = AMB degradació (valor real)
print("\n" + "=" * 60); print("E · Payback (amb degradació)"); print("=" * 60)

for cas, df_v, cas_nom in [("CASA", cv, "CASA"), ("PIS", pv, "PIS")]:
    sub = df_v.copy()
    sub['Escenari'] = sub['FV'].astype(str) + ' | ' + sub['Mobilitat'] + ' | ' + sub['Carrega']
    sub = sub.sort_values('PB_Bidir_Deg_anys', ascending=True)

    fig, ax = plt.subplots(figsize=(12, max(5, len(sub) * 0.55)))
    y = np.arange(len(sub))
    h = 0.25

    # Barra base = payback sense degradació (blau clar)
    # Barra overlay = increment per degradació (violeta)
    pb_base = sub['Payback_Bidir_anys'].clip(upper=55).values
    pb_deg = sub['PB_Bidir_Deg_anys'].clip(upper=55).values
    pb_delta = pb_deg - pb_base

    ax.barh(y + h, pb_base, h, color=C_BIDIR, alpha=0.75, label='PB bidir. (base)')
    ax.barh(y + h, pb_delta, h, left=pb_base, color=C_DEG, alpha=0.85,
            label='+ cost degradació')

    ax.barh(y, sub['Payback_Unidir_anys'].clip(upper=55), h,
            label='PB unidir. (1.350 €)', color=C_UNIDIR, alpha=0.85)
    ax.barh(y - h, sub['PB_Difer_Deg_anys'].clip(upper=55), h,
            label='PB diferencial (3.650 €)', color=C_DIFER, alpha=0.85)

    ax.axvline(8, color='red', ls='--', lw=1.5, label=f'Llindar {LLINDAR} anys')
    ax.set_yticks(y)
    ax.set_yticklabels(sub['Escenari'], fontsize=9)
    ax.set_xlabel('Payback (anys)', fontsize=11)
    ax.set_title(f'CAS {cas_nom} — Payback amb cost de degradació '
                 f'(excloent càrrega pública)',
                 fontsize=12, pad=10)
    ax.legend(loc='lower right', fontsize=8, framealpha=0.9)
    pb_max = sub[['PB_Bidir_Deg_anys', 'Payback_Unidir_anys',
                   'PB_Difer_Deg_anys']].clip(upper=55).max().max()
    ax.set_xlim(0, min(55, pb_max * 1.05 + 2))

    # Anotació delta degradació
    for i, (base_v, deg_v) in enumerate(zip(pb_base, pb_deg)):
        if deg_v < 55:
            delta = deg_v - base_v
            ax.text(deg_v + 0.2, y[i] + h, f'+{delta:.1f}',
                    va='center', fontsize=6.5, color=C_DEG, fontweight='bold')

    guardar(f"E_{cas_nom}_Paybacks.png")

# Efecte FV sobre payback bidireccional (amb degradació)
fig, ax = plt.subplots(figsize=(10, 5))
ca = cv[cv['FV'] == 'Si'].sort_values(['Mobilitat', 'Carrega'])
cs = cv[cv['FV'] == 'No'].sort_values(['Mobilitat', 'Carrega'])
mg = cs.merge(ca, on=['Mobilitat', 'Carrega'], suffixes=('_s', '_a'))
mg['Esc'] = mg['Mobilitat'] + '\n' + mg['Carrega']
mg['D'] = mg['PB_Bidir_Deg_anys_a'] - mg['PB_Bidir_Deg_anys_s']
cols_d = ['#5DCA5D' if d < 0 else '#E24B4A' for d in mg['D']]
bars = ax.bar(range(len(mg)), mg['D'], color=cols_d, alpha=0.85,
              edgecolor='white', width=0.6)
ax.set_xticks(range(len(mg)))
ax.set_xticklabels(mg['Esc'], fontsize=9)
ax.set_ylabel('Δ Payback bidir. amb deg. (anys)', fontsize=10)
ax.set_title('CAS CASA — Efecte marginal de la FV sobre el payback bidireccional',
             fontsize=13, pad=10)
ax.axhline(0, color='black', lw=0.5)
for i, v in enumerate(mg['D']):
    ax.text(i, v, f'{v:+.1f}', ha='center',
            va='bottom' if v >= 0 else 'top', fontsize=9)
guardar("E_CASA_Efecte_FV_Payback.png")

# Descomposició: estalvi net vs cost degradació (CASA)
fig, ax = plt.subplots(figsize=(11, max(5, len(cv) * 0.5)))
cv_s = cv.copy()
cv_s['Escenari'] = cv_s['FV'].astype(str) + ' | ' + cv_s['Mobilitat'] + ' | ' + cv_s['Carrega']
cv_s = cv_s.sort_values('Estalvi_Net_EUR', ascending=True)
y = np.arange(len(cv_s))
ax.barh(y, cv_s['Estalvi_Net_EUR'], 0.7, label='Estalvi net V2H (€/any)',
        color=C_BIDIR, alpha=0.85)
ax.barh(y, -cv_s['Cost_Degradacio_EUR'], 0.7, label='Cost degradació (€/any)',
        color=C_DEG, alpha=0.85)
ax.set_yticks(y)
ax.set_yticklabels(cv_s['Escenari'], fontsize=8)
ax.set_xlabel('€/any', fontsize=10)
ax.set_title('CASA — Estalvi net vs cost de degradació',
             fontsize=12, pad=10)
ax.legend(fontsize=9, loc='lower right')
ax.axvline(0, color='black', lw=0.5)
for i, (en, cd) in enumerate(zip(cv_s['Estalvi_Net_EUR'], cv_s['Cost_Degradacio_EUR'])):
    pct = cd / en * 100 if en > 0 else 0
    ax.text(en + 5, i, f'{en:.0f} € (deg: {cd:.1f} €, {pct:.1f}%)',
            va='center', fontsize=7)
guardar("E_CASA_Estalvi_vs_Degradacio.png")

print("\n" + "=" * 60); print("F · Comparació PIS vs CASA"); print("=" * 60)

ref = df[(df['Mobilitat'] == 'Moderada') & (df['Carrega'] == 'Domestica')]
ref_vars = ['Estalvi_V2H_EUR', 'Estalvi_FV_VE_EUR', 'Estalvi_Total_EUR',
            'Cost_Degradacio_EUR', 'Autoconsum_pct', 'EFC_V2H']
ref_noms = ['Estalvi V2H\n(€/any)', 'Estalvi FV→VE\n(€/any)', 'Estalvi total\n(€/any)',
            'Cost degradació\n(€/any)', 'Autoconsum\n(%)',
            'EFC\n(cicles/any)']

fig, axes = plt.subplots(2, 3, figsize=(15, 9))
fig.suptitle('Comparació PIS vs CASA — Escenari referència (moderada, domèstica)',
             fontsize=14)
for idx, (v, n) in enumerate(zip(ref_vars, ref_noms)):
    ax = axes.flat[idx]
    vals = ref.groupby('Habitatge')[v].mean().reindex(['PIS', 'CASA_FV', 'CASA_noFV'])
    colors = [C_PIS, C_FV, C_CASA]
    bars = ax.bar(range(len(vals)), vals.values, color=colors, alpha=0.85,
                  edgecolor='white', width=0.6)
    ax.set_xticks(range(len(vals)))
    ax.set_xticklabels(vals.index, fontsize=9)
    ax.set_ylabel(n, fontsize=9)
    for bar, val in zip(bars, vals.values):
        fmt = '.2f' if v == 'Cost_Degradacio_EUR' else '.1f'
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() * 1.01,
                f'{val:{fmt}}', ha='center', va='bottom', fontsize=9)
guardar("F_Comparacio_PIS_CASA.png")

print("\n" + "=" * 60); print("G · Emissions"); print("=" * 60)

fig, ax = plt.subplots(figsize=(11, max(5, len(dv) * 0.38)))
dfs = dv.sort_values('tCO2_Anual')
dfs['Esc'] = dfs['Habitatge'] + ' | ' + dfs['Mobilitat'] + ' | ' + dfs['Carrega']
cols_e = [C_PIS if h == 'PIS' else (C_FV if fv == 'Si' else C_CASA)
          for h, fv in zip(dfs['Tipus_Habitatge'], dfs['FV'])]
ax.barh(range(len(dfs)), dfs['tCO2_Anual'], color=cols_e, alpha=0.85,
        edgecolor='white', height=0.7)
ax.set_yticks(range(len(dfs)))
ax.set_yticklabels(dfs['Esc'], fontsize=8)
ax.set_xlabel('tCO$_2$ / any', fontsize=10)
ax.set_title('Rànking d\'emissions — Escenaris viables (excloent pública)',
             fontsize=13, pad=10)
ax.legend(handles=[Patch(facecolor=C_PIS, label='PIS'),
                   Patch(facecolor=C_FV, label='CASA amb FV'),
                   Patch(facecolor=C_CASA, label='CASA sense FV')], fontsize=9)
for i, v in enumerate(dfs['tCO2_Anual']):
    ax.text(v + 0.01, i, f'{v:.3f}', va='center', fontsize=7)
guardar("G_Ranking_Emissions.png")

# Efecte FV emissions
fig, ax = plt.subplots(figsize=(10, 5))
mg2 = cs.merge(ca, on=['Mobilitat', 'Carrega'], suffixes=('_s', '_a'))
mg2['D'] = mg2['tCO2_Anual_a'] - mg2['tCO2_Anual_s']
mg2['Esc'] = mg2['Mobilitat'] + '\n' + mg2['Carrega']
ax.bar(range(len(mg2)), mg2['D'],
       color=['#5DCA5D' if d < 0 else '#E24B4A' for d in mg2['D']],
       alpha=0.85, edgecolor='white', width=0.6)
ax.set_xticks(range(len(mg2)))
ax.set_xticklabels(mg2['Esc'], fontsize=9)
ax.set_ylabel('Δ tCO$_2$ (amb FV − sense FV)', fontsize=10)
ax.set_title('CAS CASA — Reducció d\'emissions per FV', fontsize=13, pad=10)
ax.axhline(0, color='black', lw=0.5)
for i, v in enumerate(mg2['D']):
    ax.text(i, v - 0.02, f'{v:.3f}', ha='center', va='top', fontsize=9)
guardar("G_Efecte_FV_Emissions.png")

# Trade-off emissions vs benefici econòmic net, amb frontera de Pareto
# (escenaris no dominats: cap altre escenari amb menys emissions I més benefici)
fig, ax = plt.subplots(figsize=(10.5, 6.5))
marc_c = {'Domestica': 'o', 'Feina': 's'}
pts = []
for _, r in dv.iterrows():
    c = C_PIS if r['Tipus_Habitatge'] == 'PIS' else (
        C_FV if r['FV'] == 'Si' else C_CASA)
    m = marc_c.get(r['Carrega'], 'o')
    estalvi_adj = r['Estalvi_Net_EUR'] - r['Cost_Degradacio_EUR']
    ax.scatter(r['tCO2_Anual'], estalvi_adj, c=c, marker=m, s=95,
               edgecolors='black', linewidths=0.4, alpha=0.85, zorder=3)
    pts.append((r['tCO2_Anual'], estalvi_adj,
                f"{r['Habitatge']} · {r['Mobilitat']} · {r['Carrega']}"))

front = sorted((x, y, l) for x, y, l in pts
               if not any(x2 <= x and y2 >= y and (x2 < x or y2 > y)
                          for x2, y2, _ in pts))
ax.plot([p[0] for p in front], [p[1] for p in front], 'k--', lw=1.3,
        alpha=0.65, zorder=2, label='Frontera de Pareto')
print("  Frontera de Pareto (no dominats):")
for k, (x, y, lbl) in enumerate(front):
    print(f"    {lbl}: {x:.3f} tCO₂, {y:.0f} €/any")
    if k == 0:
        off, ha = (9, -12), 'left'      # punt inferior esquerre: etiqueta a sota-dreta
    elif k == len(front) - 1:
        off, ha = (9, 7), 'left'        # punt superior dret: etiqueta a dalt-dreta
    else:
        off, ha = (-9, 9), 'right'      # punts intermedis: etiqueta a dalt-esquerra
    ax.annotate(lbl, (x, y), textcoords='offset points', xytext=off, ha=ha,
                fontsize=7.5, color='#333333', fontweight='bold')

ax.set_xlabel('Emissions (tCO$_2$/any)', fontsize=10)
ax.set_ylabel('Estalvi net − cost degradació (€/any)', fontsize=10)
ax.set_title('Trade-off: emissions vs benefici econòmic net (incl. degradació)',
             fontsize=13, pad=10)
handles = [Patch(facecolor=C_PIS, label='PIS'),
           Patch(facecolor=C_FV, label='CASA amb FV'),
           Patch(facecolor=C_CASA, label='CASA sense FV'),
           plt.Line2D([0], [0], marker='o', color='grey', linestyle='None',
                      markersize=8, label='Càrrega domèstica'),
           plt.Line2D([0], [0], marker='s', color='grey', linestyle='None',
                      markersize=8, label='Càrrega a feina'),
           plt.Line2D([0], [0], color='black', ls='--', lw=1.3,
                      label='Frontera de Pareto')]
ax.legend(handles=handles, fontsize=8, loc='center right')
guardar("G_Tradeoff_CO2_Estalvi.png")

print("\n" + "=" * 60); print("H · Correlacions"); print("=" * 60)

vars_corr = ['Estalvi_V2H_EUR', 'Estalvi_Net_EUR', 'Estalvi_FV_VE_EUR',
             'Cost_Degradacio_EUR',
             'Autoconsum_pct', 'EFC_V2H', 'V2H_kWh', 'Carg_FV_VE_kWh', 'tCO2_Anual']
noms_c = {
    'Estalvi_V2H_EUR': 'Estalvi V2H (€)',
    'Estalvi_Net_EUR': 'Estalvi net (€)',
    'Estalvi_FV_VE_EUR': 'Estalvi FV→VE (€)',
    'Cost_Degradacio_EUR': 'Cost degradació (€)',
    'Autoconsum_pct': 'Autoconsum (%)',
    'EFC_V2H': 'EFC (cicles/any)',
    'V2H_kWh': 'V2H (kWh)',
    'Carg_FV_VE_kWh': 'FV→VE (kWh)',
    'tCO2_Anual': 'Emissions CO₂ (tCO₂/any)',
}

all_corr_rows = []
for cas, df_cas in [("PIS", df_pis), ("CASA", df_casa)]:
    sub = df_cas[vars_corr].replace(999, np.nan)
    corr = sub.corr(method='spearman')
    for i in range(len(vars_corr)):
        for j in range(i + 1, len(vars_corr)):
            rho = corr.iloc[i, j]
            if not np.isnan(rho) and abs(rho) >= 0.7:
                all_corr_rows.append({
                    'Cas': cas,
                    'Variable 1': noms_c[vars_corr[i]],
                    'Variable 2': noms_c[vars_corr[j]],
                    'rho_Spearman': round(rho, 3),
                    'Abs_rho': round(abs(rho), 3),
                })

df_corr_key = pd.DataFrame(all_corr_rows).sort_values(['Cas', 'Abs_rho'], ascending=[True, False])
df_corr_key.drop(columns='Abs_rho', inplace=True)
df_corr_key.to_csv(CARPETA / "H_Correlacions_Clau.csv", sep=';', decimal=',',
                   index=False, encoding='utf-8-sig')
print("  ✓ H_Correlacions_Clau.csv")
for cas in ['PIS', 'CASA']:
    sub_c = df_corr_key[df_corr_key['Cas'] == cas]
    print(f"  CAS {cas} — correlacions |rho| ≥ 0,70:")
    for _, r in sub_c.iterrows():
        print(f"    {r['Variable 1']}  ↔  {r['Variable 2']}:  rho = {r['rho_Spearman']:.3f}")

print("\n" + "=" * 60); print("I · Pareto"); print("=" * 60)

facs_p = ['FV', 'Mobilitat', 'Carrega']
pb_mean = cv['PB_Bidir_Deg_anys'].mean()  # AMB degradació
eta2_pb = {}
t = anova_factors(cv, facs_p, 'PB_Bidir_Deg_anys')
if t is not None:
    for _, r in t.iterrows():
        if r['Font'] in facs_p:
            eta2_pb[r['Font']] = r['eta2']

eta_sorted = sorted(eta2_pb.items(), key=lambda x: x[1], reverse=True)
fac_n = [f[0] for f in eta_sorted]
fac_v = [f[1] for f in eta_sorted]
fac_col = {'Carrega': '#E24B4A', 'Mobilitat': '#3B8BD4', 'FV': '#EF9F27'}

fig, ax = plt.subplots(figsize=(8, 5.5))
bars = ax.bar(range(len(fac_n)), [v * 100 for v in fac_v],
              color=[fac_col.get(f, 'grey') for f in fac_n],
              alpha=0.85, edgecolor='white', width=0.55)
ax.set_xticks(range(len(fac_n)))
ax.set_xticklabels(fac_n, fontsize=12)
ax.set_ylabel('η² (%)', fontsize=11)
ax.set_ylim(0, 110)
ax.set_title('DIAGRAMA DE PARETO — Variància explicada (PB bidir. amb degradació)',
             fontsize=12, pad=10)

cum = np.cumsum(fac_v)
ax2 = ax.twinx()
ax2.plot(range(len(fac_n)), cum, 'k-o', markersize=7, lw=2)
ax2.set_ylim(0, 1.1)
ax2.set_ylabel('Proporció acumulada', fontsize=11)
ax2.axhline(0.8, color='red', ls='--', alpha=0.6, lw=1.2, label='Llindar 80%')

for i, (v, c) in enumerate(zip(fac_v, cum)):
    ax.text(i, v * 100 + 2, f'{v:.1%}', ha='center', va='bottom',
            fontsize=11, fontweight='bold')

ax2.legend(fontsize=9, loc='center right')
guardar("I_Pareto_Payback.png")

# Efectes per nivell
efectes = []
for fac in facs_p:
    for nivel, grp in cv.groupby(fac)['PB_Bidir_Deg_anys']:
        eff = grp.mean() - pb_mean
        efectes.append({'Factor': fac, 'Nivell': nivel,
                        'Efecte_anys': round(eff, 3),
                        'Efecte_mesos': round(eff * 12, 1)})
df_eff = pd.DataFrame(efectes).sort_values('Efecte_anys', key=abs, ascending=False)
df_eff.to_csv(CARPETA / "T9_Pareto_Efectes.csv", sep=';', decimal=',',
              index=False, encoding='utf-8-sig')
print("  ✓ T9_Pareto_Efectes.csv")

print("\n" + "=" * 60); print("J · Matriu de decisió"); print("=" * 60)

reco_c = {'Bidireccional': C_BIDIR, 'Unidireccional': C_UNIDIR,
           'Unidir. (retorn lent)': '#A8D08D', 'No rendible': '#CCCCCC'}

# Distribució de recomanacions
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.suptitle('Distribució de recomanacions per tipologia (incl. degradació)',
             fontsize=14)
for idx, (cas, ax) in enumerate(zip(['PIS', 'CASA'], axes)):
    sub = df_dec[df_dec['Habitatge'] == cas]
    counts = sub['Recomanacio'].value_counts()
    cols_p = [reco_c.get(r, '#CCCCCC') for r in counts.index]
    wedges, texts, autotexts = ax.pie(
        counts.values, labels=counts.index, autopct='%1.0f%%',
        colors=cols_p, startangle=90, textprops={'fontsize': 9},
        pctdistance=0.75)
    for at in autotexts:
        at.set_fontsize(10)
        at.set_fontweight('bold')
    ax.set_title(f'{cas} ({len(sub)} escenaris)', fontsize=12)
guardar("J_Distribucio_Recomanacions.png")


# Resum degradació integrat
df_deg = dv[['Habitatge', 'FV', 'Mobilitat', 'Carrega',
             'EFC_V2H', 'Cost_Degradacio_EUR', 'Carg_FV_VE_kWh',
             'Estalvi_Net_EUR',
             'PB_Bidir_Deg_anys', 'PB_Difer_Deg_anys',
             'Payback_Bidir_anys', 'Payback_Diferencial_anys']].copy()
df_deg['Delta_PB_Bidir'] = df_deg['PB_Bidir_Deg_anys'] - df_deg['Payback_Bidir_anys']
df_deg['Pct_Deg_vs_Estalvi'] = (df_deg['Cost_Degradacio_EUR'] /
                                 df_deg['Estalvi_Net_EUR'] * 100).round(2)
df_deg.to_csv(CARPETA / "T10_Degradacio_Integrat.csv", sep=';', decimal=',',
              index=False, encoding='utf-8-sig')
print("  ✓ T10_Degradacio_Integrat.csv")

fitxers = sorted(os.listdir(CARPETA))
n_png = sum(1 for f in fitxers if f.endswith('.png'))
n_csv = sum(1 for f in fitxers if f.endswith('.csv'))
print(f"\n{'=' * 60}")
print(f"COMPLETAT: {n_png} PNG + {n_csv} CSV a {CARPETA}/")
print(f"{'=' * 60}")
for f in fitxers:
    print(f"  {f}")
