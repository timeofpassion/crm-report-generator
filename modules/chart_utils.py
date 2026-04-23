import io
import warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as mpatches

warnings.filterwarnings('ignore')

# ── 한국어 폰트 설정 ──────────────────────────────────────────────────────────
def _setup_font():
    candidates = [
        'C:/Windows/Fonts/malgun.ttf',
        'C:/Windows/Fonts/NanumGothic.ttf',
    ]
    for path in candidates:
        try:
            fm.fontManager.addfont(path)
            prop = fm.FontProperties(fname=path)
            matplotlib.rcParams['font.family'] = prop.get_name()
            break
        except Exception:
            continue
    matplotlib.rcParams['axes.unicode_minus'] = False

_setup_font()

# ── 컬러 팔레트 ───────────────────────────────────────────────────────────────
PALETTE = ['#2D7DD2', '#E8A020', '#27AE60', '#E74C3C', '#9B59B6',
           '#1ABC9C', '#E67E22', '#3498DB', '#F39C12', '#2ECC71']

C_PRIMARY   = '#1A3A6B'
C_SECONDARY = '#2D7DD2'
C_ACCENT    = '#E8A020'
C_LIGHT     = '#EBF3FB'
C_GRAY      = '#6B7280'

STYLE = {
    'axes.spines.top':    False,
    'axes.spines.right':  False,
    'axes.facecolor':     '#FAFBFC',
    'figure.facecolor':   'white',
    'grid.color':         '#E5E7EB',
    'grid.linestyle':     '--',
    'grid.alpha':         0.6,
}


def _buf(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf


def _apply_style(ax):
    for k, v in STYLE.items():
        try:
            ax.set(**{k.split('.')[-1]: v})
        except Exception:
            pass
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_facecolor('#FAFBFC')
    ax.yaxis.grid(True, linestyle='--', alpha=0.6, color='#E5E7EB')
    ax.set_axisbelow(True)


def _fmt(v):
    if v >= 100_000_000:
        return f'{v/100_000_000:.1f}억'
    elif v >= 10_000:
        return f'{v/10_000:.0f}만'
    return f'{v:,.0f}'


# ── 도넛 차트 ────────────────────────────────────────────────────────────────
def donut_chart(labels, values, title='', figsize=(6, 5)) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')

    colors = PALETTE[:len(labels)]
    wedges, _, autotexts = ax.pie(
        values, labels=None, autopct='%1.1f%%',
        colors=colors, startangle=90,
        pctdistance=0.82, wedgeprops=dict(width=0.5, edgecolor='white', linewidth=2)
    )
    for at in autotexts:
        at.set_fontsize(8)
        at.set_fontweight('bold')
        at.set_color('white')

    ax.legend(
        wedges, [f'{l} ({_fmt(v)}원)' for l, v in zip(labels, values)],
        loc='center left', bbox_to_anchor=(1, 0.5),
        fontsize=8, frameon=False
    )
    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=10, color=C_PRIMARY)
    return _buf(fig)


# ── 수평 막대 차트 ────────────────────────────────────────────────────────────
def hbar_chart(labels, values, title='', figsize=(7, 5), highlight_top=True) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    y = range(len(labels))
    colors = [C_ACCENT if (highlight_top and i == 0) else C_SECONDARY for i in range(len(labels))]
    bars = ax.barh(list(y), values, color=colors, height=0.6, edgecolor='white')

    ax.set_yticks(list(y))
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    ax.xaxis.set_visible(False)

    for bar, v in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.01, bar.get_y() + bar.get_height() / 2,
                f'{_fmt(v)}원', va='center', ha='left', fontsize=8, color=C_PRIMARY)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 세로 막대 차트 ────────────────────────────────────────────────────────────
def bar_chart(labels, values, title='', color=C_SECONDARY, figsize=(7, 4), label_rotate=30) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    x = range(len(labels))
    ax.bar(list(x), values, color=color, width=0.6, edgecolor='white')
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=label_rotate, ha='right', fontsize=8)
    ax.yaxis.set_visible(False)

    for i, v in enumerate(values):
        ax.text(i, v + max(values) * 0.01, _fmt(v) + '원',
                ha='center', va='bottom', fontsize=7, color=C_PRIMARY)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 그룹 막대 차트 (진료의별) ─────────────────────────────────────────────────
def grouped_bar_chart(df: pd.DataFrame, group_col: str, stack_col: str, value_col: str,
                      title='', figsize=(8, 5)) -> io.BytesIO:
    pivot = df.pivot_table(index=group_col, columns=stack_col, values=value_col,
                           aggfunc='sum', fill_value=0)
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    n = len(pivot.columns)
    x = np.arange(len(pivot.index))
    width = 0.7 / max(n, 1)

    for i, col in enumerate(pivot.columns):
        offset = (i - (n - 1) / 2) * width
        ax.bar(x + offset, pivot[col], width=width * 0.9,
               label=col, color=PALETTE[i % len(PALETTE)], edgecolor='white')

    ax.set_xticks(x)
    ax.set_xticklabels(pivot.index, rotation=20, ha='right', fontsize=8)
    ax.legend(fontsize=7, frameon=False, bbox_to_anchor=(1, 1), loc='upper left')
    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 라인 차트 (일별 추이) ─────────────────────────────────────────────────────
def line_chart(dates, values, title='', figsize=(8, 3.5)) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    ax.plot(dates, values, color=C_SECONDARY, linewidth=2, marker='o', markersize=3)
    ax.fill_between(dates, values, alpha=0.15, color=C_SECONDARY)

    # 주차 구분선
    if len(dates) > 7:
        for i in range(7, len(dates), 7):
            ax.axvline(x=dates[i], color=C_GRAY, linewidth=0.5, linestyle='--', alpha=0.5)

    ax.set_xticks([dates[i] for i in range(0, len(dates), max(1, len(dates) // 6))])
    ax.tick_params(axis='x', rotation=30, labelsize=7)
    ax.yaxis.set_visible(False)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 요일별 막대 (색상 구분) ───────────────────────────────────────────────────
def weekday_bar_chart(labels, values, title='', figsize=(6, 3.5)) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    colors = [C_ACCENT if v == max(values) else C_SECONDARY for v in values]
    ax.bar(labels, values, color=colors, width=0.6, edgecolor='white')
    ax.yaxis.set_visible(False)

    for i, v in enumerate(values):
        ax.text(i, v + max(values) * 0.01, _fmt(v) + '원',
                ha='center', va='bottom', fontsize=7, color=C_PRIMARY)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 분포 히스토그램 ───────────────────────────────────────────────────────────
def histogram(values, title='', figsize=(6, 4)) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')
    _apply_style(ax)

    ax.hist(values, bins=20, color=C_SECONDARY, edgecolor='white', alpha=0.85)
    mean_v = np.mean(values)
    ax.axvline(mean_v, color=C_ACCENT, linewidth=2, linestyle='--',
               label=f'평균: {_fmt(mean_v)}원')
    ax.legend(fontsize=8, frameon=False)
    ax.yaxis.set_visible(False)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.tight_layout()
    return _buf(fig)


# ── 히트맵 ────────────────────────────────────────────────────────────────────
def heatmap(df: pd.DataFrame, title='', figsize=(8, 5)) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_facecolor('white')

    data = df.values.astype(float)
    im = ax.imshow(data, cmap='Blues', aspect='auto')

    ax.set_xticks(range(len(df.columns)))
    ax.set_xticklabels(df.columns, rotation=30, ha='right', fontsize=8)
    ax.set_yticks(range(len(df.index)))
    ax.set_yticklabels(df.index, fontsize=8)

    row_max = data.max(axis=1, keepdims=True)
    for r in range(data.shape[0]):
        for c in range(data.shape[1]):
            v = data[r, c]
            txt_color = 'white' if v > data.max() * 0.6 else C_PRIMARY
            ax.text(c, r, _fmt(v), ha='center', va='center',
                    fontsize=7, color=txt_color)

    if title:
        ax.set_title(title, fontsize=11, fontweight='bold', pad=8, color=C_PRIMARY)
    plt.colorbar(im, ax=ax, shrink=0.6)
    plt.tight_layout()
    return _buf(fig)
