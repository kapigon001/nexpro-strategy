"""matplotlibチャート生成（4枚のみ）"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# フォント設定
FONT_PATH = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
JP_FONT = fm.FontProperties(fname=FONT_PATH)
plt.rcParams['font.family'] = 'IPAGothic'
fm.fontManager.addfont(FONT_PATH)

NAVY = '#1B2A4A'
GOLD = '#C8A951'
RED = '#C0392B'
GREEN = '#27AE60'
BLUE = '#3498DB'
LIGHT_GRAY = '#F2F3F5'

OUT_DIR = '/home/user/nexpro-strategy/assets/charts'
os.makedirs(OUT_DIR, exist_ok=True)


def chart_market():
    """Slide 6: 日本動画配信市場"""
    fig, ax = plt.subplots(figsize=(6, 4))
    years = ['2023', '2024', '2025\n(推定)', '2026\n(推定)', '2027\n(推定)', '2028']
    values = [513, 635, 780, 950, 1200, 1529]
    colors = [NAVY if y in ['2023','2024'] else BLUE for y in ['2023','2024','2025','2026','2027','2028']]
    bars = ax.bar(years, values, color=colors, width=0.6, edgecolor='white')
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 20,
                f'¥{val}B', ha='center', va='bottom', fontsize=10,
                fontproperties=JP_FONT, fontweight='bold')
    ax.set_ylabel('市場規模（十億円）', fontproperties=JP_FONT, fontsize=11)
    ax.set_title('日本法人向け動画配信市場', fontproperties=JP_FONT, fontsize=14, fontweight='bold')
    ax.text(4.5, 1400, 'CAGR 24.4%', fontsize=13, color=RED, fontweight='bold',
            fontproperties=JP_FONT)
    ax.set_ylim(0, 1700)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis='x', labelsize=9)
    for label in ax.get_xticklabels():
        label.set_fontproperties(JP_FONT)
    for label in ax.get_yticklabels():
        label.set_fontproperties(JP_FONT)
    ax.yaxis.label.set_fontproperties(JP_FONT)
    plt.tight_layout()
    fig.savefig(f'{OUT_DIR}/market.png', dpi=150, bbox_inches='tight',
                facecolor='white', transparent=False)
    plt.close()


def chart_revenue():
    """Slide 8: 売上・成長推移コンボチャート"""
    fig, ax1 = plt.subplots(figsize=(8, 4.5))
    years = ['FY22', 'FY23', 'FY24', 'FY25P', 'FY26P', 'FY27P']
    mrr =    [225.9, 252.7, 287.8, 330.0, 417.4, 526.8]
    option = [192.2, 244.8, 225.1, 283.4, 355.7, 477.0]
    compound=[0, 0, 0, 3.2, 17.0, 51.0]
    salesdx= [0, 0, 0, 33.5, 121.8, 327.0]
    x = np.arange(len(years))
    w = 0.55
    ax1.bar(x, mrr, w, label='MRR', color=NAVY)
    ax1.bar(x, option, w, bottom=mrr, label='オプション', color=BLUE)
    bot2 = [m+o for m,o in zip(mrr, option)]
    ax1.bar(x, compound, w, bottom=bot2, label='コンパウンド', color=GREEN)
    bot3 = [b+c for b,c in zip(bot2, compound)]
    ax1.bar(x, salesdx, w, bottom=bot3, label='営業DX', color=GOLD)
    totals = [m+o+c+s for m,o,c,s in zip(mrr, option, compound, salesdx)]
    for i, t in enumerate(totals):
        ax1.text(i, t+15, f'¥{t:.0f}M', ha='center', fontsize=9,
                fontweight='bold', fontproperties=JP_FONT)
    ax1.set_ylabel('売上（百万円）', fontproperties=JP_FONT, fontsize=11)
    ax1.set_ylim(0, 1550)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years, fontproperties=JP_FONT, fontsize=10)
    for label in ax1.get_yticklabels():
        label.set_fontproperties(JP_FONT)
    ax1.yaxis.label.set_fontproperties(JP_FONT)
    ax1.legend(loc='upper left', prop=JP_FONT, fontsize=9)
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    # YoY line
    ax2 = ax1.twinx()
    yoy = [None, 19.0, 3.1, 26.7, 40.3, 51.6]
    yoy_x = [i for i,v in enumerate(yoy) if v is not None]
    yoy_v = [v for v in yoy if v is not None]
    ax2.plot(yoy_x, yoy_v, 'o-', color=RED, linewidth=2, markersize=6)
    for xi, yi in zip(yoy_x, yoy_v):
        c = RED if yi < 10 else NAVY
        ax2.annotate(f'+{yi}%', (xi, yi), textcoords="offset points",
                    xytext=(0, 12), ha='center', fontsize=9, color=c,
                    fontweight='bold', fontproperties=JP_FONT)
    ax2.set_ylabel('YoY成長率（%）', fontproperties=JP_FONT, fontsize=11)
    ax2.set_ylim(0, 70)
    ax2.spines['top'].set_visible(False)
    for label in ax2.get_yticklabels():
        label.set_fontproperties(JP_FONT)
    ax2.yaxis.label.set_fontproperties(JP_FONT)
    plt.title('売上推移と成長率', fontproperties=JP_FONT, fontsize=14, fontweight='bold')
    plt.tight_layout()
    fig.savefig(f'{OUT_DIR}/revenue.png', dpi=150, bbox_inches='tight',
                facecolor='white', transparent=False)
    plt.close()


def chart_positioning_map1():
    """Slide 13: 機能深度 × 日本企業適合性"""
    fig, ax = plt.subplots(figsize=(7, 5.5))
    competitors = {
        'Zoom':        (5.0, 3.0, '#3498DB', 'o'),
        'ON24':        (8.5, 4.0, '#8E44AD', 'o'),
        'EventHub':    (4.5, 7.5, '#1ABC9C', 'o'),
        'bizibl':      (3.0, 8.0, '#E67E22', 'o'),
        'FanGrowth':   (2.5, 8.5, '#E91E63', 'o'),
        'V-CUBE':      (5.5, 7.0, '#7F8C8D', 'o'),
        'Cocripo':     (2.0, 6.5, '#2ECC71', 'o'),
        'Teams/Webex': (4.0, 3.5, '#566573', 'o'),
    }
    for name, (x, y, c, m) in competitors.items():
        ax.scatter(x, y, c=c, s=120, zorder=5, marker=m, edgecolors='white', linewidth=1)
        ax.annotate(name, (x, y), textcoords="offset points",
                   xytext=(8, 5), fontsize=9, fontproperties=JP_FONT, color=c)

    # Nexpro current
    ax.scatter(6.0, 7.0, c=GOLD, s=250, zorder=6, marker='*', edgecolors=NAVY, linewidth=1.5)
    ax.annotate('ネクプロ\n(現在)', (6.0, 7.0), textcoords="offset points",
               xytext=(-40, -20), fontsize=10, fontweight='bold',
               fontproperties=JP_FONT, color=NAVY)
    # Nexpro target
    ax.scatter(8.5, 8.5, c=GOLD, s=300, zorder=6, marker='*',
              edgecolors=RED, linewidth=2, linestyle='--')
    ax.annotate('ネクプロ\n(目標)', (8.5, 8.5), textcoords="offset points",
               xytext=(10, -15), fontsize=10, fontweight='bold',
               fontproperties=JP_FONT, color=RED)
    # Arrow
    ax.annotate('', xy=(8.3, 8.3), xytext=(6.2, 7.2),
               arrowprops=dict(arrowstyle='->', color=GOLD, lw=2.5, ls='--'))

    # Quadrant lines
    ax.axhline(y=5, color='gray', linestyle=':', alpha=0.5)
    ax.axvline(x=5, color='gray', linestyle=':', alpha=0.5)

    # Sweet spot label
    ax.text(7.5, 9.5, '空白地帯', fontsize=14, color=RED, fontweight='bold',
           fontproperties=JP_FONT, ha='center',
           bbox=dict(boxstyle='round,pad=0.3', facecolor='#FADBD8', alpha=0.8))

    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.set_xlabel('機能深度（配信＋分析＋実行）', fontproperties=JP_FONT, fontsize=12)
    ax.set_ylabel('日本企業適合性', fontproperties=JP_FONT, fontsize=12)
    ax.set_title('ポジショニングマップ(1)', fontproperties=JP_FONT, fontsize=14, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(JP_FONT)
    ax.xaxis.label.set_fontproperties(JP_FONT)
    ax.yaxis.label.set_fontproperties(JP_FONT)
    plt.tight_layout()
    fig.savefig(f'{OUT_DIR}/pos_map1.png', dpi=150, bbox_inches='tight',
                facecolor='white', transparent=False)
    plt.close()


def chart_positioning_map2():
    """Slide 14: データ活用高度性 × 導入容易性"""
    fig, ax = plt.subplots(figsize=(7, 5.5))
    competitors = {
        'Zoom':        (3.0, 3.5, '#3498DB'),
        'ON24':        (8.5, 3.0, '#8E44AD'),
        'EventHub':    (3.5, 6.5, '#1ABC9C'),
        'bizibl':      (2.0, 8.0, '#E67E22'),
        'FanGrowth':   (2.5, 8.5, '#E91E63'),
        'V-CUBE':      (3.0, 4.0, '#7F8C8D'),
        'Cocripo':     (1.5, 9.0, '#2ECC71'),
        'Teams/Webex': (2.5, 5.0, '#566573'),
    }
    # Sweet spot zone
    rect = plt.Rectangle((5, 5), 5, 5, linewidth=0, edgecolor=None,
                         facecolor='#D5F5E3', alpha=0.4, zorder=1)
    ax.add_patch(rect)
    ax.text(7.5, 9.3, 'Sweet Spot', fontsize=13, color=GREEN, fontweight='bold',
           fontproperties=JP_FONT, ha='center')

    for name, (x, y, c) in competitors.items():
        ax.scatter(x, y, c=c, s=120, zorder=5, marker='o', edgecolors='white', linewidth=1)
        ax.annotate(name, (x, y), textcoords="offset points",
                   xytext=(8, 5), fontsize=9, fontproperties=JP_FONT, color=c)

    ax.scatter(4.5, 6.0, c=GOLD, s=250, zorder=6, marker='*', edgecolors=NAVY, linewidth=1.5)
    ax.annotate('ネクプロ\n(現在)', (4.5, 6.0), textcoords="offset points",
               xytext=(-40, -20), fontsize=10, fontweight='bold',
               fontproperties=JP_FONT, color=NAVY)
    ax.scatter(7.5, 7.0, c=GOLD, s=300, zorder=6, marker='*',
              edgecolors=RED, linewidth=2)
    ax.annotate('ネクプロ\n(目標)', (7.5, 7.0), textcoords="offset points",
               xytext=(10, -15), fontsize=10, fontweight='bold',
               fontproperties=JP_FONT, color=RED)
    ax.annotate('', xy=(7.3, 6.9), xytext=(4.7, 6.1),
               arrowprops=dict(arrowstyle='->', color=GOLD, lw=2.5, ls='--'))

    ax.axhline(y=5, color='gray', linestyle=':', alpha=0.5)
    ax.axvline(x=5, color='gray', linestyle=':', alpha=0.5)

    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.set_xlabel('データ活用高度性', fontproperties=JP_FONT, fontsize=12)
    ax.set_ylabel('導入容易性（上=容易）', fontproperties=JP_FONT, fontsize=12)
    ax.set_title('ポジショニングマップ(2)', fontproperties=JP_FONT, fontsize=14, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_fontproperties(JP_FONT)
    ax.xaxis.label.set_fontproperties(JP_FONT)
    ax.yaxis.label.set_fontproperties(JP_FONT)
    plt.tight_layout()
    fig.savefig(f'{OUT_DIR}/pos_map2.png', dpi=150, bbox_inches='tight',
                facecolor='white', transparent=False)
    plt.close()


def generate_all_charts():
    print("  チャート1: 市場規模...")
    chart_market()
    print("  チャート2: 売上推移...")
    chart_revenue()
    print("  チャート3: ポジショニングマップ1...")
    chart_positioning_map1()
    print("  チャート4: ポジショニングマップ2...")
    chart_positioning_map2()
    print("  全チャート生成完了")


if __name__ == '__main__':
    generate_all_charts()
