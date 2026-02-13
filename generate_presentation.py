"""ネクプロ 全社戦略プレゼンテーション — 配信×データ×AI 収束版"""
import os
import sys
sys.path.insert(0, '/home/user/nexpro-strategy')

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from slide_helpers import *
from generate_charts import generate_all_charts

CHART_DIR = '/home/user/nexpro-strategy/assets/charts'
OUT_DIR = '/home/user/nexpro-strategy/output'
os.makedirs(OUT_DIR, exist_ok=True)


# =====================================================
# Ch1: 結論 — 3つの掛け算
# =====================================================

def slide_01_title(prs):
    slide = add_slide(prs)
    set_bg(slide, NAVY)
    add_textbox(slide, Inches(1), Inches(1.8), Inches(11), Inches(1.2),
                "ネクプロ 全社戦略提案", font_size=40, bold=True, color=WHITE,
                alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(3.2), Inches(11), Inches(1.2),
                "配信 × データ × AI",
                font_size=32, bold=True, color=GOLD, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(4.6), Inches(11), Inches(0.6),
                "ウェビナーの先にある¥1兆市場を、最初に取りに行く",
                font_size=18, color=WHITE, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(6.0), Inches(11), Inches(0.5),
                "2026年2月 | CEO + 全マネージャーMTG | Confidential",
                font_size=12, color=MED_GRAY, alignment=PP_ALIGN.CENTER)
    add_notes(slide, "配信は強み。その上にデータとAIを掛け合わせることで¥200B→¥2T市場のプレイヤーになれる。本日はこの戦略の承認と実行体制の決定を行います。")


def slide_02_multiplication(prs):
    """Ch1: 3つの掛け算 — 配信×データ×AI"""
    slide = add_slide(prs)
    add_header_bar(slide, "結論：3つの掛け算",
                   "配信 × データ × AI ＝ ¥200B市場から¥2T市場へのプレイヤーになる切符")
    # 3 multiplication boxes
    layers = [
        ("配信（強み）", "ウェビナー配信・擬似ライブ\nメディアサイト機能",
         "高品質な1stパーティ\nデータの生成エンジン", NAVY),
        ("× データ", "視聴ログの蓄積\n（現在は未活用）",
         "エンゲージメントスコア\n→ 商談化予測 → SF連携", BLUE),
        ("× AI", "未実装",
         "日本語コンテンツ自動生成\nインテント分析・レコメンド", GREEN),
    ]
    for i, (label, now, future, color) in enumerate(layers):
        y = Inches(2.0 + i * 1.7)
        # Label
        lbl = add_rounded_rect(slide, Inches(0.5), y, Inches(2.2), Inches(1.3), color)
        set_shape_text(lbl, label, font_size=18, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
        # Now
        now_box = add_rounded_rect(slide, Inches(3.0), y, Inches(4.5), Inches(1.3),
                                    LIGHT_GRAY, color)
        set_shape_multiline(now_box, ["現在の価値", now], font_size=11, color=DARK_TEXT)
        now_box.text_frame.paragraphs[0].font.bold = True
        now_box.text_frame.paragraphs[0].font.size = Pt(9)
        now_box.text_frame.paragraphs[0].font.color.rgb = MED_GRAY
        # Arrow
        add_textbox(slide, Inches(7.7), y + Inches(0.3), Inches(0.6), Inches(0.6),
                    "→", font_size=24, bold=True, color=GOLD, alignment=PP_ALIGN.CENTER)
        # Future
        fut_box = add_rounded_rect(slide, Inches(8.5), y, Inches(4.5), Inches(1.3),
                                    color)
        set_shape_multiline(fut_box, ["掛け算後の価値", future], font_size=11, color=WHITE)
        fut_box.text_frame.paragraphs[0].font.bold = True
        fut_box.text_frame.paragraphs[0].font.size = Pt(9)
    add_footer(slide, 2)
    add_notes(slide, "配信力は強みであり守るべき資産。問題は配信「だけ」に留まること。データとAIを掛けることで到達市場が桁違いに広がる。")


def slide_03_kpi_target(prs):
    """Ch1: KPI — 現在 → 36ヶ月後"""
    slide = add_slide(prs)
    add_header_bar(slide, "目標KPI",
                   "掛け算の成果を4つのKPIで測る")
    data = [
        ["指標", "FY24実績", "FY27目標", "変化"],
        ["売上", "¥512M", "¥1,382M", "2.7倍"],
        ["ARPA", "¥148K/月", "¥204K/月", "+38%"],
        ["月次解約率", "1.7%", "1.0%", "-0.7pt"],
        ["新収益比率", "0%", "27%", "配信の上に積む新収益"],
    ]
    add_table(slide, 5, 4, data, Inches(1.5), Inches(2.2), Inches(10.3), Inches(3.0))
    # Message
    add_textbox(slide, Inches(1.5), Inches(5.8), Inches(10.3), Inches(0.8),
                "FY24 +3.1%の成長鈍化 → 掛け算なしでは¥1B突破は構造的に困難\n"
                "新収益（コンパウンド¥51M + 営業DX¥327M）がFY27で27%を占める計画",
                font_size=12, color=DARK_TEXT)
    add_footer(slide, 3)
    add_notes(slide, "売上2.7倍は野心的だが、Gate Reviewで段階検証する。新収益27%は配信の上にデータ×AIを載せた結果。")


# =====================================================
# Ch2: なぜ今か — 窓が開いている
# =====================================================

def slide_04_window(prs):
    """Ch2: 3つの追い風"""
    slide = add_slide(prs)
    add_header_bar(slide, "なぜ今か：窓が開いている",
                   "12-18ヶ月の機会ウィンドウ。閉まる前に動く")
    winds = [
        ("1", "競合の隙間", NAVY,
         ["ON24 → Cvent買収で日本優先度低下",
          "Zoom/Teams → AI搭載するが日本語適合が構造的に弱い",
          "国内勢 → データ活用が浅い",
          "「日本×高機能×データ」を埋めている企業がいない"]),
        ("2", "データの価値シフト", BLUE,
         ["3rdパーティCookie廃止 → 1stパーティデータ価値急増",
          "ウェビナー = 行動・関心・熱量を直接取得できる稀有なチャネル",
          "配信していること自体がデータ資産になる"]),
        ("3", "市場の構造的拡大", GREEN,
         ["企業あたりウェビナー回数: 13→47回/年（3.6倍）",
          "法人向け動画配信市場: CAGR 24.4%",
          "CRM/MA市場（¥4,190B）との接続で到達市場が桁違いに"]),
    ]
    for i, (num, title, color, items) in enumerate(winds):
        x = Inches(0.3 + i * 4.3)
        # Header
        hdr = add_rounded_rect(slide, x, Inches(2.0), Inches(4.0), Inches(0.7), color)
        set_shape_text(hdr, f"  {num}  {title}", font_size=16, bold=True, color=WHITE)
        # Items
        add_multiline_textbox(slide, x + Inches(0.1), Inches(2.9), Inches(3.8), Inches(3.8),
                               items, font_size=10, bullet=True, color=DARK_TEXT,
                               line_spacing=1.4)
    add_footer(slide, 4)
    add_notes(slide, "3つの追い風が同時に吹いている。日本SaaS浸透率4%、AI統合12-18ヶ月遅れ。この猶予は永続しない。SF Agentforce対応国内最先行の優位を活かす窓が今。")


def slide_05_market(prs):
    """Ch2: 市場チャート"""
    slide = add_slide(prs)
    add_header_bar(slide, "日本法人向け動画配信市場",
                   "CAGR 24.4% — 配信市場は拡大中。さらにデータ接続でCRM/MA市場（¥4T+）に到達")
    slide.shapes.add_picture(f'{CHART_DIR}/market.png',
                             Inches(0.3), Inches(1.9), Inches(6.5), Inches(5.0))
    # 市場拡張の図
    markets = [
        ("配信単体", "¥635B", "動画配信市場", NAVY),
        ("+ データ接続", "¥4,190B", "+ CRM/MA市場", BLUE),
        ("+ AI活用", "¥2T+", "+ デジタルマーケ全体", GREEN),
    ]
    for i, (label, size, desc, color) in enumerate(markets):
        y = Inches(2.3 + i * 1.5)
        box = add_rounded_rect(slide, Inches(7.2), y, Inches(5.5), Inches(1.2), color)
        set_shape_multiline(box, [f"  {label}", f"  到達可能市場: {size}", f"  {desc}"],
                           font_size=11, color=WHITE)
        box.text_frame.paragraphs[0].font.bold = True
        box.text_frame.paragraphs[0].font.size = Pt(14)
        box.text_frame.paragraphs[0].font.color.rgb = GOLD
    add_footer(slide, 5)
    add_notes(slide, "配信市場自体も成長中だが、データ・AIを掛けることで到達可能市場が桁違いに広がる。これが「掛け算」の意味。")


# =====================================================
# Ch3: 勝てる理由 — 空白ポジション
# =====================================================

def slide_06_pos_map1(prs):
    """Ch3: ポジショニングマップ1"""
    slide = add_slide(prs)
    add_header_bar(slide, "勝てる理由：空白ポジション",
                   "「高機能×日本適合×データ活用」の交差点は空白。ネクプロだけが到達可能")
    slide.shapes.add_picture(f'{CHART_DIR}/pos_map1.png',
                             Inches(0.3), Inches(1.8), Inches(7.5), Inches(5.5))
    # Key insight
    add_textbox(slide, Inches(8.0), Inches(2.2), Inches(5.0), Inches(2.0),
                "ON24レベルの機能深度\n× 国内勢レベルの日本適合性\n\nこの交差点に到達できるのは\nネクプロだけ",
                font_size=14, bold=True, color=NAVY, line_spacing=1.5)
    # Competitor summary
    mini = [
        ["", "配信", "CRM連携", "日本適合", "データ"],
        ["ネクプロ", "◎", "◎", "◎", "○→◎"],
        ["Zoom", "◎", "○", "△", "△"],
        ["ON24", "○", "◎", "△", "◎"],
        ["国内勢", "○-△", "△", "◎", "△"],
    ]
    add_table(slide, 5, 5, mini, Inches(8.0), Inches(4.5), Inches(5.0), Inches(2.5))
    add_footer(slide, 6)
    add_notes(slide, "右上の空白地帯が目標ポジション。ON24は機能は深いが日本適合が弱い。国内勢は逆。ネクプロは両方を持っている。")


def slide_07_pos_map2(prs):
    """Ch3: ポジショニングマップ2"""
    slide = add_slide(prs)
    add_header_bar(slide, "ポジショニングマップ(2): データ活用 × 導入容易性",
                   "「高データ活用×低障壁」のSweet Spotを狙う")
    slide.shapes.add_picture(f'{CHART_DIR}/pos_map2.png',
                             Inches(1.5), Inches(1.8), Inches(10.3), Inches(5.5))
    add_footer(slide, 7)
    add_notes(slide, "ON24はデータ活用が高いが導入が重い。ネクプロはテンプレ・伴走支援で障壁を下げつつデータ活用を高度化する。")


# =====================================================
# Ch4: 掛け算の実装 — 3つのギア
# =====================================================

def slide_08_three_gears_overview(prs):
    """Ch4: 3ギア全体像"""
    slide = add_slide(prs)
    add_header_bar(slide, "掛け算の実装：3つのギア",
                   "配信を深め（Gear1）、データを載せ（Gear2）、AIで回す（Gear3）")
    gears = [
        ("Gear 1", "配信の深化", "0-6M", GREEN,
         ["業種別PKG×4", "オンボーディング標準化", "CS→プロフィット化", "ヘルススコア導入"],
         "解約率1.3%以下\nARPA ¥160K以上"),
        ("Gear 2", "データレイヤー", "0-12M", BLUE,
         ["スコアMVP開発", "SF連携強化", "価格3層化", "商談化接続"],
         "スコアMVP 5社\nSF連携率80%"),
        ("Gear 3", "AIレイヤー", "6-18M", NAVY,
         ["AIコンテンツ生成", "インテント分析", "API-first化", "HubSpot/Marketo連携"],
         "新収益比率15%超\n3層稼働"),
    ]
    for i, (gear, title, period, color, items, goal) in enumerate(gears):
        x = Inches(0.3 + i * 4.3)
        # Gear header
        hdr = add_rounded_rect(slide, x, Inches(2.0), Inches(4.0), Inches(0.9), color)
        set_shape_multiline(hdr, [f"{gear}: {title}", f"期間: {period}"],
                           font_size=14, color=WHITE, bold=True,
                           alignment=PP_ALIGN.CENTER)
        hdr.text_frame.paragraphs[1].font.size = Pt(10)
        hdr.text_frame.paragraphs[1].font.bold = False
        # Items
        add_multiline_textbox(slide, x + Inches(0.1), Inches(3.1), Inches(3.8), Inches(2.0),
                               items, font_size=11, bullet=True, color=DARK_TEXT)
        # Goal
        goal_box = add_rounded_rect(slide, x, Inches(5.3), Inches(4.0), Inches(0.9),
                                     LIGHT_GRAY, color)
        set_shape_multiline(goal_box, ["ゴール:", goal], font_size=10, color=DARK_TEXT)
        goal_box.text_frame.paragraphs[0].font.bold = True
        goal_box.text_frame.paragraphs[0].font.color.rgb = color

    # Timeline bar
    add_rect(slide, Inches(0.3), Inches(6.5), Inches(12.7), Inches(0.06), MED_GRAY)
    marks = [("M0", Inches(0.3)), ("M6\nGate1", Inches(4.5)),
             ("M12\nGate2", Inches(8.5)), ("M18", Inches(12.0))]
    for label, x in marks:
        add_textbox(slide, x, Inches(6.6), Inches(1.2), Inches(0.6),
                    label, font_size=8, bold=True, color=NAVY, alignment=PP_ALIGN.CENTER)
    add_footer(slide, 8)
    add_notes(slide, "3つのギアは順次実装。Gear1は既存業務の型化で即着手可能。Gear2がこの戦略の核心。Gear3はGear2の成果を前提に展開。")


def slide_09_gear1(prs):
    """Ch4: Gear1 詳細"""
    slide = add_slide(prs)
    add_header_bar(slide, "Gear 1: 配信の深化（0-6ヶ月）",
                   "強みをさらに強くする — 止血と基盤固め")
    data = [
        ["施策", "内容", "成果KPI", "担当"],
        ["業種別パッケージ×4", "IT/製造/金融/医療のテンプレ+運用ガイド", "成約率+5pt", "営業"],
        ["オンボーディング標準化", "60日内完了率90%を目標に型化", "解約率改善の基盤", "CS"],
        ["CS→プロフィットセンター化", "伴走支援を収益化", "CS起点年¥30M", "CS"],
        ["ヘルススコア導入", "利用データで解約予兆を検知", "解約予測精度70%", "CS/Prod"],
    ]
    add_table(slide, 5, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(3.0))
    # Goal box
    goal = add_rounded_rect(slide, Inches(3.0), Inches(5.5), Inches(7.3), Inches(1.2), GREEN)
    set_shape_multiline(goal,
        ["Gear 1 ゴール（Gate 1基準）",
         "月次解約率: 1.7% → 1.3%以下 / ARPA: ¥148K → ¥160K以上"],
        font_size=14, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    goal.text_frame.paragraphs[1].font.size = Pt(12)
    goal.text_frame.paragraphs[1].font.bold = False
    add_footer(slide, 9)
    add_notes(slide, "Gear1は既存業務の型化が中心。追加投資は最小限。業種別PKGとオンボーディング標準化は即着手可能。")


def slide_10_gear2(prs):
    """Ch4: Gear2 詳細"""
    slide = add_slide(prs)
    add_header_bar(slide, "Gear 2: データレイヤーの構築（0-12ヶ月）",
                   "配信の上にデータを載せる — この戦略の核心")
    data = [
        ["施策", "内容", "成果KPI", "担当"],
        ["エンゲージメントスコアMVP", "視聴行動→関心度スコア→SF連携", "5社導入、ARPA+20%", "Prod"],
        ["Salesforce連携強化", "Agentforce深化、双方向データ同期", "連携率80%", "Prod"],
        ["価格体系3層化", "Base（配信）+ Data（分析）+ AI", "ARPA+15%", "営業"],
        ["商談化接続強化", "スコア→商談化率の因果接続", "商談化寄与額の可視化", "営業/CS"],
    ]
    add_table(slide, 5, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(3.0))
    goal = add_rounded_rect(slide, Inches(3.0), Inches(5.5), Inches(7.3), Inches(1.2), BLUE)
    set_shape_multiline(goal,
        ["Gear 2 ゴール（Gate 2基準）",
         "スコアMVP 5社導入 / SF連携率80% / 「配信×データ」の価値実証"],
        font_size=14, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    goal.text_frame.paragraphs[1].font.size = Pt(12)
    goal.text_frame.paragraphs[1].font.bold = False
    add_footer(slide, 10)
    add_notes(slide, "エンゲージメントスコアMVPが最重要施策。視聴行動データを関心度スコアに変換し、SFの商談データと接続する。これが「配信×データ」の実体。")


def slide_11_gear3(prs):
    """Ch4: Gear3 詳細"""
    slide = add_slide(prs)
    add_header_bar(slide, "Gear 3: AIレイヤーの実装（6-18ヶ月）",
                   "データの上にAIを載せる — 市場拡張フェーズ")
    data = [
        ["施策", "内容", "成果KPI", "担当"],
        ["AI日本語コンテンツ自動生成", "ウェビナー→ブログ/メール/SNS/FAQ", "採用率30%", "Prod"],
        ["インテント分析", "行動データからの購買意欲推定", "データpt 40+/人", "Prod"],
        ["API-first化", "外部連携の基盤整備", "API網羅率80%", "Prod"],
        ["HubSpot/Marketo連携", "SF以外のMA/CRM展開", "新規20社/年", "Prod/営業"],
    ]
    add_table(slide, 5, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(3.0))
    goal = add_rounded_rect(slide, Inches(3.0), Inches(5.5), Inches(7.3), Inches(1.2), NAVY)
    set_shape_multiline(goal,
        ["Gear 3 ゴール（Gate 3基準）",
         "新収益比率15%超 / 「配信×データ×AI」の3層が稼働"],
        font_size=14, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    goal.text_frame.paragraphs[1].font.size = Pt(12)
    goal.text_frame.paragraphs[1].font.bold = False
    add_footer(slide, 11)
    add_notes(slide, "Gear3はGear2の成果を前提に展開。ON24のAI Propel+が先例：AIコンテンツで従来比7倍のエンゲージメント。Cventが¥400億で買収した価値がここにある。")


def slide_12_revenue(prs):
    """Ch4: 売上推移チャート"""
    slide = add_slide(prs)
    add_header_bar(slide, "売上推移と成長率",
                   "FY24 +3.1%成長鈍化 → 掛け算で新収益柱を積む")
    slide.shapes.add_picture(f'{CHART_DIR}/revenue.png',
                             Inches(0.5), Inches(1.9), Inches(12.3), Inches(5.2))
    add_footer(slide, 12)
    add_notes(slide, "FY24の+3.1%成長を正面から見せる。MRR+オプションだけでは横ばい。コンパウンド+営業DXの新収益柱がFY27で27%を占める計画。")


# =====================================================
# Ch5: いつ・誰が・どう測る
# =====================================================

def slide_13_gate_review(prs):
    """Ch5: Gate Review基準"""
    slide = add_slide(prs)
    add_header_bar(slide, "Gate Reviewで段階検証",
                   "仮説ではなくファクトでギアを上げる")
    data = [
        ["Gate", "時期", "判定KPI", "基準", "達成時", "未達時"],
        ["Gate 1", "M6", "スコアMVP導入社数", "5社以上", "Gear3本格化", "延長・修正"],
        ["", "", "月次解約率", "≤1.3%", "投資継続", "CS施策見直し"],
        ["", "", "ARPA（長期）", "≥¥160K", "投資継続", "価格再検討"],
        ["Gate 2", "M12", "新収益比率", "15%超", "組織拡大", "ピボット検討"],
        ["", "", "営業DX売上", "≥¥15M/半期", "投資拡大", "縮小検討"],
        ["Gate 3", "M24", "売上", "¥1Bペース", "成長加速", "戦略再設計"],
    ]
    add_table(slide, 7, 6, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(3.8))
    # Message
    add_textbox(slide, Inches(0.5), Inches(6.2), Inches(12.0), Inches(0.6),
                "Gate Reviewは「計画通りか」ではなく「投資を続けるべきか」をファクトベースに判断する場",
                font_size=13, bold=True, color=RED)
    add_footer(slide, 13)
    add_notes(slide, "全力投球ではなく、実績で信頼を獲得しながらギアを上げる。未達時のピボットルールを予め決めておくことが重要。")


def slide_14_org_governance(prs):
    """Ch5: 組織 + ガバナンス"""
    slide = add_slide(prs)
    add_header_bar(slide, "組織とガバナンス",
                   "兼務→成果確認→専任化 / KPIオーナー制で実行を担保")
    # Org timeline
    org_data = [
        ["時期", "アクション", "備考"],
        ["M1", "KPIオーナー全員アサイン", "各KPIに個人名を紐付け"],
        ["M1-3", "PMM機能（兼務）スタート", "競合レポート月次化"],
        ["M3-6", "RevOps機能（兼務）", "ファネル可視化100%"],
        ["M6-", "成果確認→専任化（採用+3-5名）", "Gate 1結果で判断"],
    ]
    add_table(slide, 5, 3, org_data, Inches(0.3), Inches(2.0), Inches(7.5), Inches(2.8))
    # Governance
    gov_data = [
        ["会議体", "頻度", "目的"],
        ["経営KPIレビュー", "月次", "KPI進捗と阻害要因の特定"],
        ["戦略レビュー", "四半期", "ギアの進捗と方向性確認"],
        ["Gate Review", "M6/M12/M24", "投資継続・拡大の判断"],
    ]
    add_table(slide, 4, 3, gov_data, Inches(0.3), Inches(5.2), Inches(7.5), Inches(2.0),
              header_color=GREEN)
    # Org chart
    boxes = [
        (Inches(9.5), Inches(2.5), "CEO"),
        (Inches(8.2), Inches(3.8), "PMM\n(兼務)"),
        (Inches(9.8), Inches(3.8), "RevOps\n(兼務)"),
        (Inches(11.4), Inches(3.8), "Prod"),
        (Inches(8.2), Inches(5.2), "Sales"),
        (Inches(9.8), Inches(5.2), "CS"),
        (Inches(11.4), Inches(5.2), "AI/Data\n(採用)"),
    ]
    for x, y, text in boxes:
        box = add_rounded_rect(slide, x, y, Inches(1.4), Inches(0.8), NAVY)
        set_shape_text(box, text, font_size=9, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
    add_footer(slide, 14)
    add_notes(slide, "30名体制では全てを専任化できない。兼務→成果確認→専任化の段階アプローチ。KPIオーナー制度が実行の核心。")


# =====================================================
# Ch6: 決めること
# =====================================================

def slide_15_decisions(prs):
    """Ch6: 3つの決議 + Next Steps"""
    slide = add_slide(prs)
    add_header_bar(slide, "本日決めること",
                   "今日決めなければ「窓が閉まるのを見ているだけ」になる")
    # 3 decisions
    decisions = [
        ("1", "掛け算戦略の承認", "配信×データ×AIの3ギア実装を承認するか"),
        ("2", "投資配分の承認", "プロダクト45% / GTM35% / 組織20%"),
        ("3", "組織再編の承認", "PMM/RevOps兼務新設 + KPIオーナー制 + Gate Review設定"),
    ]
    for i, (num, title, desc) in enumerate(decisions):
        y = Inches(2.0 + i * 1.2)
        add_rect(slide, Inches(0.5), y, Inches(0.08), Inches(1.0), GOLD)
        box = add_rect(slide, Inches(0.58), y, Inches(6.0), Inches(1.0), DARK_NAVY)
        add_textbox(slide, Inches(0.8), y + Inches(0.05), Inches(0.6), Inches(0.6),
                    num, font_size=28, bold=True, color=GOLD)
        add_textbox(slide, Inches(1.5), y + Inches(0.05), Inches(4.8), Inches(0.4),
                    title, font_size=16, bold=True, color=WHITE)
        add_textbox(slide, Inches(1.5), y + Inches(0.5), Inches(4.8), Inches(0.4),
                    desc, font_size=10, color=MED_GRAY)
    # Next Steps
    ns = [
        ["アクション", "担当", "期限"],
        ["KPIオーナーアサイン", "CEO", "1週間以内"],
        ["PMM兼務者決定", "マネージャー会議", "2週間以内"],
        ["スコアMVP要件定義", "プロダクト", "1ヶ月以内"],
        ["業種別PKG設計", "営業", "1ヶ月以内"],
        ["Gate Review 1 日程確定", "CEO", "1週間以内"],
    ]
    add_table(slide, 6, 3, ns, Inches(7.0), Inches(2.0), Inches(5.8), Inches(4.2),
              header_color=GREEN)
    add_footer(slide, 15)
    add_notes(slide, "沈黙を恐れず一つずつ確認。「保留」は実質的に「窓が閉まるのを見ている」と同義。")


def slide_16_closing(prs):
    slide = add_slide(prs)
    set_bg(slide, NAVY)
    add_textbox(slide, Inches(1), Inches(2.0), Inches(11), Inches(1),
                "配信 × データ × AI", font_size=36, bold=True, color=GOLD,
                alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(3.5), Inches(11), Inches(0.8),
                "ウェビナーの先にある¥1兆市場を、最初に取りに行く",
                font_size=20, color=WHITE, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.5),
                "Q&A",
                font_size=24, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    add_notes(slide, "想定Q&A(10問)を準備済み。厳しい質問を歓迎する姿勢を見せる。")


# =====================================================
# Appendix
# =====================================================

def slide_app_revenue(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 売上内訳詳細")
    data = [
        ["", "FY22", "FY23", "FY24", "FY25P", "FY26P", "FY27P"],
        ["売上合計", "¥418M", "¥497M", "¥512M", "¥649M", "¥911M", "¥1,382M"],
        ["YoY", "-", "+19.0%", "+3.1%", "+26.7%", "+40.3%", "+51.6%"],
        ["MRR年額", "¥225M", "¥252M", "¥287M", "¥330M", "¥417M", "¥526M"],
        ["オプション", "¥192M", "¥244M", "¥225M", "¥283M", "¥355M", "¥477M"],
        ["コンパウンド", "-", "-", "-", "¥3M", "¥17M", "¥51M"],
        ["営業DX", "-", "-", "-", "¥33M", "¥121M", "¥327M"],
    ]
    add_table(slide, 7, 7, data, Inches(0.3), Inches(1.8), Inches(12.7), Inches(4.5))
    add_footer(slide, 17)


def slide_app_kpi(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: SaaS KPI推移")
    data = [
        ["指標", "FY22", "FY23", "FY24", "目標"],
        ["長期PFアカウント数", "160", "151", "167", "210"],
        ["ARPA長期(千円/月)", "-", "¥137.6K", "¥148.0K", "¥204K"],
        ["月次解約率(長期)", "3.6%", "2.3%", "1.7%", "1.0%"],
        ["新規長期PF成約/年", "60", "27", "50", "38"],
        ["成約率", "11.2%", "9.5%", "10.5%", "-"],
    ]
    add_table(slide, 6, 5, data, Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.5),
              highlight_rows=[3], highlight_color=RGBColor(0xFA, 0xDB, 0xD8))
    add_textbox(slide, Inches(0.5), Inches(5.8), Inches(12.0), Inches(0.8),
                "SaaS優良企業の目安: 月次解約率 < 0.42%（年5%以下）\n"
                "ネクプロは改善トレンドにあるが、Gear1の施策で1.0%到達を目指す",
                font_size=12, color=DARK_TEXT)
    add_footer(slide, 18)


def slide_app_swot(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: SWOT分析")
    qdata = {
        "S（配信力が強み）": (Inches(0.3), Inches(2.0), Inches(6.2), Inches(2.5), NAVY, WHITE, [
            "1. ITreview 13期連続受賞",
            "2. SF Agentforce国内最先行",
            "3. 400-500社エンタープライズ基盤",
            "4. 日本語運用ノウハウ・伴走支援",
            "5. 擬似ライブ・メディアサイト機能",
        ]),
        "W（規模とリソース）": (Inches(6.8), Inches(2.0), Inches(6.2), Inches(2.5), GOLD, DARK_TEXT, [
            "1. 30名の少数精鋭体制",
            "2. 総調達額¥7.5億",
            "3. ブランド認知度劣後",
            "4. インテントデータ未整備",
            "5. PMM/RevOps機能の不在",
        ]),
        "O（窓が開いている）": (Inches(0.3), Inches(4.7), Inches(6.2), Inches(2.5), GREEN, WHITE, [
            "1. AI統合12-18ヶ月遅れ",
            "2. ON24→Cvent買収で日本優先度低下",
            "3. Cookie廃止→1stパーティ価値増",
            "4. SaaS浸透率4%（DX余地大）",
            "5. ウェビナー数3.6倍増(13→47回/年)",
        ]),
        "T（窓は閉まる）": (Inches(6.8), Inches(4.7), Inches(6.2), Inches(2.5), RED, WHITE, [
            "1. Zoom/Teams AI搭載",
            "2. 配信コモディティ化",
            "3. Cvent大型統合(ON24+Goldcast)",
            "4. 価格競争激化",
            "5. 汎用AIで差別化希薄化",
        ]),
    }
    for label, (x, y, w, h, bg, tc, items) in qdata.items():
        shape = add_rounded_rect(slide, x, y, w, h, bg)
        lines = [label] + items
        set_shape_multiline(shape, lines, font_size=9, color=tc, bold=False,
                           alignment=PP_ALIGN.LEFT)
        shape.text_frame.paragraphs[0].font.bold = True
        shape.text_frame.paragraphs[0].font.size = Pt(12)
    add_footer(slide, 19)


def slide_app_on24(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: ON24ベンチマーク — 掛け算の先例")
    data = [
        ["指標", "ON24実績", "示唆"],
        ["AI Propel+", "自動ブログ/メール/SNS/FAQ生成", "コンテンツ二次活用が鍵"],
        ["エンゲージメント", "AIコンテンツで従来比7倍", "データ駆動の効果は実証済み"],
        ["デモリクエスト", "パーソナライズで4倍増", "個別最適化がCVR向上に直結"],
        ["動画クリップ", "YoY +2,903%（13万本+）", "ウェビナー→ショート動画変換が主流に"],
        ["Cvent買収額", "ON24 ¥400億+Goldcast ¥300億", "エンゲージメントデータの市場評価"],
    ]
    add_table(slide, 6, 3, data, Inches(0.3), Inches(1.8), Inches(12.7), Inches(4.0))
    add_footer(slide, 20)


def slide_app_qa(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 想定Q&A（厳しい質問10選）")
    qas = [
        ("Q1: 30名で同時実行できるのか？", "3ギアは並行ではなく順次。Gear1は既存業務の型化。Gate Reviewで実行可能性を検証。"),
        ("Q2: 結局ウェビナーツールの延長では？", "スコアMVP×SF商談接続は「配信の先」。ON24(Cvent ¥400億で買収)が先例。"),
        ("Q3: 営業DX ¥327Mは絵に描いた餅？", "FY25 ¥33Mがシード。Gate基準¥15M/半期。未達時ピボット。"),
        ("Q4: Zoom/TeamsのAI搭載で勝てるか？", "日本適合は構造的に弱い。12-18ヶ月で「配信×データ」ポジションを確立。"),
        ("Q5: SF依存は危険？", "依存ではなくエコシステム活用。Agentforce国内最先行。中期でHubSpot/Marketo展開。"),
    ]
    for i, (q, a) in enumerate(qas):
        y = Inches(1.8 + i * 1.05)
        add_textbox(slide, Inches(0.5), y, Inches(12), Inches(0.4),
                    q, font_size=11, bold=True, color=NAVY)
        add_textbox(slide, Inches(0.7), y + Inches(0.35), Inches(12), Inches(0.5),
                    "→ " + a, font_size=10, color=DARK_TEXT)
    add_footer(slide, 21)

    # Q6-Q10
    slide2 = add_slide(prs)
    add_header_bar(slide2, "Appendix: 想定Q&A（続き）")
    qas2 = [
        ("Q6: 解約率1.0%は現実的？", "FY22:3.6%→FY24:1.7%と改善中。Gate基準は1.3%(バッファ付き)。3施策で到達可能。"),
        ("Q7: PMM/RevOps兼務で機能する？", "30名で専任は非現実的。兼務3-6ヶ月→成果で専任化。Gate Reviewで判断。"),
        ("Q8: 配信の価値は本当に続くか？", "企業ウェビナー数13→47回に急増。配信需要は拡大中。データ×AIを掛けてさらに価値を高める。"),
        ("Q9: bizibl/FanGrowthとの差別化？", "「簡単・安い」vs「データ深度×SF×伴走」。競合軸が異なり価格競争は回避。"),
        ("Q10: なぜ掛け算で市場が¥200B→¥2Tに？", "配信=動画市場。+データ=CRM/MA市場。+AI=デジタルマーケ全体。ON24/Cvent時価総額がその証左。"),
    ]
    for i, (q, a) in enumerate(qas2):
        y = Inches(1.8 + i * 1.05)
        add_textbox(slide2, Inches(0.5), y, Inches(12), Inches(0.4),
                    q, font_size=11, bold=True, color=NAVY)
        add_textbox(slide2, Inches(0.7), y + Inches(0.35), Inches(12), Inches(0.5),
                    "→ " + a, font_size=10, color=DARK_TEXT)
    add_footer(slide2, 22)


# =====================================================
# MAIN
# =====================================================
def main():
    print("=" * 50)
    print("ネクプロ 全社戦略プレゼンテーション — 配信×データ×AI 収束版")
    print("=" * 50)

    print("\n[1/2] チャート生成中...")
    generate_all_charts()

    print("\n[2/2] スライド生成中...")
    prs = new_presentation()

    slides = [
        ("Slide 1: 表紙", slide_01_title),
        ("Slide 2: 3つの掛け算", slide_02_multiplication),
        ("Slide 3: 目標KPI", slide_03_kpi_target),
        ("Slide 4: なぜ今か", slide_04_window),
        ("Slide 5: 市場", slide_05_market),
        ("Slide 6: ポジショニングマップ1", slide_06_pos_map1),
        ("Slide 7: ポジショニングマップ2", slide_07_pos_map2),
        ("Slide 8: 3ギア全体像", slide_08_three_gears_overview),
        ("Slide 9: Gear1 配信深化", slide_09_gear1),
        ("Slide 10: Gear2 データレイヤー", slide_10_gear2),
        ("Slide 11: Gear3 AIレイヤー", slide_11_gear3),
        ("Slide 12: 売上推移", slide_12_revenue),
        ("Slide 13: Gate Review", slide_13_gate_review),
        ("Slide 14: 組織・ガバナンス", slide_14_org_governance),
        ("Slide 15: 決めること", slide_15_decisions),
        ("Slide 16: クロージング", slide_16_closing),
    ]
    for label, func in slides:
        print(f"  {label}")
        func(prs)

    # Appendix
    print("  Appendix slides...")
    slide_app_revenue(prs)
    slide_app_kpi(prs)
    slide_app_swot(prs)
    slide_app_on24(prs)
    slide_app_qa(prs)  # creates 2 slides

    out_path = f'{OUT_DIR}/nexpro_strategy_presentation.pptx'
    prs.save(out_path)
    size = os.path.getsize(out_path)
    n_slides = len(prs.slides)
    print(f"\n{'=' * 50}")
    print(f"生成完了!")
    print(f"  ファイル: {out_path}")
    print(f"  スライド数: {n_slides}")
    print(f"  ファイルサイズ: {size / 1024:.0f} KB")
    print(f"{'=' * 50}")


if __name__ == '__main__':
    main()
