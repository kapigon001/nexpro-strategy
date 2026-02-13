"""メイン: 全35スライド生成"""
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


def slide_01_title(prs):
    slide = add_slide(prs)
    set_bg(slide, NAVY)
    add_textbox(slide, Inches(1), Inches(2.2), Inches(11), Inches(1.2),
                "ネクプロ 全社戦略提案", font_size=40, bold=True, color=WHITE,
                alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(3.5), Inches(11), Inches(0.8),
                "ウェビナーツールから\nB2Bエンゲージメント・インテリジェンス基盤へ",
                font_size=20, color=WHITE, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.5),
                "2026年2月 | 経営層・マネージャーMTG | Confidential",
                font_size=12, color=MED_GRAY, alignment=PP_ALIGN.CENTER)
    add_notes(slide, "本資料は経営層・マネージャー全員で共有し、本日3つの意思決定を行うための戦略提案書です。")


def slide_02_decisions(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "本日の意思決定事項",
                   "「決めないこと」は「現状維持を選ぶこと」と同義")
    items = [
        ("1", "重点投資領域の資源配分承認", "プロダクト45% / GTM35% / 組織20%"),
        ("2", "事業ポジショニング転換の承認", "ウェビナーツール → B2Bエンゲージメント・インテリジェンス基盤"),
        ("3", "組織再編の承認", "PMM/RevOps新設、KPIオーナー制度、90日実行計画"),
    ]
    for i, (num, title, sub) in enumerate(items):
        y = Inches(2.0 + i * 1.6)
        # Gold left border
        add_rect(slide, Inches(0.8), y, Inches(0.08), Inches(1.3), GOLD)
        # Navy box
        box = add_rect(slide, Inches(0.88), y, Inches(11.5), Inches(1.3), DARK_NAVY)
        # Number
        add_textbox(slide, Inches(1.2), y + Inches(0.15), Inches(0.8), Inches(0.8),
                    num, font_size=32, bold=True, color=GOLD)
        # Title
        add_textbox(slide, Inches(2.2), y + Inches(0.15), Inches(9.5), Inches(0.6),
                    title, font_size=18, bold=True, color=WHITE)
        # Subtitle
        add_textbox(slide, Inches(2.2), y + Inches(0.75), Inches(9.5), Inches(0.4),
                    sub, font_size=12, color=MED_GRAY)
    add_footer(slide, 2)
    add_notes(slide, "冒頭で「今日のゴール」を共有。議論が散漫にならないよう最後にこの3つに戻ります。")


def slide_03_exec_summary(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "エグゼクティブサマリー",
                   "「配信ツール」に留まれば淘汰される。「データ基盤」に転換すれば勝てる")
    bullets = [
        "AIエージェント時代、SaaS UIレイヤーは圧縮。データレイヤーの価値が増大",
        "ネクプロの勝ち筋：Salesforce × エンゲージメントデータで唯一の存在に",
        "推奨：「竹」戦略で6ヶ月実証 → Gate Review → 「松」へ段階移行",
        "FY27目標：売上¥1,382M（FY24比2.7倍）、月次解約率1.0%",
    ]
    add_multiline_textbox(slide, Inches(0.5), Inches(2.0), Inches(6.5), Inches(4.5),
                          bullets, font_size=13, bullet=True, color=DARK_TEXT)
    # KPI table
    data = [
        ["指標", "FY24実績", "FY27目標"],
        ["売上", "¥512M", "¥1,382M"],
        ["ARPA", "¥148K/月", "¥204K/月"],
        ["月次解約率", "1.7%", "1.0%"],
        ["新収益比率", "0%", "27%"],
    ]
    add_table(slide, 5, 3, data, Inches(7.5), Inches(2.2), Inches(5.3), Inches(3.0))
    add_footer(slide, 3)
    add_notes(slide, "結論を先に。ネクプロの勝ち筋は「配信競争から離脱し、エンゲージメントデータ×SF連携で唯一の存在になる」こと。")


def slide_05_saaspocalypse(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "SaaS中間層の圧縮",
                   "AIエージェントはUIレイヤーを不要にするが、データレイヤーの価値は増大する")
    # 3層図
    layers = [
        (Inches(0.8), Inches(2.2), BLUE, "AIエージェント層（実行） ↑拡大"),
        (Inches(0.8), Inches(3.5), RED, "SaaS UIレイヤー（圧縮対象） ←→圧縮"),
        (Inches(0.8), Inches(4.8), GREEN, "データ基盤層（価値増大） ↑拡大"),
    ]
    for x, y, color, text in layers:
        shape = add_rounded_rect(slide, x, y, Inches(5.5), Inches(1.0), color)
        set_shape_text(shape, text, font_size=14, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
    # ファクト表
    data = [
        ["ファクト", "数値"],
        ["SaaS株価消失", "$2,850億(約42兆円)"],
        ["バーティカルSaaS下落", "-43%"],
        ["ワークフローSaaS下落", "-39%"],
        ["AIエージェントPJ中止予測", "40%+(2027年末)"],
        ["日本SaaS浸透率", "4%(米国15-18%)"],
    ]
    add_table(slide, 6, 2, data, Inches(7.0), Inches(2.2), Inches(5.8), Inches(3.8))
    add_footer(slide, 5)
    add_notes(slide, "SaaSpocalypseの数字で危機感を醸成。ただし「日本には時間がある」と次スライドで希望も示します。")


def slide_06_japan_market(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "日本市場の構造的特性",
                   "SaaS浸透率4%の日本は12-18ヶ月の先行者優位。今動けば間に合う")
    # チャート画像
    slide.shapes.add_picture(f'{CHART_DIR}/market.png',
                             Inches(0.3), Inches(1.9), Inches(6.5), Inches(5.0))
    # 障壁ボックス
    barriers = [
        "SaaS浸透率 4%（米国15-18%）",
        "日本語選好 72%のバイヤー",
        "デジタル人材不足 230万人",
        "商習慣の壁（稟議・根回し文化）",
    ]
    for i, text in enumerate(barriers):
        y = Inches(2.2 + i * 1.15)
        box = add_rounded_rect(slide, Inches(7.2), y, Inches(5.5), Inches(0.9), LIGHT_GRAY, NAVY)
        set_shape_text(box, f"🛡 {text}", font_size=12, bold=True, color=NAVY,
                      alignment=PP_ALIGN.LEFT)
    add_footer(slide, 6)
    add_notes(slide, "日本のAI統合は12-18ヶ月遅れ。これはネクプロの猶予時間だが永続しない。")


def slide_08_revenue(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "売上・成長推移",
                   "FY24成長率+3.1%。新収益柱なしでは¥1B突破は構造的に困難")
    slide.shapes.add_picture(f'{CHART_DIR}/revenue.png',
                             Inches(0.5), Inches(1.9), Inches(12.3), Inches(5.2))
    add_footer(slide, 8)
    add_notes(slide, "FY24の+3.1%成長を正面から見せる。新収益柱がFY27で27%を占める計画。")


def slide_09_kpi(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "SaaS KPI課題",
                   "月次解約率1.7%(年換算18.5%)は「バケツの穴」")
    data = [
        ["指標", "FY22", "FY23", "FY24", "FY25計画"],
        ["長期PFアカウント数", "160", "151", "167", "179"],
        ["ARPA長期(千円/月)", "-", "¥137.6K", "¥148.0K", "¥168.7K"],
        ["月次解約率(長期)", "3.6%", "2.3%", "1.7%", "1.0%(目標)"],
        ["新規長期PF成約/年", "60", "27", "50", "38"],
        ["成約率", "11.2%", "9.5%", "10.5%", "-"],
        ["月間商談数", "~44", "~23", "~44", "~34"],
    ]
    t = add_table(slide, 7, 5, data, Inches(0.5), Inches(2.0), Inches(12.3), Inches(4.5),
                  highlight_rows=[3], highlight_color=RGBColor(0xFA, 0xDB, 0xD8))
    add_footer(slide, 9)
    add_notes(slide, "解約率1.7%(年換算18.5%)はSaaS優良企業の目安(年5%以下)を大幅超過。改善トレンドはあるが施策が必要。")


def slide_10_swot(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "SWOT分析", "「データモート」が唯一の持続的優位性")
    qdata = {
        "S（強み）": (Inches(0.3), Inches(2.0), Inches(6.2), Inches(2.5), NAVY, WHITE, [
            "1. ITreview 13期連続受賞の製品品質",
            "2. Salesforce Agentforce国内最先行",
            "3. 400-500社エンタープライズ顧客基盤",
            "4. 日本語運用ノウハウ・伴走支援",
            "5. 擬似ライブ配信・メディアサイト機能",
        ]),
        "W（弱み）": (Inches(6.8), Inches(2.0), Inches(6.2), Inches(2.5), GOLD, DARK_TEXT, [
            "1. 30名の少数精鋭体制",
            "2. 総調達額¥7.5億(対Zoom/ON24比)",
            "3. ブランド認知度(Zoom,V-CUBEに劣後)",
            "4. インテントデータ未整備(ON24は40-50pt/人)",
            "5. PMM/RevOps機能の不在",
        ]),
        "O（機会）": (Inches(0.3), Inches(4.7), Inches(6.2), Inches(2.5), GREEN, WHITE, [
            "1. 国内AI統合12-18ヶ月遅れ(先行者優位)",
            "2. ON24のCvent買収→日本市場優先度低下",
            "3. 3rdパーティCookie廃止→1stパーティデータ価値増",
            "4. 日本B2B DX余地(SaaS浸透率4%)",
            "5. 企業ウェビナー数急増(13→47回/年)",
        ]),
        "T（脅威）": (Inches(6.8), Inches(4.7), Inches(6.2), Inches(2.5), RED, WHITE, [
            "1. Zoom/Webex/TeamsのAIエージェント搭載",
            "2. AIによる配信機能コモディティ化",
            "3. Cvent大型統合(ON24¥400億+Goldcast¥300億)",
            "4. 価格競争激化(bizibl/Cocripo低価格攻勢)",
            "5. 汎用AIで差別化の希薄化リスク",
        ]),
    }
    for label, (x, y, w, h, bg, tc, items) in qdata.items():
        shape = add_rounded_rect(slide, x, y, w, h, bg)
        lines = [label] + items
        set_shape_multiline(shape, lines, font_size=9, color=tc, bold=False,
                           alignment=PP_ALIGN.LEFT)
        # Make first line bold
        shape.text_frame.paragraphs[0].font.bold = True
        shape.text_frame.paragraphs[0].font.size = Pt(12)
    add_footer(slide, 10)
    add_notes(slide, "W×T象限が最大リスク：何もしなければ3年で地位喪失。S×O象限で「SF×AI先行でデータモート構築」が勝ち筋。")


def slide_11_cross_swot(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "クロスSWOT戦略示唆",
                   "S×Oの「データモート構築」とW×Tの「最大リスク回避」が戦略の両輪")
    data = [
        ["", "機会(O)", "脅威(T)"],
        ["強み(S)", "SF統合×AI先行で\n「データモート」構築\n→エンゲージメントスコアMVP",
         "日本適合性×伴走支援で\nグローバル勢と差別化\n→業種別パッケージ"],
        ["弱み(W)", "少数精鋭×AI遅れの隙間で\nニッチ勝ち\n→SF顧客に選択と集中",
         "【最大リスク】\n資金・人材不足×コモディティ化\n→何もしなければ地位喪失"],
    ]
    t = add_table(slide, 3, 3, data, Inches(0.5), Inches(2.2), Inches(12.3), Inches(4.5))
    add_footer(slide, 11)
    add_notes(slide, "SWOTは整理ツールではなく戦略の出発点。この4象限から重点施策が導かれます。")


def slide_12_competitor_table(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "競合比較", "「日本×データ」の空白ポジションが存在する")
    data = [
        ["評価軸", "ネクプロ", "Zoom", "ON24", "EventHub", "bizibl", "FanGrowth"],
        ["配信品質",    "◎", "◎", "○", "○", "○", "△"],
        ["MA/CRM連携",  "◎", "○", "◎", "△", "△", "△"],
        ["日本語適合",   "◎", "△", "△", "◎", "◎", "◎"],
        ["データ分析",   "○", "△", "◎", "△", "△", "△"],
        ["AI機能",      "○", "○", "◎", "△", "△", "△"],
        ["価格柔軟性",   "◎", "○", "△", "○", "◎", "◎"],
        ["導入支援",     "◎", "△", "○", "○", "○", "◎"],
    ]
    add_table(slide, 8, 7, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(4.8))
    add_footer(slide, 12)
    add_notes(slide, "ネクプロはMA/CRM連携と日本語適合で強い。データ分析とAI機能が投資ポイント。")


def slide_13_pos_map1(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "ポジショニングマップ(1): 機能深度 × 日本企業適合性",
                   "右上の「高機能×高適合」象限は空白。ネクプロが唯一到達可能")
    slide.shapes.add_picture(f'{CHART_DIR}/pos_map1.png',
                             Inches(1.5), Inches(1.8), Inches(10.3), Inches(5.5))
    add_footer(slide, 13)
    add_notes(slide, "右上の「高機能×高適合」象限は空白。ON24は機能は深いが日本適合が弱い。国内勢は逆。")


def slide_14_pos_map2(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "ポジショニングマップ(2): データ活用高度性 × 導入容易性",
                   "「高データ活用×低障壁」のSweet Spotを狙う")
    slide.shapes.add_picture(f'{CHART_DIR}/pos_map2.png',
                             Inches(1.5), Inches(1.8), Inches(10.3), Inches(5.5))
    add_footer(slide, 14)
    add_notes(slide, "ON24はデータ活用高いが導入が重い。ネクプロはテンプレ・伴走で障壁を下げつつデータ活用を高度化。")


def slide_16_mece(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "全社課題マップ（MECE）",
                   "最優先は「エンゲージメントデータ基盤化」と「解約率改善」")
    cols = [
        ("市場", [("AI時代の事業再定義遅れ", RED), ("ウェビナー市場コモディティ化", GOLD)]),
        ("顧客", [("成果指標(商談化)接続が弱い", RED), ("利用深度のばらつき", GOLD)]),
        ("プロダクト", [("エンゲージメントスコア未実装", RED), ("AI自動生成未実装", GOLD)]),
        ("GTM", [("業種別テンプレ不足", GOLD), ("価格体系が旧来型", GOLD)]),
        ("組織", [("PMM機能の不在", RED), ("30名体制ボトルネック", RED)]),
        ("財務", [("新収益柱の実行リスク", RED), ("LTV拡張余地", GOLD)]),
    ]
    col_w = Inches(2.0)
    for ci, (header, cards) in enumerate(cols):
        x = Inches(0.3 + ci * 2.1)
        # Header
        hdr = add_rect(slide, x, Inches(2.0), col_w, Inches(0.5), NAVY)
        set_shape_text(hdr, header, font_size=12, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
        for ri, (text, color) in enumerate(cards):
            y = Inches(2.7 + ri * 1.8)
            card = add_rounded_rect(slide, x, y, col_w, Inches(1.5), color)
            set_shape_text(card, text, font_size=10, bold=True, color=WHITE,
                          alignment=PP_ALIGN.CENTER)
    # 凡例
    add_rect(slide, Inches(0.5), Inches(6.8), Inches(0.3), Inches(0.3), RED)
    add_textbox(slide, Inches(0.9), Inches(6.8), Inches(1.5), Inches(0.3),
                "= 最優先", font_size=9, color=DARK_TEXT)
    add_rect(slide, Inches(2.5), Inches(6.8), Inches(0.3), Inches(0.3), GOLD)
    add_textbox(slide, Inches(2.9), Inches(6.8), Inches(1.5), Inches(0.3),
                "= 重要", font_size=9, color=DARK_TEXT)
    add_footer(slide, 16)
    add_notes(slide, "15課題をMECE整理。最優先2つ：エンゲージメントスコア実装と解約率改善。")


def slide_17_priority(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "優先度マトリクス",
                   "象限Iの5課題に90%のマネジメント注力を配分")
    # Axes
    add_rect(slide, Inches(1.5), Inches(2.0), Inches(0.05), Inches(4.8), MED_GRAY)  # Y axis
    add_rect(slide, Inches(1.5), Inches(6.8), Inches(11.0), Inches(0.05), MED_GRAY)  # X axis
    add_textbox(slide, Inches(0.2), Inches(2.0), Inches(1.2), Inches(0.4),
                "高\n緊急度", font_size=9, color=DARK_TEXT, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.2), Inches(6.2), Inches(1.2), Inches(0.4),
                "低", font_size=9, color=DARK_TEXT, alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(11.5), Inches(6.9), Inches(1.5), Inches(0.3),
                "高 財務インパクト →", font_size=9, color=DARK_TEXT)
    # Quadrant I (top-right) - Red items
    q1 = ["エンゲージメントスコアMVP", "解約率改善(1.7%→1.0%)",
           "事業価値再定義", "PMM新設", "商談化接続強化"]
    for i, t in enumerate(q1):
        y = Inches(2.3 + i * 0.85)
        dot = add_rect(slide, Inches(7.5 + (i%2)*1.5), y, Inches(3.8), Inches(0.6), RED)
        set_shape_text(dot, t, font_size=10, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
    # Quadrant II (top-left) - Gold
    q2 = ["CS→プロフィット化", "価格再設計"]
    for i, t in enumerate(q2):
        dot = add_rect(slide, Inches(2.0), Inches(2.5 + i*1.0), Inches(3.5), Inches(0.6), GOLD)
        set_shape_text(dot, t, font_size=10, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    # Quadrant III (bottom-right)
    q3 = ["新収益柱管理", "AI自動生成"]
    for i, t in enumerate(q3):
        dot = add_rect(slide, Inches(7.5), Inches(5.5 + i*0.8), Inches(3.5), Inches(0.55), BLUE)
        set_shape_text(dot, t, font_size=10, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    add_footer(slide, 17)
    add_notes(slide, "散漫にならないよう最優先5課題にフォーカス。")


def slide_19_options(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "3つの戦略オプション",
                   "「竹」で6ヶ月実証し、KPI達成を条件に「松」へ段階移行")
    data = [
        ["項目", "梅：防衛型", "竹：均衡型【推奨】", "松：攻勢型"],
        ["コンセプト", "既存効率最大化", "既存深耕+データ拡張", "基盤転換+組織再編"],
        ["投資規模", "現行維持", "+30-50%", "+80-100%"],
        ["FY27売上", "¥740M(仮説)", "¥930M(仮説)", "¥1,382M"],
        ["プロダクト", "UI改善のみ", "スコアMVP+AI要約", "フルスタック転換"],
        ["GTM", "効率化のみ", "業種別PKG+CS高度化", "新セグメント+価格再設計"],
        ["組織", "現行維持", "PMM兼務設置", "PMM/RevOps正式新設"],
        ["リスク", "低→高(中長期)", "中", "高"],
    ]
    add_table(slide, 8, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(5.0))
    add_footer(slide, 19)
    add_notes(slide, "3案を公平に提示。竹案を推奨する理由は次スライドで説明。")


def slide_20_recommended(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "推奨戦略の論拠",
                   "全力投球ではなく、実績で信頼を獲得しながらギアを上げる")
    # 3 logic boxes
    logics = [
        ("1", "キャッシュ耐性", "30名・¥7.5億で全方位投資は自殺行為"),
        ("2", "小さく証明", "MVP 6ヶ月→5社実証→本格投資"),
        ("3", "段階構築", "兼務→成果確認→専任化"),
    ]
    for i, (num, title, desc) in enumerate(logics):
        x = Inches(0.5 + i * 4.2)
        box = add_rounded_rect(slide, x, Inches(2.0), Inches(3.8), Inches(1.2), NAVY)
        set_shape_multiline(box, [f" {num}  {title}", f"    {desc}"],
                           font_size=11, color=WHITE, bold=False)
        box.text_frame.paragraphs[0].font.bold = True
        box.text_frame.paragraphs[0].font.size = Pt(14)
        box.text_frame.paragraphs[0].font.color.rgb = GOLD
    # Do/Don't table
    data = [
        ["やること ✓", "やらないこと ✗"],
        ["エンゲージメントスコアMVP", "汎用AIチャットボット開発"],
        ["Salesforce連携深化", "HubSpot/Marketo(中期以降)"],
        ["業種別ソリューション化", "低価格競争"],
        ["CS→プロフィットセンター化", "機能の無秩序追加"],
        ["解約率改善(1.7%→1.0%)", "新規大量獲得(質を優先)"],
    ]
    add_table(slide, 6, 2, data, Inches(0.5), Inches(3.6), Inches(12.3), Inches(3.3),
              header_color=GREEN)
    add_footer(slide, 20)
    add_notes(slide, "「やらないこと」の明示が重要。30名体制では全てはできない。")


def slide_21_gate(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "段階移行の判断基準",
                   "6ヶ月後のGate Reviewでファクトベースに松案移行を判断")
    data = [
        ["KPI", "Gate基準", "達成時", "未達時"],
        ["スコアMVP導入社数", "5社以上", "松案移行", "竹延長・修正"],
        ["月次解約率", "≤1.3%", "松案移行", "CS施策見直し"],
        ["ARPA(長期)", "≥¥160K", "松案移行", "価格体系再検討"],
        ["営業DX売上", "≥¥15M/半期", "投資拡大", "ピボット検討"],
    ]
    add_table(slide, 5, 4, data, Inches(0.5), Inches(2.0), Inches(12.3), Inches(3.0))
    # Timeline
    add_rect(slide, Inches(1.5), Inches(5.5), Inches(10.0), Inches(0.08), NAVY)
    # M0
    add_textbox(slide, Inches(1.2), Inches(5.7), Inches(1), Inches(0.4),
                "M0\n開始", font_size=10, bold=True, color=NAVY, alignment=PP_ALIGN.CENTER)
    # M6 Gate
    gate = add_rounded_rect(slide, Inches(5.5), Inches(5.2), Inches(2.5), Inches(0.6), RED)
    set_shape_text(gate, "M6: Gate Review", font_size=12, bold=True, color=WHITE,
                  alignment=PP_ALIGN.CENTER)
    # Branches
    add_textbox(slide, Inches(8.5), Inches(5.2), Inches(3), Inches(0.4),
                "達成 → 松案移行", font_size=11, bold=True, color=GREEN)
    add_textbox(slide, Inches(8.5), Inches(5.8), Inches(3), Inches(0.4),
                "未達 → 竹延長", font_size=11, bold=True, color=RED)
    add_footer(slide, 21)
    add_notes(slide, "Gate Reviewは「計画通りか」ではなく「投資を続けるべきか」を判断する場。")


def slide_22_product(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "重点施策A：プロダクト",
                   "最優先はエンゲージメントスコアMVPとSalesforce連携強化")
    data = [
        ["#", "施策", "時期", "成果KPI", "難易度"],
        ["A1", "エンゲージメントスコアMVP", "0-6M", "5社導入、ARPA+20%", "高"],
        ["A2", "Salesforce連携強化", "0-6M", "連携率80%", "中"],
        ["A3", "AI日本語コンテンツ自動生成", "6-12M", "採用率30%", "高"],
        ["A4", "API-first化", "6-18M", "API網羅率80%", "高"],
        ["A5", "HubSpot/Marketo連携", "12-18M", "新規20社/年", "中"],
    ]
    add_table(slide, 6, 5, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(4.0))
    add_footer(slide, 22)
    add_notes(slide, "A1とA2が0-6ヶ月の最優先。A3以降は中期で順次。")


def slide_23_sales_cs(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "重点施策B/C：営業 × CS",
                   "業種別パッケージで成約率向上、CSをプロフィットセンター化")
    # Sales table (left)
    add_textbox(slide, Inches(0.3), Inches(1.9), Inches(2), Inches(0.4),
                "B: 営業", font_size=14, bold=True, color=NAVY)
    sd = [
        ["#", "施策", "時期", "成果KPI"],
        ["B1", "業種別PKG×4", "0-6M", "成約率+5pt"],
        ["B2", "価格体系再設計", "6-12M", "ARPA+15%"],
        ["B3", "営業DX加速", "0-12M", "¥33M→¥121M"],
        ["B4", "SFエコシステム販路", "6-18M", "月2件以上"],
    ]
    add_table(slide, 5, 4, sd, Inches(0.3), Inches(2.4), Inches(6.2), Inches(3.0))
    # CS table (right)
    add_textbox(slide, Inches(6.8), Inches(1.9), Inches(2), Inches(0.4),
                "C: CS", font_size=14, bold=True, color=NAVY)
    cd = [
        ["#", "施策", "時期", "成果KPI"],
        ["C1", "オンボーディング標準化", "0-6M", "60日完了90%"],
        ["C2", "ヘルススコア導入", "0-6M", "解約予測70%"],
        ["C3", "CS収益化", "6-12M", "年¥30M"],
        ["C4", "戦略アカウント管理", "0-12M", "Top20 NRR120%+"],
    ]
    add_table(slide, 5, 4, cd, Inches(6.8), Inches(2.4), Inches(6.2), Inches(3.0))
    add_footer(slide, 23)
    add_notes(slide, "B1(業種別PKG)とC1(オンボーディング標準化)は即着手可能。")


def slide_24_org(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "重点施策D：組織",
                   "「兼務→成果確認→専任化」の段階アプローチ")
    data = [
        ["#", "施策", "時期", "成果KPI"],
        ["D1", "PMM機能新設(兼務)", "M1-3", "競合レポート月次化"],
        ["D2", "RevOps機能(兼務)", "M3-6", "ファネル可視化100%"],
        ["D3", "戦略採用(AI/PMM/CS/SalesDX)", "M1-12", "+3-5名"],
        ["D4", "KPIオーナー制度導入", "M1", "全主要KPIに個人名"],
    ]
    add_table(slide, 5, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(2.8))
    # Org chart
    boxes = [
        (Inches(5.5), Inches(5.0), "CEO"),
        (Inches(1.5), Inches(6.2), "PMM(兼務)"),
        (Inches(3.8), Inches(6.2), "RevOps(兼務)"),
        (Inches(6.1), Inches(6.2), "Prod"),
        (Inches(8.4), Inches(6.2), "Sales"),
        (Inches(10.7), Inches(6.2), "CS"),
    ]
    for x, y, text in boxes:
        box = add_rounded_rect(slide, x, y, Inches(1.8), Inches(0.6), NAVY)
        set_shape_text(box, text, font_size=11, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
    add_footer(slide, 24)
    add_notes(slide, "D4(KPIオーナー制度)は来週着手可能。各KPIに個人名を紐付ける。")


def slide_25_roadmap(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "36ヶ月ロードマップ",
                   "止血(0-6M) → 転換(6-18M) → 成長(18-36M)")
    phases = [
        ("Phase1: 止血・改善 (0-6M)", GREEN, Inches(2.0), [
            "KPIオーナー全員アサイン(M1)", "PMM兼務スタート(M1-3)",
            "オンボーディング標準化(M2)", "業種別PKG 4種(M3)",
            "ヘルススコア導入(M4-6)", "★スコアMVP完成 / Gate Review(M6)",
        ]),
        ("Phase2: 転換・仕込み (6-18M)", GOLD, Inches(3.8), [
            "分析ティア化(M8)", "価格3層構造(M9)",
            "AIコンテンツ生成(M10)", "CS収益化(M12)",
            "営業DX ¥100M/年ペース(M12)", "★採用+3-5名 / Gate Review(M12)",
        ]),
        ("Phase3: 成長加速 (18-36M)", BLUE, Inches(5.6), [
            "HubSpot/Marketo連携(M18-24)", "API-first完了(M24)",
            "リポジショニング完了(M24)", "売上¥1B突破(M24-30)",
            "営業DX ¥300M/年(M36)", "★M&A検討 / Gate Review(M36)",
        ]),
    ]
    for title, color, y, items in phases:
        # Phase header
        hdr = add_rect(slide, Inches(0.3), y, Inches(3.0), Inches(1.4), color)
        set_shape_text(hdr, title, font_size=11, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
        # Items
        for i, item in enumerate(items):
            x = Inches(3.5 + (i % 3) * 3.3)
            iy = y + Inches(0 if i < 3 else 0.7)
            box = add_rounded_rect(slide, x, iy, Inches(3.1), Inches(0.55),
                                   LIGHT_GRAY, color)
            fc = RED if "★" in item else DARK_TEXT
            set_shape_text(box, item, font_size=8, bold=("★" in item),
                          color=fc, alignment=PP_ALIGN.LEFT)
    add_footer(slide, 25)
    add_notes(slide, "最初の6ヶ月に最も具体的な情報。Gate Reviewのタイミングを明示。")


def slide_26_kpi_tree(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "KPIツリー＆ガバナンス",
                   "North Star → 事業KPI → 先行KPIの因果連鎖を設計")
    # North Star
    ns = add_rounded_rect(slide, Inches(3.5), Inches(2.0), Inches(6.3), Inches(0.6), NAVY)
    set_shape_text(ns, "North Star: 顧客エンゲージメント成果価値（商談化寄与額）",
                  font_size=11, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    # Level 1
    l1 = ["ARR成長率", "NRR(目標110%+)", "粗利率", "顧客基盤"]
    for i, t in enumerate(l1):
        x = Inches(0.8 + i * 3.2)
        box = add_rounded_rect(slide, x, Inches(3.0), Inches(2.8), Inches(0.5), BLUE)
        set_shape_text(box, t, font_size=10, bold=True, color=WHITE,
                      alignment=PP_ALIGN.CENTER)
    # Level 2
    l2 = [
        ["MRR", "新規MRR", "営業DX売上"],
        ["Expansion", "Churn MRR", ""],
        ["ARPA", "オプション売上", ""],
        ["累計アカウント", "月次解約率", "成約率"],
    ]
    for i, items in enumerate(l2):
        for j, t in enumerate(items):
            if not t:
                continue
            x = Inches(0.5 + i * 3.2 + j * 1.0)
            box = add_rounded_rect(slide, x, Inches(3.8 + j * 0.0), Inches(0.95), Inches(0.4),
                                   LIGHT_GRAY, BLUE)
            set_shape_text(box, t, font_size=7, color=DARK_TEXT, alignment=PP_ALIGN.CENTER)
    # Governance table
    gdata = [
        ["会議体", "頻度", "参加者", "議題"],
        ["経営KPIレビュー", "月次", "CEO+KPIオーナー", "KPI進捗確認"],
        ["戦略レビュー", "四半期", "全マネージャー", "戦略方向性確認"],
        ["Gate Review", "M6/M12/M18", "経営層全員", "投資継続判断"],
        ["スプリントレビュー", "隔週", "開発チーム", "実行進捗"],
    ]
    add_table(slide, 5, 4, gdata, Inches(0.5), Inches(4.8), Inches(12.3), Inches(2.3))
    add_footer(slide, 26)
    add_notes(slide, "KPIオーナーを個人名で紐付けることが核心。月次レビューで進捗管理。")


def slide_27_decision_closing(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "本会議で決めること",
                   "今日決めなければ「何もしない」を選択したことになる")
    data = [
        ["#", "決議内容", "選択肢", "推奨"],
        ["1", "投資配分(P45/GTM35/Org20)", "A:承認/B:修正/C:保留", "A"],
        ["2", "ポジショニング転換承認", "A:承認/B:段階的/C:不承認", "A"],
        ["3", "組織再編(PMM/RevOps/KPI)", "A:承認/B:PMM先行/C:保留", "A"],
        ["4", "90日実行計画承認", "A:承認/B:範囲縮小/C:保留", "A"],
        ["5", "Gate Review設定(M6)", "A:承認/B:期間変更/C:保留", "A"],
    ]
    add_table(slide, 6, 4, data, Inches(0.3), Inches(2.0), Inches(12.7), Inches(2.8))
    # Next Steps
    ns = [
        ["アクション", "担当", "期限"],
        ["KPIオーナーアサイン", "CEO", "1週間以内"],
        ["PMM兼務者決定", "マネージャー会議", "2週間以内"],
        ["スコアMVP要件定義", "プロダクト", "1ヶ月以内"],
        ["業種別PKG設計", "営業", "1ヶ月以内"],
        ["Gate Review日程確定", "CEO", "1週間以内"],
    ]
    add_table(slide, 6, 3, ns, Inches(0.3), Inches(5.0), Inches(12.7), Inches(2.2),
              header_color=GREEN)
    add_footer(slide, 27)
    add_notes(slide, "沈黙を恐れず一つずつ確認。「保留」は実質的に「不承認」と同義であることを伝える。")


def slide_28_closing(prs):
    slide = add_slide(prs)
    set_bg(slide, NAVY)
    add_textbox(slide, Inches(1), Inches(2.8), Inches(11), Inches(1),
                "ご清聴ありがとうございます", font_size=32, bold=True, color=WHITE,
                alignment=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(1), Inches(4.2), Inches(11), Inches(0.8),
                "Q&A", font_size=24, bold=True, color=GOLD,
                alignment=PP_ALIGN.CENTER)
    add_notes(slide, "想定Q&A(10問)を準備済み。厳しい質問を歓迎する姿勢を見せる。")


# === Appendix slides ===
def slide_29_revenue_detail(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 売上内訳詳細")
    data = [
        ["", "FY22", "FY23", "FY24", "FY25P", "FY26P", "FY27P"],
        ["売上合計", "¥418M", "¥497M", "¥512M", "¥649M", "¥911M", "¥1,382M"],
        ["YoY", "-", "+19.0%", "+3.1%", "+26.7%", "+40.3%", "+51.6%"],
        ["MRR年額", "¥225M", "¥252M", "¥287M", "¥330M", "¥417M", "¥526M"],
        ["オプション", "¥192M", "¥244M", "¥225M", "¥319M", "¥494M", "¥855M"],
        ["コンパウンド", "-", "-", "-", "¥3M", "¥17M", "¥51M"],
        ["営業DX", "-", "-", "-", "¥33M", "¥121M", "¥327M"],
    ]
    add_table(slide, 7, 7, data, Inches(0.3), Inches(1.8), Inches(12.7), Inches(4.5))
    add_footer(slide, 29)
    add_notes(slide, "財務詳細。FY25以降は計画値。")


def slide_30_churn(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 解約率推移")
    data = [
        ["指標", "FY22", "FY23", "FY24", "目標"],
        ["月次解約率(長期)", "3.6%", "2.3%", "1.7%", "1.0%"],
        ["年換算", "35.6%", "24.3%", "18.5%", "11.4%"],
        ["改善幅", "-", "-1.3pt", "-0.6pt", "-0.7pt"],
    ]
    add_table(slide, 4, 5, data, Inches(1.0), Inches(2.0), Inches(11.0), Inches(2.5))
    add_textbox(slide, Inches(1.0), Inches(5.0), Inches(11.0), Inches(1.0),
                "SaaS優良企業の目安: 月次解約率 < 0.42%（年5%以下）\n"
                "ネクプロは改善トレンドにあるが、目標到達にはCS施策の抜本強化が必要",
                font_size=13, color=DARK_TEXT)
    add_footer(slide, 30)
    add_notes(slide, "解約率の改善トレンドは続いているが、ベストプラクティスとの差は依然大きい。")


def slide_31_market_data(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 市場データ")
    data = [
        ["市場", "2023", "2024", "2025(予)", "2028(予)", "CAGR"],
        ["法人向け動画配信", "¥513B", "¥635B", "¥780B", "¥1,529B", "24.4%"],
        ["デジタルマーケ(CRM/MA)", "-", "¥3,672B", "¥4,190B", "-", "+14.1%"],
        ["グローバルWeb会議", "-", "$6.36B", "-", "$11.28B", "12.1%"],
    ]
    add_table(slide, 4, 6, data, Inches(0.5), Inches(2.0), Inches(12.3), Inches(2.5))
    add_footer(slide, 31)
    add_notes(slide, "動画配信市場はCAGR24.4%で急成長中。デジタルマーケも14%成長。")


def slide_32_competitor_profiles(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 競合プロファイル")
    data = [
        ["企業", "ポジション", "強み", "弱み"],
        ["Zoom", "汎用会議+ウェビナー", "圧倒的認知度・低価格", "日本適合△、分析浅い"],
        ["ON24", "IEP(データ分析型)", "40-50データpt/人、AI", "日本市場非注力、導入重い"],
        ["EventHub", "イベント管理統合", "ハイブリッド対応", "ウェビナー特化ではない"],
        ["bizibl", "簡単ウェビナー", "低価格・簡単", "分析・連携が弱い"],
        ["FanGrowth", "成果報酬型", "リスク低い導入", "機能限定的"],
        ["V-CUBE", "大規模配信", "放送品質・運用支援", "高コスト・SaaS的でない"],
    ]
    add_table(slide, 7, 4, data, Inches(0.3), Inches(1.8), Inches(12.7), Inches(5.0))
    add_footer(slide, 32)
    add_notes(slide, "各社の特徴を整理。ネクプロの差別化ポイントは「日本×データ×SF連携」。")


def slide_33_on24(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: ON24ベンチマーク")
    data = [
        ["指標", "ON24実績", "示唆"],
        ["AI Propel+", "自動ブログ/メール/SNS/FAQ生成", "コンテンツ二次活用が鍵"],
        ["エンゲージメント", "AIコンテンツで従来比7倍", "データ駆動の効果は実証済み"],
        ["デモリクエスト", "パーソナライズで4倍増", "個別最適化がCVR向上に直結"],
        ["動画クリップ", "YoY +2,903%（13万本+）", "ウェビナー→ショート動画変換が主流に"],
        ["Cvent買収額", "ON24 ¥400億+Goldcast ¥300億", "エンゲージメントデータの市場評価"],
    ]
    add_table(slide, 6, 3, data, Inches(0.3), Inches(1.8), Inches(12.7), Inches(4.0))
    add_footer(slide, 33)
    add_notes(slide, "ON24の数字は「エンゲージメントデータ基盤」の市場価値を実証している。")


def slide_34_qa(prs):
    slide = add_slide(prs)
    add_header_bar(slide, "Appendix: 想定Q&A（厳しい質問10選）")
    qas = [
        ("Q1: 30名で同時実行できるのか？", "竹案は5課題に絞る「選択と集中」。兼務→専任の段階アプローチ。"),
        ("Q2: 結局ウェビナーツールの延長では？", "エンゲージメントスコアは「配信の先」。SF商談接続で新カテゴリー定義。"),
        ("Q3: 営業DX ¥327Mは絵に描いた餅？", "FY25 ¥33Mがシード。Gate基準¥15M/半期。未達時ピボット。"),
        ("Q4: Zoom/TeamsのAI搭載で勝てるか？", "日本適合の構造的弱さ。12-18ヶ月でデータモート構築。"),
        ("Q5: Salesforce依存は危険？", "依存ではなくエコシステム活用。中期でHubSpot/Marketo展開。"),
    ]
    for i, (q, a) in enumerate(qas):
        y = Inches(1.8 + i * 1.1)
        add_textbox(slide, Inches(0.5), y, Inches(12), Inches(0.4),
                    q, font_size=11, bold=True, color=NAVY)
        add_textbox(slide, Inches(0.7), y + Inches(0.35), Inches(12), Inches(0.5),
                    "→ " + a, font_size=10, color=DARK_TEXT)
    add_footer(slide, 34)
    add_notes(slide, "Q1-Q5。Q6-Q10は次スライド。")

    # Q6-Q10
    slide2 = add_slide(prs)
    add_header_bar(slide2, "Appendix: 想定Q&A（続き）")
    qas2 = [
        ("Q6: 解約率1.0%は現実的？", "FY22:3.6%→FY24:1.7%と改善中。Gate基準は1.3%(バッファ付き)。"),
        ("Q7: PMM/RevOps兼務で機能する？", "兼務3-6ヶ月→成果で専任化。30名で最初から専任は非現実的。"),
        ("Q8: FY26-27計画は攻めすぎ？", "竹案FY27は¥930M(仮説)。松案¥1,382Mは条件付き目標。"),
        ("Q9: bizibl/FanGrowthとの差別化？", "「簡単・安い」vs「データ深度×SF×伴走」。競合軸が異なる。"),
        ("Q10: ウェビナー市場は縮小しない？", "逆に拡大(13→47回/年)。「配信だけ」は縮小、「データ活用」は成長。"),
    ]
    for i, (q, a) in enumerate(qas2):
        y = Inches(1.8 + i * 1.1)
        add_textbox(slide2, Inches(0.5), y, Inches(12), Inches(0.4),
                    q, font_size=11, bold=True, color=NAVY)
        add_textbox(slide2, Inches(0.7), y + Inches(0.35), Inches(12), Inches(0.5),
                    "→ " + a, font_size=10, color=DARK_TEXT)
    add_footer(slide2, 35)
    add_notes(slide2, "Q6-Q10。全質問に対し具体的数値と論拠で回答可能。")


# === MAIN ===
def main():
    print("=" * 50)
    print("ネクプロ 全社戦略プレゼンテーション生成")
    print("=" * 50)

    print("\n[1/2] チャート生成中...")
    generate_all_charts()

    print("\n[2/2] スライド生成中...")
    prs = new_presentation()

    print("  Slide 1: 表紙")
    slide_01_title(prs)
    print("  Slide 2: 意思決定事項")
    slide_02_decisions(prs)
    print("  Slide 3: エグゼクティブサマリー")
    slide_03_exec_summary(prs)
    print("  Slide 4: セクション - 外部環境")
    make_section_divider(prs, "01", "外部環境分析", "AIエージェント時代のSaaS構造変化", 4)
    print("  Slide 5: SaaSpocalypse")
    slide_05_saaspocalypse(prs)
    print("  Slide 6: 日本市場")
    slide_06_japan_market(prs)
    print("  Slide 7: セクション - 自社現状")
    make_section_divider(prs, "02", "自社現状分析", "定量データが示す「安定基盤と成長の踊り場」", 7)
    print("  Slide 8: 売上推移")
    slide_08_revenue(prs)
    print("  Slide 9: SaaS KPI")
    slide_09_kpi(prs)
    print("  Slide 10: SWOT")
    slide_10_swot(prs)
    print("  Slide 11: クロスSWOT")
    slide_11_cross_swot(prs)
    print("  Slide 12: 競合比較")
    slide_12_competitor_table(prs)
    print("  Slide 13: ポジショニングマップ1")
    slide_13_pos_map1(prs)
    print("  Slide 14: ポジショニングマップ2")
    slide_14_pos_map2(prs)
    print("  Slide 15: セクション - 重要課題")
    make_section_divider(prs, "03", "重要課題のMECE整理", "6視点 × 緊急度 × インパクト × 難易度", 15)
    print("  Slide 16: MECE課題マップ")
    slide_16_mece(prs)
    print("  Slide 17: 優先度マトリクス")
    slide_17_priority(prs)
    print("  Slide 18: セクション - 戦略オプション")
    make_section_divider(prs, "04", "戦略オプション比較", "3つの道 — 守るか、備えるか、攻めるか", 18)
    print("  Slide 19: 3つの戦略オプション")
    slide_19_options(prs)
    print("  Slide 20: 推奨戦略")
    slide_20_recommended(prs)
    print("  Slide 21: Gate Review基準")
    slide_21_gate(prs)
    print("  Slide 22: 重点施策 - プロダクト")
    slide_22_product(prs)
    print("  Slide 23: 重点施策 - 営業×CS")
    slide_23_sales_cs(prs)
    print("  Slide 24: 重点施策 - 組織")
    slide_24_org(prs)
    print("  Slide 25: ロードマップ")
    slide_25_roadmap(prs)
    print("  Slide 26: KPIツリー")
    slide_26_kpi_tree(prs)
    print("  Slide 27: 意思決定アジェンダ")
    slide_27_decision_closing(prs)
    print("  Slide 28: クロージング")
    slide_28_closing(prs)
    print("  Slide 29-35: Appendix")
    slide_29_revenue_detail(prs)
    slide_30_churn(prs)
    slide_31_market_data(prs)
    slide_32_competitor_profiles(prs)
    slide_33_on24(prs)
    slide_34_qa(prs)  # creates 2 slides (34+35)

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
