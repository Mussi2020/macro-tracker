# -*- coding: utf-8 -*-
"""
宏观指标跟踪 · 自动更新脚本 v3（宏观决策支持系统）
================================================
功能：
  1. 从 Yahoo Finance 抓取 13 个宏观指标全量历史数据（最早至1976年）：
     美债10Y/2Y、布伦特、WTI、黄金、铜、美元指数、美元兑人民币、美元兑日元、
     VIX、标普500、纳斯达克、高收益债ETF(HYG)；
     其中 US2Y/Brent/WTI 另用 FRED（圣路易斯联储）补长历史至 1976/1987/1986 年
  2. 数据质量层（P0）：单日异常变动检测与剔除（双向验证防误杀）
  3. 计算派生指标：10Y-2Y 期限利差、Brent-WTI 价差、金油比、美元动量、信用状态
  4. 宏观因子引擎（P1）：Growth/Inflation/Liquidity/Risk 四因子评分 + Regime 象限
     （Reflation/Goldilocks/Stagflation/Deflation）
  5. Z-Score + 历史百分位（P2）、30/90/365日收益率相关性（P4）、
     宏观分歧检测（P5）、Regime 历史（P6）、资产倾向（P7）
  6. 规则引擎生成「宏观状态」5 维度 + 自动宏观解读
  7. 合并进「宏观指标跟踪.xlsx」：历史数据 / 最新看板 / 趋势图
  8. 生成多层结构 HTML 决策仪表盘
  9. 每次写入前自动备份到 backups/（保留最近 10 份）

用法：
  手动运行：  python update_macro.py
  定时运行：  由 Windows 任务计划程序 "MacroTracker_DailyUpdate" 每日 09:00 调用
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import datetime as dt
import calendar
import traceback
from pathlib import Path

# 依赖装在脚本同级的 .pylibs 里（迁移到新机器后无需重建 venv）
_PYLIBS = Path(__file__).resolve().parent / ".pylibs"
if _PYLIBS.is_dir():
    sys.path.insert(0, str(_PYLIBS))

import requests
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------------
# 跟随脚本自身位置，换电脑/换盘无需改路径
BASE = Path(__file__).resolve().parent
XLSX = BASE / "宏观指标跟踪.xlsx"
HTML = BASE / "宏观趋势图.html"
BACKUP_DIR = BASE / "backups"
LOG = BASE / "update.log"
BACKUP_KEEP = 10          # 保留备份份数
FRED_SKIP_DAYS = 9000     # 主源历史已超过该天数时，跳过 FRED 补历史（约 1990 年）

# 前期摆动高低点（ZigZag 反转确认阈值）。不同指标波动率差异大，按指标单独设：
# 阈值越小越灵敏（摆动点越多、越贴近短期），越大越只认大级别转折。
SWING_PCT_DEFAULT = 0.05
SWING_PCT = {
    "US10Y": 0.10, "US2Y": 0.10,          # 利率：百分点级别的摆动才算数
    "Brent": 0.10, "WTI": 0.10,           # 原油
    "Gold": 0.08, "Copper": 0.10,         # 商品
    "DXY": 0.03, "USDCNY": 0.02, "USDJPY": 0.03,   # 汇率：波动小
    "VIX": 0.25,                          # 恐慌指数本身波动极大
    "SPX": 0.05, "NDX": 0.06,             # 股指：5%~6% 回调/反弹算一次摆动
    "HYG": 0.03,                          # 信用债 ETF
    # 月度宏观/估值：变化平缓，阈值相应调小
    "CPI": 0.02, "PCE": 0.02, "PPI": 0.05,
    "NFP": 0.01, "PMI": 0.40, "CAPE": 0.10,
}
HIST_KEEP = 12000         # 历史数据最大行数（1990年起约36年 ≈ 9300+ 交易日）
FETCH_RANGE = "max"       # Yahoo 抓取范围：全量历史（各指标自上市/可得日起，最早1990年前）
HIST_SHEET = "历史数据"
SNAP_SHEET = "最新看板"
CHART_SHEET = "趋势图"

# Cloudflare Pages 自动发布（存在 cf_config.json 才启用；发布失败不影响本地更新）
CF_CONFIG = BASE / "cf_config.json"
_HOME = Path.home()


def _find_node() -> Path:
    """按顺序找第一个存在的 node.exe（WorkBuddy 托管版 → 系统版）。"""
    cands = [p for p in (_HOME / ".workbuddy" / "binaries" / "node"
                         / "versions").glob("*/node.exe")]
    cands += [p for p in (_HOME / ".workbuddy" / "binaries" / "node"
                          / "versions").glob("*/node")]
    cands += [Path(r"C:\Program Files\nodejs\node.exe")]
    for p in cands:
        if p.exists():
            return p
    return cands[-1]


WRANGLER_NODE = _find_node()
WRANGLER_JS = _HOME / ".workbuddy" / "binaries" / "node" / "workspace" \
    / "node_modules" / "wrangler" / "bin" / "wrangler.js"

# (历史表列名, Yahoo代码, 看板指标名, 默认备注, 显示精度, 图表颜色, 分组)
INDICATORS = [
    ("US10Y",  "^TNX",     "美国10年期国债收益率", "全球资产定价之锚；高位压制估值", 3, "#C00000", "利率"),
    ("US2Y",   "2YY=F",    "美国2年期国债收益率",  "更贴近美联储政策利率预期",       3, "#E8985E", "利率"),
    ("Brent",  "BZ=F",     "布伦特原油",           "全球通胀与地缘温度计",           2, "#843C0C", "商品"),
    ("WTI",    "CL=F",     "WTI原油",              "美国油价基准",                   2, "#BF8F00", "商品"),
    ("Gold",   "GC=F",     "黄金",                 "避险+抗通胀；与实际利率反向",    1, "#806000", "商品"),
    ("Copper", "HG=F",     "COMEX铜",              "全球经济领先指标；铜涨=需求回暖", 1, "#B87333", "商品"),
    ("DXY",    "DX-Y.NYB", "美元指数 DXY",         "美元强弱；强→新兴市场承压",      2, "#375623", "美元"),
    ("USDCNY", "USDCNY=X", "美元兑人民币",         "影响海外资产换算与资金流向",     4, "#1F4E79", "美元"),
    ("USDJPY", "JPY=X",    "美元兑日元",           "套息交易风向标；急升破160有风险", 3, "#2E75B6", "美元"),
    ("VIX",    "^VIX",     "VIX恐慌指数",          "<15平静 20-30紧张 >30恐慌",      2, "#7030A0", "风险"),
    ("SPX",    "^GSPC",    "标普500",              "全球风险资产风向标",             1, "#E36C09", "风险"),
    ("NDX",    "^IXIC",    "纳斯达克综合",         "成长股风险偏好",                 1, "#6C3FA0", "风险"),
    ("HYG",    "HYG",      "高收益债ETF(HYG)",     "信用利差代理；大跌=信用压力",    2, "#548235", "风险"),
    # ↓ 以下为月度宏观基本面/估值，Yahoo 无对应序列，直接走 FRED / Shiller 源
    ("CPI",    "",         "美国CPI(城市消费者)",  "通胀总口径；美联储目标 2%",      2, "#C55A11", "宏观"),
    ("PCE",    "",         "美国PCE物价指数",      "美联储决策最看重的通胀口径",     2, "#ED7D31", "宏观"),
    ("PPI",    "",         "美国PPI(全部商品)",    "上游通胀，通常领先CPI 1-3个月",  2, "#FFC000", "宏观"),
    ("NFP",    "",         "美国非农就业人数",     "就业总规模（千人）；月度新增见备注", 0, "#2E75B6", "宏观"),
    ("PMI",    "",         "制造业景气指数(纽约联储)", "ISM PMI 无免费历史源，改用纽约联储制造业指数；0上方=扩张", 1, "#7030A0", "宏观"),
    ("CAPE",   "",         "席勒CAPE(周期调整PE)", "标普10年周期调整市盈率；>30 偏贵", 2, "#833C00", "估值"),
]

# 美国宏观基本面月度序列（FRED，无需 API Key），Yahoo 无对应数据
MACRO_FRED = {
    "CPI": "CPIAUCSL",            # CPI 城市消费者季调，1947 起
    "PCE": "PCEPI",               # PCE 物价指数，1959 起
    "PPI": "PPIACO",              # PPI 全部商品，1913 起
    "NFP": "PAYEMS",              # 非农就业人数（千人），1939 起
    "PMI": "GACDISA066MSFRBNY",   # 纽约联储 Empire State 制造业指数，1967 起
}
# 月度指标：涨跌幅按「数据点个数」比较，而非自然日
MONTHLY_COLS = set(MACRO_FRED) | {"CAPE"}

# 席勒 CAPE：耶鲁官方 ie_data.xls（已停更于 2023-09），之后的数据用
# cape_manual.csv 手工补充，每行格式：2023-10,31.2
CAPE_URL = "http://www.econ.yale.edu/~shiller/data/ie_data.xls"
CAPE_MANUAL = BASE / "cape_manual.csv"

# 数据质量层（P0）：
# 1) VALID_RANGE 各指标数值合理区间（硬校验，主要拦截列错位级别的脏数据）
# 2) SPIKE_PCT 单日双向 V 型跳变阈值（%）：与前一日 AND 后一日相比都超阈值才剔除，
#    阈值设得很宽（60%+），只拦截不可能的真实行情，保留历史真实暴跌暴涨
VALID_RANGE = {
    "US10Y": (0, 25), "US2Y": (0, 20), "Brent": (0, 250), "WTI": (-100, 250),
    "Gold": (100, 6000), "Copper": (0.3, 12), "DXY": (70, 165),
    "USDCNY": (5, 10), "USDJPY": (60, 250), "VIX": (5, 150),
    "SPX": (50, 20000), "NDX": (50, 40000), "HYG": (20, 120),
    # 宏观基本面（月度）与估值
    "CPI": (10, 1000), "PCE": (5, 1000), "PPI": (5, 3000),
    "NFP": (20000, 300000), "PMI": (-60, 60), "CAPE": (4, 60),
}
SPIKE_PCT = 60.0

GROUP_ORDER = ["利率", "宏观", "估值", "风险", "商品", "美元"]
GROUP_ICON = {"利率": "📉", "宏观": "🏛", "估值": "⚖️",
              "风险": "📊", "商品": "🛢", "美元": "💵"}

# HTML 分组键（语言无关），JS 端再按语言显示
GKEY = {"利率": "rates", "宏观": "macro", "估值": "val",
        "风险": "risk", "商品": "cmd", "美元": "fx"}
GKEY_ORDER = ["rates", "macro", "val", "risk", "cmd", "fx"]

# ----------------------------------------------------------------------------
# 多语言支持（i18n）：默认英语，共 8 大语言
# TR[key] = (en, zh, es, fr, de, ja, ru, pt)；占位符 {name}，用 str.format 渲染；
# 参数值以 "@" 开头表示嵌套翻译键（如 "@d4_lead_a"），渲染时再查表。
# ----------------------------------------------------------------------------
LANGS = ("en", "zh", "es", "fr", "de", "ja", "ru", "pt")
LANG_NAMES = {"en": "English", "zh": "简体中文", "es": "Español", "fr": "Français",
              "de": "Deutsch", "ja": "日本語", "ru": "Русский", "pt": "Português"}
DEFAULT_LANG = "en"

TR = {
    # ---- 界面文字（ui_ 前缀 → JS 端 UI 字典）----
    "ui_title": ("Macro Decision Dashboard", "宏观决策仪表盘",
                 "Panel de Decisión Macro", "Tableau de Bord Macro",
                 "Makro-Entscheidungs-Dashboard", "マクロ・ディシジョン・ダッシュボード",
                 "Макроаналитическая панель решений", "Painel de Decisão Macro"),
    "ui_dstat_label": ("Data status:", "数据状态：", "Estado de los datos:",
                       "État des données :", "Datenstatus:", "データ状態：",
                       "Состояние данных:", "Status dos dados:"),
    "ui_last_update": ("Last updated:", "最后更新：", "Última actualización:",
                       "Dernière mise à jour :", "Letzte Aktualisierung:", "最終更新：",
                       "Последнее обновление:", "Última actualização:"),
    "ui_data_source": ("Data source:", "数据源：", "Fuente de datos:",
                       "Source des données :", "Datenquelle:", "データソース：",
                       "Источник данных:", "Fonte dos dados:"),
    "ui_color_note": ("Red = up, Green = down", "涨红跌绿",
                      "Rojo = subida, Verde = bajada", "Rouge = hausse, Vert = baisse",
                      "Rot = steigend, Grün = fallend", "赤=上昇、緑=下落",
                      "Красный = рост, Зелёный = падение", "Vermelho = alta, Verde = baixa"),
    "ui_lang_label": ("Language", "语言", "Idioma", "Langue", "Sprache", "言語",
                      "Язык", "Idioma"),
    "ui_eval_label": ("Evaluation start date", "评估起始日",
                      "Fecha inicial de evaluación", "Date de début d'évaluation",
                      "Bewertungsstartdatum", "評価開始日",
                      "Начальная дата оценки", "Data inicial de avaliação"),
    "ui_reset": ("Reset", "重置", "Restablecer", "Réinitialiser", "Zurücksetzen",
                 "リセット", "Сброс", "Redefinir"),
    "ui_reset_tip": ("Restore default (last 12 months)", "恢复默认（近12个月）",
                     "Restaurar el valor predeterminado (últimos 12 meses)",
                     "Rétablir la valeur par défaut (12 derniers mois)",
                     "Standard wiederherstellen (letzte 12 Monate)",
                     "既定値に戻す（直近12か月）",
                     "Вернуть значение по умолчанию (последние 12 месяцев)",
                     "Restaurar padrão (últimos 12 meses)"),
    "ui_eval_dyn_tip": ("Defaults to the trailing 12 months; editable",
                        "默认为最近 12 个月，可自行修改",
                        "Por defecto, los últimos 12 meses; editable",
                        "Par défaut, les 12 derniers mois ; modifiable",
                        "Standard: die letzten 12 Monate; anpassbar",
                        "既定値は直近12か月。変更可能",
                        "По умолчанию — последние 12 месяцев; можно изменить",
                        "Padrão: últimos 12 meses; editável"),
    "ui_swing_tip": ("H = highest in {r}, L = lowest in {r}; highlighted = whichever came later",
                     "H＝{r}内最高点　L＝{r}内最低点，黄底＝更晚出现的那个",
                     "H = máximo en {r}, L = mínimo en {r}; resaltado = el más reciente",
                     "H = plus haut de {r}, L = plus bas de {r}; surligné = le plus récent",
                     "H = Höchstwert in {r}, L = Tiefstwert in {r}; hervorgehoben = der Spätere",
                     "H＝{r}の最高値　L＝{r}の最安値、黄背景＝より後に出た方",
                     "H = максимум за {r}, L = минимум за {r}; выделено = более поздний",
                     "H = máxima em {r}, L = mínima em {r}; destacado = o mais recente"),
    "ui_sec_regime": ("🌎 Current Macro State", "🌎 当前宏观状态",
                      "🌎 Estado Macro Actual", "🌎 État Macro Actuel",
                      "🌎 Aktueller Makro-Zustand", "🌎 現在のマクロ状態",
                      "🌎 Текущее макросостояние", "🌎 Estado Macro Atual"),
    "ui_sec_diverge": ("⚠️ Macro Divergence Alerts", "⚠️ 宏观分歧检测",
                       "⚠️ Alertas de Divergencia Macro", "⚠️ Alertes de Divergence Macro",
                       "⚠️ Makro-Divergenz-Warnungen", "⚠️ マクロ・ダイバージェンス検出",
                       "⚠️ Оповещения о макрорасхождениях", "⚠️ Alertas de Divergência Macro"),
    "ui_sec_derived": ("🧮 Derived Indicators", "🧮 派生指标",
                       "🧮 Indicadores Derivados", "🧮 Indicateurs Dérivés",
                       "🧮 Abgeleitete Indikatoren", "🧮 派生指標",
                       "🧮 Производные индикаторы", "🧮 Indicadores Derivados"),
    "ui_sec_zscore": ("📐 Momentum Z-Score & Historical Percentile",
                      "📐 动量 Z-Score 与历史百分位",
                      "📐 Z-Score de Momento y Percentil Histórico",
                      "📐 Z-Score de Momentum et Percentile Historique",
                      "📐 Momentum Z-Score & historisches Perzentil",
                      "📐 モメンタム Z-Score と履歴パーセンタイル",
                      "📐 Z-оценка импульса и исторический перцентиль",
                      "📐 Z-Score de Momentum e Percentil Histórico"),
    "ui_zscore_hint": (
        "Z-Score: standardized 3-month momentum vs. its full-history distribution "
        "(|Z|>2 = historically rare). Percentile: where the current level sits "
        "within full history.",
        "Z-Score：近3月动量相对全历史分布的标准化得分（|Z|>2 = 历史罕见）；"
        "百分位：当前价格水平在全历史中的位置。",
        "Z-Score: momentum de 3 meses estandarizado frente a su distribución "
        "histórica completa (|Z|>2 = históricamente raro); Percentil: posición "
        "del nivel actual en la historia completa.",
        "Z-Score : momentum sur 3 mois standardisé par rapport à sa distribution "
        "historique complète (|Z|>2 = rare historiquement) ; Percentile : "
        "position du niveau actuel dans l'historique complet.",
        "Z-Score: 3-Monats-Momentum, standardisiert gegenüber der Gesamthistorie "
        "(|Z|>2 = historisch selten); Perzentil: Position des aktuellen Niveaus "
        "in der Gesamthistorie.",
        "Z-Score：過去全期間の分布に対する直近3ヶ月モメンタムの標準化スコア"
        "（|Z|>2 は歴史的に珍しい）。パーセンタイル：現在水準の全期間中での位置。",
        "Z-оценка: стандартизированный 3-месячный импульс относительно всей "
        "истории (|Z|>2 — редкость); перцентиль: положение текущего уровня во "
        "всей истории.",
        "Z-Score: momentum de 3 meses padronizado frente à sua distribuição "
        "histórica completa (|Z|>2 = historicamente raro); Percentil: posição do "
        "nível atual na história completa."),
    "ui_sec_corr": ("🔗 Asset Return Correlation Matrix", "🔗 资产收益率相关性矩阵",
                    "🔗 Matriz de Correlación de Rentabilidades",
                    "🔗 Matrice de Corrélation des Rendements",
                    "🔗 Korrelationsmatrix der Anlageerträge", "🔗 資産リターン相関行列",
                    "🔗 Матрица корреляции доходностей активов",
                    "🔗 Matriz de Correlação de Retornos"),
    "ui_corr_hint": (
        "Red = positive, green = negative correlation; computed on daily "
        "close-to-close returns. Darker = stronger relationship.",
        "红=正相关，绿=负相关；基于收盘价日收益率（%变化）计算。深色=关系更强。",
        "Rojo = correlación positiva, verde = negativa; calculado con rendimientos "
        "diarios de cierre. Más oscuro = relación más fuerte.",
        "Rouge = corrélation positive, vert = négative ; calculé sur les "
        "rendements quotidiens de clôture. Plus foncé = relation plus forte.",
        "Rot = positive, Grün = negative Korrelation; berechnet auf "
        "Tagesschlussrenditen. Dunkler = stärkere Beziehung.",
        "赤=正の相関、緑=負の相関。終値ベースの日次リターンで計算。"
        "濃い色ほど関係が強い。",
        "Красный = положительная, зелёный = отрицательная корреляция; рассчитано "
        "по дневным доходностям закрытия. Темнее — сильнее связь.",
        "Vermelho = correlação positiva, verde = negativa; calculado sobre "
        "retornos diários de fechamento. Mais escuro = relação mais forte."),
    "ui_sec_hist": ("🕰 Macro Regime History (Last 6 Months)", "🕰 宏观状态历史（近半年）",
                    "🕰 Historial del Régimen Macro (últimos 6 meses)",
                    "🕰 Historique du Régime Macro (6 derniers mois)",
                    "🕰 Makro-Regime-Historie (letzte 6 Monate)",
                    "🕰 マクロ・レジーム履歴（直近6ヶ月）",
                    "🕰 История макрорежима (последние 6 месяцев)",
                    "🕰 Histórico do Regime Macro (últimos 6 meses)"),
    "ui_sec_bias": ("⚖️ Asset Bias (Regime Bias)", "⚖️ 资产倾向（Regime Bias）",
                    "⚖️ Sesgo de Activos (Regime Bias)", "⚖️ Biais d'Actifs (Regime Bias)",
                    "⚖️ Anlage-Tendenzen (Regime Bias)", "⚖️ 資産バイアス（レジーム・バイアス）",
                    "⚖️ Перекос активов (Regime Bias)", "⚖️ Viés de Ativos (Regime Bias)"),
    "ui_bias_hint": ("Directional tilts based on current macro factors — not "
                     "trading signals.",
                     "基于当前宏观因子的方向性倾向参考，非买卖信号。",
                     "Inclinaciones direccionales según los factores macro "
                     "actuales; no son señales de trading.",
                     "Inclinaisons directionnelles basées sur les facteurs macro "
                     "actuels — il ne s'agit pas de signaux de trading.",
                     "Direktionale Neigungen basierend auf aktuellen "
                     "Makrofaktoren — keine Handelssignale.",
                     "現在のマクロファクターに基づく方向的な傾向。売買シグナルではない。",
                     "Направленные перекосы на основе текущих макрофакторов — "
                     "не торговые сигналы.",
                     "Inclinações direcionais com base nos fatores macro atuais "
                     "— não são sinais de negociação."),
    "ui_sec_insight": ("🧠 Today's Macro Read", "🧠 今日宏观解读",
                       "🧠 Lectura Macro del Día", "🧠 Lecture Macro du Jour",
                       "🧠 Makro-Lage des Tages", "🧠 本日のマクロ解説",
                       "🧠 Макрообзор дня", "🧠 Leitura Macro do Dia"),
    "ui_live": ("LIVE", "LIVE 实时", "LIVE en vivo", "LIVE en direct", "LIVE",
                "LIVE リアルタイム", "LIVE (онлайн)", "LIVE ao vivo"),
    "ui_partial": ("PARTIAL ({t})", "PARTIAL 部分缓存（{t}）", "PARCIAL ({t})",
                   "PARTIEL ({t})", "TEILWEISE ({t})", "部分的（{t}）",
                   "ЧАСТИЧНО ({t})", "PARCIAL ({t})"),
    "ui_cache": ("CACHED", "CACHE 缓存数据", "CACHE (datos en caché)",
                 "CACHE (données en cache)", "CACHE (zwischengespeichert)",
                 "CACHE（キャッシュ）", "CACHE (кэш)", "CACHE (em cache)"),
    "ui_chg1m": ("1M", "近1月", "1M", "1M", "1M", "1ヶ月", "1 мес.", "1M"),
    "ui_chg3m": ("3M", "近3月", "3M", "3M", "3M", "3ヶ月", "3 мес.", "3M"),
    "ui_since": ("since {ym}", "自{ym}", "desde {ym}", "depuis {ym}", "seit {ym}",
                 "{ym}以降", "с {ym}", "desde {ym}"),
    "ui_since_star": ("since {ym}*", "自{ym}*", "desde {ym}*", "depuis {ym}*",
                      "seit {ym}*", "{ym}以降*", "с {ym}*", "desde {ym}*"),
    "ui_since_tip": ("This series starts in {ym}; change is measured from that base",
                     "该指标数据自 {ym} 起，涨跌幅以此为基准",
                     "Esta serie comienza en {ym}; la variación se mide desde esa base",
                     "Cette série démarre en {ym} ; la variation est mesurée depuis cette base",
                     "Diese Serie beginnt im {ym}; die Veränderung wird ab dieser Basis gemessen",
                     "この系列は{ym}開始。騰落率はこの基準から算出",
                     "Данный ряд начинается с {ym}; изменение считается от этой базы",
                     "Esta série começa em {ym}; a variação é medida a partir dessa base"),
    "ui_eval_tip": ("Follow the global evaluation start date (currently {d})",
                    "跟随顶部评估起始日（当前 {d}）",
                    "Seguir la fecha inicial de evaluación global (actualmente {d})",
                    "Suivre la date de début d'évaluation globale (actuellement {d})",
                    "Dem globalen Bewertungsstartdatum folgen (aktuell {d})",
                    "上部の評価開始日に従う（現在 {d}）",
                    "Следовать глобальной начальной дате оценки (сейчас {d})",
                    "Seguir a data inicial de avaliação global (atualmente {d})"),
    "ui_v_prefix": ("Overall: ", "综合判断：", "Síntesis: ", "Synthèse : ",
                    "Gesamturteil: ", "総合判断：", "Общий вывод: ", "Síntese: "),
    "ui_sum_title": ("Overall", "综合判断", "Síntesis", "Synthèse", "Gesamturteil",
                     "総合判断", "Общий вывод", "Síntese"),
    "ui_f_k": ("{n} Factor", "{n}因子", "Factor {n}", "Facteur {n}", "Faktor {n}",
               "{n}ファクター", "Фактор: {n}", "Fator {n}"),
    "ui_f_score": ("score {s}σ", "评分 {s}σ", "puntuación {s}σ", "score {s}σ",
                   "Score {s}σ", "スコア {s}σ", "оценка {s}σ", "pontuação {s}σ"),
    "ui_z_pctile": ("pct {p}%", "百分位 {p}%", "pctil {p}%", "pctile {p}%",
                    "Perz. {p}%", "パーセンタイル {p}%", "перц. {p}%",
                    "percentil {p}%"),
    "ui_th_month": ("Month", "月份", "Mes", "Mois", "Monat", "月", "Месяц", "Mês"),
    "ui_th_growth": ("Growth", "增长", "Crecimiento", "Croissance", "Wachstum",
                     "成長", "Рост", "Crescimento"),
    "ui_th_infl": ("Inflation", "通胀", "Inflación", "Inflation", "Inflation",
                   "インフレ", "Инфляция", "Inflação"),
    "ui_th_asset": ("Asset", "资产", "Activo", "Actif", "Anlage", "資産",
                    "Актив", "Ativo"),
    "ui_th_bias": ("Bias", "倾向", "Sesgo", "Biais", "Tendenz", "バイアス",
                   "Перекос", "Viés"),
    "ui_th_note": ("Note", "说明", "Nota", "Note", "Hinweis", "説明",
                   "Примечание", "Nota"),
    "ui_t30": ("30D", "30日", "30D", "30J", "30T", "30日", "30 дн.", "30D"),
    "ui_t90": ("90D", "90日", "90D", "90J", "90T", "90日", "90 дн.", "90D"),
    "ui_t1y": ("1Y", "1年", "1A", "1A", "1J", "1年", "1 г.", "1A"),
    "ui_r_eval": ("Start", "起始日", "Inicio", "Début", "Start", "起点",
                  "Старт", "Início"),
    "ui_r_w1": ("1W", "1周", "1S", "1S", "1W", "1週", "1 нед.", "1S"),
    "ui_r_m1": ("1M", "1月", "1M", "1M", "1M", "1ヶ月", "1 мес.", "1M"),
    "ui_r_q1": ("3M", "季度", "3M", "3M", "3M", "四半期", "3 мес.", "3M"),
    "ui_r_h1": ("6M", "半年", "6M", "6M", "6M", "半年", "6 мес.", "6M"),
    "ui_r_y1": ("1Y", "1年", "1A", "1A", "1J", "1年", "1 г.", "1A"),
    "ui_r_y3": ("3Y", "3年", "3A", "3A", "3J", "3年", "3 г.", "3A"),
    "ui_r_y5": ("5Y", "5年", "5A", "5A", "5J", "5年", "5 лет", "5A"),
    "ui_r_y10": ("10Y", "10年", "10A", "10A", "10J", "10年", "10 лет", "10A"),
    "ui_r_all": ("All", "全部", "Todo", "Tout", "Alle", "全期間", "Всё", "Tudo"),
    "ui_yoy": ("YoY", "同比", "Interanual", "En glissement annuel",
               "Gegenüber Vorjahr", "前年比", "Год к году", "12 meses"),
    "ui_mom": ("MoM", "环比", "Mensual", "Mensuel", "Gegenüber Vormonat",
               "前月比", "Месяц к месяцу", "Mensal"),
    "ui_g_rates": ("📉 Rates", "📉 利率", "📉 Tasas", "📉 Taux", "📉 Zinsen",
                   "📉 金利", "📉 Ставки", "📉 Taxas"),
    "ui_g_macro": ("🏛 US Macro", "🏛 美国宏观", "🏛 Macro EE.UU.",
                   "🏛 Macro É.-U.", "🏛 US-Makro", "🏛 米国マクロ",
                   "🏛 Макро США", "🏛 Macro dos EUA"),
    "ui_g_val": ("⚖️ Valuation", "⚖️ 估值", "⚖️ Valoración",
                 "⚖️ Valorisation", "⚖️ Bewertung", "⚖️ バリュエーション",
                 "⚖️ Оценка", "⚖️ Valuation"),
    "ui_g_fx": ("💵 USD & FX", "💵 美元", "💵 Dólar y FX", "💵 Dollar & FX",
                "💵 USD & Devisen", "💵 米ドル・為替", "💵 Доллар и валюты",
                "💵 Dólar e Câmbio"),
    "ui_g_risk": ("📊 Risk Assets", "📊 风险资产", "📊 Activos de Riesgo",
                  "📊 Actifs Risqués", "📊 Risiko-Assets", "📊 リスク資産",
                  "📊 Рисковые активы", "📊 Ativos de Risco"),
    "ui_g_cmd": ("🛢 Commodities", "🛢 商品", "🛢 Materias Primas",
                 "🛢 Matières Premières", "🛢 Rohstoffe", "🛢 コモディティ",
                 "🛢 Сырьё", "🛢 Commodities"),
}


TR.update({
    # ---- 宏观状态 5 维度 ----
    "lbl_risk": ("Risk Appetite", "风险偏好", "Apetito de Riesgo",
                 "Appétit pour le Risque", "Risikoappetit", "リスク選好",
                 "Аппетит к риску", "Apetito de Risco"),
    "lbl_infl": ("Inflation Pressure", "通胀压力", "Presión Inflacionaria",
                 "Pression Inflationniste", "Inflationsdruck", "インフレ圧力",
                 "Инфляционное давление", "Pressão Inflacionária"),
    "lbl_rates": ("Rate Pressure", "利率压力", "Presión de Tasas",
                  "Pression des Taux", "Zinsdruck", "金利圧力",
                  "Давление ставок", "Pressão das Taxas"),
    "lbl_usd": ("US Dollar", "美元", "Dólar EE.UU.", "Dollar US", "US-Dollar",
                "米ドル", "Доллар США", "Dólar dos EUA"),
    "lbl_vol": ("Volatility", "波动率", "Volatilidad", "Volatilité", "Volatilität",
                "ボラティリティ", "Волатильность", "Volatilidade"),
    "risk_on": ("Risk appetite strong", "风险偏好偏强", "Apetito de riesgo fuerte",
                "Appétit pour le risque fort", "Risikoappetit stark",
                "リスク選好は強め", "Аппетит к риску силён", "Apetito de risco forte"),
    "risk_off": ("Risk appetite weak", "风险偏好偏弱", "Apetito de riesgo débil",
                 "Appétit pour le risque faible", "Risikoappetit schwach",
                 "リスク選好は弱め", "Аппетит к риску слаб", "Apetito de risco fraco"),
    "risk_neu": ("Risk appetite neutral", "风险偏好中性", "Apetito de riesgo neutral",
                 "Appétit pour le risque neutre", "Risikoappetit neutral",
                 "リスク選好は中立", "Аппетит к риску нейтрален", "Apetito de risco neutro"),
    "infl_up": ("Energy prices rising markedly, inflation pressure building",
                "能源价格明显走高，通胀压力升温",
                "Los precios de la energía suben notablemente, la presión "
                "inflacionaria aumenta",
                "Les prix de l'énergie montent nettement, la pression "
                "inflationniste s'intensifie",
                "Energiepreise steigen deutlich, Inflationsdruck nimmt zu",
                "エネルギー価格が明確に上昇し、インフレ圧力が高まっている",
                "Цены на энергоносители заметно растут, инфляционное давление "
                "усиливается",
                "Preços de energia subindo acentuadamente, pressão inflacionária "
                "aumentando"),
    "infl_ease": ("Energy and gold pulling back, inflation pressure easing",
                  "能源与黄金回落，通胀压力缓和",
                  "Energía y oro retroceden, la presión inflacionaria se modera",
                  "Énergie et or reculent, la pression inflationniste s'apaise",
                  "Energie und Gold geben nach, Inflationsdruck lässt nach",
                  "エネルギーと金が下落し、インフレ圧力が和らいでいる",
                  "Энергоносители и золото откатываются, инфляционное давление "
                  "ослабевает",
                  "Energia e ouro recuando, pressão inflacionária amenizando"),
    "infl_mid": ("Inflation pressure moderate, not the dominant driver",
                 "通胀压力温和，暂非主导变量",
                 "Presión inflacionaria moderada, no es la variable dominante",
                 "Pression inflationniste modérée, pas la variable dominante",
                 "Inflationsdruck moderat, nicht die dominierende Variable",
                 "インフレ圧力は穏やかで、今のところ主導要因ではない",
                 "Инфляционное давление умеренное, пока не главный фактор",
                 "Pressão inflacionária moderada, não é a variável dominante"),
    "rates_high": ("10Y yield at {y}% is elevated, weighing on valuations",
                   "10Y 收益率 {y}% 处于高位，估值承压",
                   "El rendimiento del 10Y en {y}% es elevado, presiona las "
                   "valoraciones",
                   "Le rendement du 10Y à {y}% est élevé, il pèse sur les "
                   "valorisations",
                   "10Y-Rendite von {y}% ist hoch und belastet die Bewertungen",
                   "10Y利回り {y}% は高水準で、バリュエーションを圧迫",
                   "Доходность 10Y на уровне {y}% высока, давит на оценки активов",
                   "O rendimento do 10Y em {y}% está elevado, pressionando as "
                   "avaliações"),
    "rates_mh": ("10Y yield at {y}%, moderately high", "10Y 收益率 {y}%，中性偏高",
                 "Rendimiento del 10Y en {y}%, moderadamente alto",
                 "Rendement du 10Y à {y}%, modérément élevé",
                 "10Y-Rendite von {y}%, moderat hoch",
                 "10Y利回り {y}%、やや高め",
                 "Доходность 10Y на уровне {y}%, умеренно высокая",
                 "Rendimento do 10Y em {y}%, moderadamente alto"),
    "rates_ok": ("10Y yield at {y}%, rate environment friendly",
                 "10Y 收益率 {y}%，利率环境友好",
                 "Rendimiento del 10Y en {y}%, entorno de tasas favorable",
                 "Rendement du 10Y à {y}%, environnement de taux favorable",
                 "10Y-Rendite von {y}%, Zinsumfeld günstig",
                 "10Y利回り {y}%、金利環境は良好",
                 "Доходность 10Y на уровне {y}%, процентная среда благоприятна",
                 "Rendimento do 10Y em {y}%, ambiente de taxas favorável"),
    "na_txt": ("-", "-", "-", "-", "-", "-", "-", "-"),
    "usd_strong": ("USD strengthening, pressure on EM and commodities",
                   "美元走强，新兴市场与大宗承压",
                   "El dólar se fortalece, presiona a emergentes y materias primas",
                   "Le dollar se renforce, pression sur les émergents et les "
                   "matières premières",
                   "Der Dollar stärkt sich, Druck auf Schwellenländer und Rohstoffe",
                   "米ドルが強含み、新興市場とコモディティに下押し圧力",
                   "Доллар укрепляется, давление на развивающиеся рынки и сырьё",
                   "Dólar se fortalecendo, pressão sobre emergentes e commodities"),
    "usd_weak": ("USD weakening, supportive for risk assets and gold",
                 "美元走弱，利多风险资产与黄金",
                 "El dólar se debilita, favorece a los activos de riesgo y al oro",
                 "Le dollar s'affaiblit, favorable aux actifs risqués et à l'or",
                 "Der Dollar schwächt sich, stützt Risiko-Assets und Gold",
                 "米ドルが弱含み、リスク資産と金に追い風",
                 "Доллар слабеет, поддерживает рисковые активы и золото",
                 "Dólar se enfraquecendo, favorecendo ativos de risco e ouro"),
    "usd_mid": ("USD range-bound", "美元区间震荡", "Dólar en rango lateral",
                "Dollar dans un range", "Dollar seitwärts", "米ドルはレンジ相場",
                "Доллар движется в диапазоне", "Dólar em consolidação"),
    "vol_low": ("VIX {v}, market calm (beware complacency)",
                "VIX {v}，市场平静（需警惕自满）",
                "VIX {v}, mercado en calma (cuidado con la complacencia)",
                "VIX {v}, marché calme (méfiez-vous de la complaisance)",
                "VIX {v}, Markt ruhig (Vorsicht vor Selbstzufriedenheit)",
                "VIX {v}、市場は平静（慢心に注意）",
                "VIX {v}, рынок спокоен (осторожно с самоуспокоенностью)",
                "VIX {v}, mercado calmo (cuidado com a complacência)"),
    "vol_norm": ("VIX {v}, volatility normal", "VIX {v}，波动正常",
                 "VIX {v}, volatilidad normal", "VIX {v}, volatilité normale",
                 "VIX {v}, Volatilität normal", "VIX {v}、ボラティリティは正常",
                 "VIX {v}, волатильность нормальная", "VIX {v}, volatilidade normal"),
    "vol_tense": ("VIX {v}, markets nervous", "VIX {v}，市场情绪紧张",
                  "VIX {v}, nerviosismo en el mercado", "VIX {v}, marché nerveux",
                  "VIX {v}, nervöser Markt", "VIX {v}、市場心理は緊張",
                  "VIX {v}, рынок напряжен", "VIX {v}, mercado tenso"),
    "vol_panic": ("VIX {v}, panic dominates, watch for liquidity shocks",
                  "VIX {v}，恐慌情绪主导，谨防流动性冲击",
                  "VIX {v}, domina el pánico, cuidado con shocks de liquidez",
                  "VIX {v}, la panique domine, méfiez-vous des chocs de liquidité",
                  "VIX {v}, Panik dominiert, Vorsicht vor Liquiditätsschocks",
                  "VIX {v}、パニックが支配的、流動性ショックに注意",
                  "VIX {v}, царит паника, опасайтесь шоков ликвидности",
                  "VIX {v}, pânico dominante, cuidado com choques de liquidez"),
    # ---- 综合判断片段 ----
    "sb_risk_s": ("solid risk appetite", "风险偏好较强", "apetito de riesgo sólido",
                  "appétit pour le risque solide", "stabiler Risikoappetit",
                  "リスク選好は比較的強い", "устойчивый аппетит к риску",
                  "apetito de risco sólido"),
    "sb_risk_w": ("weak risk appetite", "风险偏好偏弱", "apetito de riesgo débil",
                  "appétit pour le risque faible", "schwacher Risikoappetit",
                  "リスク選好は弱め", "слабый аппетит к риску",
                  "apetito de risco fraco"),
    "sb_risk_n": ("neutral risk appetite", "风险偏好中性", "apetito de riesgo neutral",
                  "appétit pour le risque neutre", "neutraler Risikoappetit",
                  "リスク選好は中立", "нейтральный аппетит к риску",
                  "apetito de risco neutro"),
    "sb_rate_hi": ("rates staying high", "利率维持高位", "tasas se mantienen altas",
                   "les taux restent élevés", "Zinsen bleiben hoch",
                   "金利は高止まり", "ставки остаются высокими",
                   "taxas permanecem elevadas"),
    "sb_rate_ok": ("moderate rates", "利率适中", "tasas moderadas", "taux modérés",
                   "moderate Zinsen", "金利は適度", "умеренные ставки",
                   "taxas moderadas"),
    "sb_infl_up": ("energy-driven inflation pressure", "能源价格带来通胀压力",
                   "presión inflacionaria por energía",
                   "pression inflationniste tirée par l'énergie",
                   "energiegetriebener Inflationsdruck",
                   "エネルギー主導のインフレ圧力",
                   "инфляционное давление из-за энергоносителей",
                   "pressão inflacionária vinda da energia"),
    "sb_infl_ease": ("easing inflation pressure", "通胀压力缓和",
                     "presión inflacionaria en descenso",
                     "pression inflationniste en baisse", "nachlassender Inflationsdruck",
                     "インフレ圧力は和らいでいる", "ослабевающее инфляционное давление",
                     "pressão inflacionária em queda"),
    "sb_usd_weak": ("weaker USD", "美元走弱", "dólar más débil", "dollar plus faible",
                    "schwächerer Dollar", "米ドル安", "более слабый доллар",
                    "dólar mais fraco"),
    "sb_usd_strong": ("firmer USD", "美元偏强", "dólar más fuerte", "dollar plus ferme",
                      "stabilerer Dollar", "米ドル高", "более сильный доллар",
                      "dólar mais forte"),
    "sb_vol_low": ("low volatility", "波动率低位", "volatilidad baja",
                   "volatilité basse", "niedrige Volatilität", "ボラティリティは低水準",
                   "низкая волатильность", "volatilidade baixa"),
    "sum_prefix": ("Current market mix: ", "当前市场组合：",
                   "Mezcla actual de mercado: ", "Composition actuelle du marché : ",
                   "Aktuelle Marktkonstellation: ", "現在の市場コンビネーション：",
                   "Текущая рыночная комбинация: ", "Combinação atual de mercado: "),
    "sum_suffix": (".", "。", ".", ".", ".", "。", ".", "."),
})


TR.update({
    # ---- 今日解读：标题 ----
    "it_rates": ("Rates", "利率", "Tasas", "Taux", "Zinsen", "金利", "Ставки",
                 "Taxas"),
    "it_usd": ("USD", "美元", "Dólar", "Dollar", "USD", "米ドル", "Доллар", "Dólar"),
    "it_risk": ("Risk", "风险", "Riesgo", "Risque", "Risiko", "リスク", "Риск",
                "Risco"),
    "it_infl": ("Inflation", "通胀", "Inflación", "Inflation", "Inflation",
                "インフレ", "Инфляция", "Inflação"),
    "it_gold": ("Gold", "黄金", "Oro", "Or", "Gold", "金", "Золото", "Ouro"),
    # ---- 今日解读：利率 ----
    "ir_s1": ("US 10Y yield at {y}%, {c}% over the past month. ",
              "美国10Y收益率 {y}%，近1月 {c}%。",
              "Rendimiento del Tesoro 10Y EE.UU. en {y}%, {c}% en el último mes. ",
              "Rendement du Trésor américain 10Y à {y}%, {c}% sur le mois écoulé. ",
              "US-10Y-Rendite bei {y}%, {c}% im letzten Monat. ",
              "米10年金利 {y}%、直近1ヶ月 {c}%。",
              "Доходность 10-летних казначейских облигаций США {y}%, за месяц {c}%. ",
              "Rendimento do Tesouro 10Y dos EUA em {y}%, {c}% no último mês. "),
    "ir_s1b": ("US 10Y yield at {y}%. ", "美国10Y收益率 {y}%。",
               "Rendimiento del Tesoro 10Y EE.UU. en {y}%. ",
               "Rendement du Trésor américain 10Y à {y}%. ",
               "US-10Y-Rendite bei {y}%. ", "米10年金利 {y}%。",
               "Доходность 10-летних казначейских облигаций США {y}%. ",
               "Rendimento do Tesouro 10Y dos EUA em {y}%. "),
    "ir_hi": ("Yields remain elevated; markets stay sensitive to long-term "
              "funding costs, weighing on richly-valued growth stocks and gold.",
              "收益率维持高位，市场对长期资金价格仍敏感，对高估值成长股与黄金形成压制。",
              "Los rendimientos se mantienen elevados; el mercado sigue siendo "
              "sensible al costo del financiamiento a largo plazo, lo que "
              "presiona a las acciones de crecimiento con valoraciones altas y al oro.",
              "Les rendements restent élevés ; le marché demeure sensible au "
              "coût du financement à long terme, ce qui pèse sur les valeurs de "
              "croissance chèrement valorisées et sur l'or.",
              "Die Renditen bleiben hoch; der Markt bleibt sensibel für "
              "langfristige Finanzierungskosten, was hoch bewertete "
              "Wachstumsaktien und Gold belastet.",
              "利回りは高水準が続いており、長期資金コストへの感応度が高く、"
              "高バリュエーションのグロース株と金が圧迫されている。",
              "Доходности остаются высокими; рынок чувствителен к стоимости "
              "долгосрочного финансирования, что давит на дорого оцененные акции "
              "роста и золото.",
              "Os rendimentos permanecem elevados; o mercado segue sensível ao "
              "custo do financiamento de longo prazo, pressionando ações de "
              "crescimento com avaliações altas e o ouro."),
    "ir_mid": ("Yields are in a neutral range; the marginal rate pressure on "
               "assets is limited.",
               "收益率处于中性区间，利率对资产的边际压力有限。",
               "Los rendimientos están en un rango neutral; la presión marginal "
               "de las tasas sobre los activos es limitada.",
               "Les rendements sont dans une zone neutre ; la pression marginale "
               "des taux sur les actifs est limitée.",
               "Die Renditen liegen in einem neutralen Bereich; der marginale "
               "Zinsdruck auf Anlagen ist begrenzt.",
               "利回りは中立的なレンジにあり、資産への限界的な金利圧力は限定的。",
               "Доходности находятся в нейтральном диапазоне; предельное давление "
               "ставок на активы ограничено.",
               "Os rendimentos estão em uma faixa neutra; a pressão marginal das "
               "taxas sobre os ativos é limitada."),
    "ir_curve": ("10Y-2Y term spread at {c}%. ", "期限利差(10Y-2Y)为 {c}%，",
                 "El diferencial 10Y-2Y está en {c}%. ",
                 "L'écart 10Y-2Y est de {c}%. ",
                 "Die Zinsdifferenz 10Y-2Y beträgt {c}%. ",
                 "期間スプレッド（10Y-2Y）は {c}%。",
                 "Спред между 10Y и 2Y составляет {c}%. ",
                 "O spread 10Y-2Y está em {c}%. "),
    "ir_curve_inv": ("The curve remains inverted — keep watching the recession "
                     "signal.", "曲线倒挂仍需关注衰退信号。",
                     "La curva sigue invertida; vigilar la señal de recesión.",
                     "La courbe reste inversée ; surveiller le signal de récession.",
                     "Die Kurve bleibt invertiert — das Rezessionssignal im Blick "
                     "behalten.", "イールドカーブは逆転したままで、景気後退シグナルに注視が必要。",
                     "Кривая остаётся инвертированной — следите за сигналом рецессии.",
                     "A curva segue invertida; monitorar o sinal de recessão."),
    "ir_curve_ok": ("Curve shape is normal.", "曲线形态正常。",
                    "La forma de la curva es normal.", "La forme de la courbe est normale.",
                    "Die Kurvenform ist normal.", "カーブの形状は正常。",
                    "Форма кривой нормальная.", "O formato da curva é normal."),
    # ---- 今日解读：美元 ----
    "iu_s1": ("DXY at {d}, {c}% over the past month", "DXY 报 {d}，近1月 {c}%",
              "DXY en {d}, {c}% en el último mes", "DXY à {d}, {c}% sur le mois écoulé",
              "DXY bei {d}, {c}% im letzten Monat", "DXY {d}、直近1ヶ月 {c}%",
              "DXY на {d}, за месяц {c}%", "DXY em {d}, {c}% no último mês"),
    "iu_s1b": ("DXY at {d}", "DXY 报 {d}", "DXY en {d}", "DXY à {d}", "DXY bei {d}",
               "DXY {d}", "DXY на {d}", "DXY em {d}"),
    "iu_weak": (". The dollar has weakened recently, supporting gold and "
               "emerging markets.", "。美元近期走弱，对黄金与新兴市场形成支撑。",
               ". El dólar se ha debilitado recientemente, lo que apoya al oro y "
               "a los mercados emergentes.",
               ". Le dollar s'est récemment affaibli, ce qui soutient l'or et les "
               "marchés émergents.",
               ". Der Dollar hat zuletzt nachgegeben, was Gold und "
               "Schwellenländer stützt.",
               "。米ドルは直近で下落しており、金と新興市場を下支えしている。",
               ". Доллар недавно ослаб, что поддерживает золото и развивающиеся "
               "рынки.",
               ". O dólar se enfraqueceu recentemente, apoiando o ouro e os "
               "mercados emergentes."),
    "iu_strong": (". The dollar has strengthened recently, pressuring EM markets "
                  "and commodities.", "。美元近期走强，新兴市场与大宗商品承压。",
                  ". El dólar se ha fortalecido recientemente, presionando a los "
                  "emergentes y las materias primas.",
                  ". Le dollar s'est récemment renforcé, pesant sur les émergents "
                  "et les matières premières.",
                  ". Der Dollar hat zuletzt angezogen, was Schwellenländer und "
                  "Rohstoffe belastet.",
                  "。米ドルは直近で上昇しており、新興市場とコモディティに圧力。",
                  ". Доллар недавно укрепился, давя на развивающиеся рынки и сырьё.",
                  ". O dólar se fortaleceu recentemente, pressionando os "
                  "emergentes e as commodities."),
    "iu_mid": (". The dollar is range-bound overall with little direction.",
               "。美元整体区间震荡，方向感不强。",
               ". El dólar oscila en un rango, sin dirección clara.",
               ". Le dollar évolue dans un range, sans direction nette.",
               ". Der Dollar notiert insgesamt seitwärts ohne klare Richtung.",
               "。米ドルは全体的にレンジ内で方向感に欠ける。",
               ". Доллар в целом движется в диапазоне без ясного направления.",
               ". O dólar oscila em range, sem direção clara."),
    "iu_cny_dn": ("CNY appreciated {c}% over the same period.",
                  "人民币同期升值 {c}%。",
                  "El CNY se apreció {c}% en el mismo período.",
                  "Le CNY s'est apprécié de {c}% sur la même période.",
                  "Der CNY wertete im selben Zeitraum um {c}% auf.",
                  "人民元は同期間に{c}%切り上げ。",
                  "Юань за тот же период укрепился на {c}%.",
                  "O CNY se apreciou {c}% no mesmo período."),
    "iu_cny_up": ("CNY depreciated {c}% over the same period.",
                  "人民币同期贬值 {c}%。",
                  "El CNY se depreció {c}% en el mismo período.",
                  "Le CNY s'est déprécié de {c}% sur la même période.",
                  "Der CNY wertete im selben Zeitraum um {c}% ab.",
                  "人民元は同期間に{c}%切り下げ。",
                  "Юань за тот же период ослаб на {c}%.",
                  "O CNY se desvalorizou {c}% no mesmo período."),
    # ---- 今日解读：风险 ----
    "iv_s1_low": ("VIX at {v} — a low level; S&P 500 {s}% over 3 months. ",
                  "VIX 处于 {v} 的低位水平，标普500 近3月 {s}%。",
                  "VIX en {v}, un nivel bajo; S&P 500 {s}% en 3 meses. ",
                  "VIX à {v}, un niveau bas ; S&P 500 {s}% sur 3 mois. ",
                  "VIX bei {v}, ein niedriges Niveau; S&P 500 {s}% in 3 Monaten. ",
                  "VIX は {v} の低水準。S&P500 は直近3ヶ月 {s}%。",
                  "VIX на {v} — низкий уровень; S&P 500 {s}% за 3 месяца. ",
                  "VIX em {v}, um nível baixo; S&P 500 {s}% em 3 meses. "),
    "iv_s1_norm": ("VIX at {v} — a normal level; S&P 500 {s}% over 3 months. ",
                   "VIX 处于 {v} 的正常水平，标普500 近3月 {s}%。",
                   "VIX en {v}, un nivel normal; S&P 500 {s}% en 3 meses. ",
                   "VIX à {v}, un niveau normal ; S&P 500 {s}% sur 3 mois. ",
                   "VIX bei {v}, ein normales Niveau; S&P 500 {s}% in 3 Monaten. ",
                   "VIX は {v} の通常水準。S&P500 は直近3ヶ月 {s}%。",
                   "VIX на {v} — нормальный уровень; S&P 500 {s}% за 3 месяца. ",
                   "VIX em {v}, um nível normal; S&P 500 {s}% em 3 meses. "),
    "iv_s1_high": ("VIX at {v} — an elevated level; S&P 500 {s}% over 3 months. ",
                   "VIX 处于 {v} 的偏高水平，标普500 近3月 {s}%。",
                   "VIX en {v}, un nivel alto; S&P 500 {s}% en 3 meses. ",
                   "VIX à {v}, un niveau élevé ; S&P 500 {s}% sur 3 mois. ",
                   "VIX bei {v}, ein hohes Niveau; S&P 500 {s}% in 3 Monaten. ",
                   "VIX は {v} の高め水準。S&P500 は直近3ヶ月 {s}%。",
                   "VIX на {v} — повышенный уровень; S&P 500 {s}% за 3 месяца. ",
                   "VIX em {v}, um nível elevado; S&P 500 {s}% em 3 meses. "),
    "iv_low_up": ("Low volatility plus rising equities: risk sentiment is "
                  "stable, but beware complacency risk at low VIX.",
                  "低波动+股市上行，风险资产情绪稳定，但需留意低位VIX的自满风险。",
                  "Baja volatilidad y acciones al alza: el sentimiento de riesgo "
                  "es estable, pero cuidado con la complacencia con el VIX bajo.",
                  "Faible volatilité et actions en hausse : le sentiment de "
                  "risque est stable, mais méfiez-vous de la complaisance avec "
                  "un VIX bas.",
                  "Niedrige Volatilität plus steigende Aktien: die Risikostimmung "
                  "ist stabil, aber Vorsicht vor Selbstzufriedenheit bei "
                  "niedrigem VIX.",
                  "低ボラティリティと株価上昇：リスクセンチメントは安定しているが、"
                  "低VIXによる慢心リスクに注意。",
                  "Низкая волатильность плюс рост акций: настроения стабильны, но "
                  "остерегайтесь самоуспокоенности при низком VIX.",
                  "Baixa volatilidade e ações em alta: o sentimento de risco está "
                  "estável, mas cuidado com a complacência com o VIX baixo."),
    "iv_high": ("Volatility is rising — consider reducing risk exposure and "
                "tilting toward defense.", "波动抬升，建议降低风险敞口、关注防御。",
                "La volatilidad sube: conviene reducir la exposición al riesgo y "
                "prestar atención a la defensa.",
                "La volatilité monte — envisager de réduire l'exposition au "
                "risque et de privilégier la défense.",
                "Die Volatilität steigt — Risikopositionen reduzieren und auf "
                "Defensive achten.",
                "ボラティリティが上昇。リスクエクスポージャーの縮小とディフェンシブ"
                "重視を検討。",
                "Волатильность растёт — подумайте о снижении риска и перекосе в "
                "защиту.",
                "A volatilidade está subindo — considere reduzir a exposição ao "
                "risco e priorizar a defesa."),
    "iv_mid": ("Risk sentiment is broadly stable.", "风险情绪整体平稳。",
               "El sentimiento de riesgo es en general estable.",
               "Le sentiment de risque est globalement stable.",
               "Die Risikostimmung ist insgesamt stabil.",
               "リスクセンチメントは概して安定している。",
               "Настроения в целом стабильны.",
               "O sentimento de risco é em geral estável."),
    "iv_hyg": ("High yield bonds {h}% over 3 months — credit spreads widening, "
               "a marginal warning sign.",
               "高收益债近3月 {h}%，信用利差走扩，是边际上的风险预警。",
               "La renta alta {h}% en 3 meses: los diferenciales de crédito se "
               "amplían, una señal de alerta marginal.",
               "Les obligations à haut rendement {h}% sur 3 mois — les spreads "
               "de crédit s'élargissent, un signal d'alerte marginal.",
               "High-Yield-Anleihen {h}% in 3 Monaten — Kreditspreads weiten "
               "sich, ein marginales Warnsignal.",
               "ハイイールド債は直近3ヶ月 {h}%。クレジットスプレッドが拡大しており、"
               "限界的なリスク警報。",
               "Высокодоходные облигации {h}% за 3 месяца — кредитные спреды "
               "расширяются, пограничный предупреждающий сигнал.",
               "Títulos de alto rendimento {h}% em 3 meses — spreads de crédito "
               "se ampliando, um sinal de alerta marginal."),
    # ---- 今日解读：通胀（油价）----
    "io_s1": ("Brent {c1}% over the past month ({c3}% over 3 months). ",
              "布伦特近1月 {c1}%（近3月 {c3}%）。",
              "Brent {c1}% en el último mes ({c3}% en 3 meses). ",
              "Brent {c1}% sur le mois écoulé ({c3}% sur 3 mois). ",
              "Brent {c1}% im letzten Monat ({c3}% in 3 Monaten). ",
              "ブレントは直近1ヶ月 {c1}%（3ヶ月 {c3}%）。",
              "Brent {c1}% за месяц ({c3}% за 3 месяца). ",
              "Brent {c1}% no último mês ({c3}% em 3 meses). "),
    "io_hot": ("Energy prices have surged in the short term — the inflation "
               "variable to watch most closely; could delay Fed rate cuts.",
               "能源价格短期明显走强，成为当前最值得关注的通胀变量，或推迟美联储降息。",
               "Los precios de la energía se han disparado a corto plazo: la "
               "variable inflacionaria más importante a vigilar; podría retrasar "
               "las bajadas de tasas de la Fed.",
               "Les prix de l'énergie ont nettement augmenté à court terme — la "
               "variable inflationniste à surveiller de près ; cela pourrait "
               "retarder les baisses de taux de la Fed.",
               "Die Energiepreise sind kurzfristig deutlich gestiegen — die "
               "wichtigste Inflationsvariable; könnte Fed-Zinssenkungen "
               "verzögern.",
               "エネルギー価格が短期的に明確に上昇。現在最も注視すべきインフレ変数で"
               "あり、FRBの利下げを遅らせる可能性。",
               "Цены на энергоносители заметно выросли в краткосрочном периоде — "
               "ключевая инфляционная переменная; могут отложить снижение ставок "
               "ФРС.",
               "Os preços de energia subiram acentuadamente no curto prazo — a "
               "variável inflacionária mais importante a monitorar; pode atrasar "
               "os cortes de taxas do Fed."),
    "io_cool": ("Oil's momentum has faded after the spike; inflation "
                "expectations cooling at the margin.",
                "油价短期冲高后动能减弱，通胀预期边际降温。",
                "El impulso del petróleo se ha moderado tras el repunte; las "
                "expectativas de inflación se enfrían en el margen.",
                "L'élan du pétrole s'est essoufflé après le pic ; les "
                "anticipations d'inflation se refroidissent à la marge.",
                "Das Öl-Momentum hat nach dem Anstieg nachgelassen; "
                "Inflationserwartungen kühlen sich am Rand ab.",
                "油価は短期的な急騰後に勢いが鈍化。インフレ期待は限界的に低下。",
                "Импульс нефти угас после всплеска; инфляционные ожидания "
                "охлаждаются на краю.",
                "O momento do petróleo enfraqueceu após a alta; as expectativas "
                "de inflação esfriam na margem."),
    "io_flat": ("Oil is steady overall — not an inflation disturbance for now.",
                "油价整体平稳，暂不构成通胀扰动。",
                "El petróleo se mantiene estable; por ahora no es una "
                "perturbación inflacionaria.",
                "Le pétrole est globalement stable — pas de perturbation "
                "inflationniste pour l'instant.",
                "Das Öl ist insgesamt stabil — derzeit keine Inflationsstörung.",
                "油価は全体的に安定しており、当面インフレ攪乱要因ではない。",
                "Нефть в целом стабильна — пока не является инфляционным "
                "возмущением.",
                "O petróleo está estável no geral — por enquanto não é uma "
                "perturbação inflacionária."),
    # ---- 今日解读：黄金 ----
    "ig_s1": ("Gold {c1}% over the past month, {c3}% over 3 months. ",
              "黄金近1月 {c1}%、近3月 {c3}%。",
              "Oro {c1}% en el último mes, {c3}% en 3 meses. ",
              "Or {c1}% sur le mois écoulé, {c3}% sur 3 mois. ",
              "Gold {c1}% im letzten Monat, {c3}% in 3 Monaten. ",
              "金は直近1ヶ月 {c1}%、3ヶ月 {c3}%。",
              "Золото {c1}% за месяц, {c3}% за 3 месяца. ",
              "Ouro {c1}% no último mês, {c3}% em 3 meses. "),
    "ig_up_wd": ("Gold rising alongside a weaker dollar — a classic "
                 "easing/supportive mix; gold strength is fundamentally backed.",
                 "黄金上涨与美元走弱同现，属典型的宽松/支撑组合，黄金强势有基本面配合。",
                 "El oro sube junto con un dólar más débil: una combinación "
                 "clásica de relajación/apoyo; la fortaleza del oro tiene "
                 "respaldo fundamental.",
                 "L'or monte avec un dollar plus faible — un mix classique "
                 "d'assouplissement/soutien ; la force de l'or est "
                 "fondamentalement adossée.",
                 "Gold steigt bei schwächerem Dollar — eine klassische "
                 "Lockerungs-/Stützkombination; die Goldstärke ist fundamental "
                 "gestützt.",
                 "金の上昇と米ドル安が同時発生。典型的な緩和/支援の組合せで、金の強さには"
                 "ファンダメンタルズの裏付けがある。",
                 "Золото растёт на фоне слабого доллара — классическая комбинация "
                 "смягчения/поддержки; сила золота подкреплена фундаментально.",
                 "O ouro sobe junto com um dólar mais fraco — uma combinação "
                 "clássica de afrouxamento/apoio; a força do ouro tem respaldo "
                 "fundamental."),
    "ig_up_sd": ("Gold and the dollar rising together is unusual — typically "
                 "driven by safe-haven demand or central-bank buying; watch "
                 "event risk.",
                 "黄金与美元同涨，属异常组合，多由避险需求或央行购金主导，需警惕事件风险。",
                 "Que el oro y el dólar suban a la vez es inusual: suele deberse "
                 "a la demanda refugio o a compras de bancos centrales; vigilar "
                 "el riesgo de eventos.",
                 "L'or et le dollar montant ensemble est inhabituel — "
                 "généralement tiré par la demande de valeurs refuges ou les "
                 "achats de banques centrales ; surveiller le risque "
                 "événementiel.",
                 "Gold und Dollar steigen gemeinsam — ungewöhnlich, meist von "
                 "Krisenschutz-Nachfrage oder Zentralbankkäufen getrieben; "
                 "Ereignisrisiko beachten.",
                 "金と米ドルが同時に上昇するのは異例。多くは有事の需要や中央銀行の買いが"
                 "主導で、イベントリスクに注意が必要。",
                 "Золото и доллар растут вместе — нетипичная комбинация, обычно "
                 "движимая спросом на защитные активы или покупками центробанков; "
                 "следите за событийными рисками.",
                 "Ouro e dólar subindo juntos é incomum — geralmente impulsionado "
                 "por demanda de porto seguro ou compras de bancos centrais; "
                 "atenção ao risco de eventos."),
    "ig_dn": ("Gold pulling back; if the dollar is strengthening at the same "
              "time, it's rate/FX pressure — the trend is not broken.",
              "黄金回调，若美元同期走强则属利率/汇率压制，趋势未破坏。",
              "El oro retrocede; si el dólar se fortalece al mismo tiempo, es "
              "presión de tasas/FX; la tendencia no está rota.",
              "L'or recule ; si le dollar se renforce en même temps, c'est une "
              "pression de taux/changes — la tendance n'est pas rompue.",
              "Gold gibt nach; stärkt sich der Dollar gleichzeitig, handelt es "
              "sich um Zins-/FX-Druck — der Trend ist nicht gebrochen.",
              "金は調整中。同期に米ドルが強ければ金利・為替による圧力であり、トレンドは"
              "破れていない。",
              "Золото откатывается; если доллар одновременно крепнет, это "
              "давление ставок/валюты — тренд не сломан.",
              "O ouro recua; se o dólar se fortalecer ao mesmo tempo, é pressão "
              "de taxas/câmbio — a tendência não está rompida."),
    "ig_flat": ("Gold is consolidating, awaiting direction.", "黄金横盘，等待方向。",
                "El oro se consolida, a la espera de dirección.",
                "L'or consolide, dans l'attente d'une direction.",
                "Gold konsolidiert und wartet auf Richtung.", "金は横ばいで方向待ち。",
                "Золото консолидируется, ожидая направления.",
                "O ouro está consolidando, aguardando direção."),
})

TR.update({
    # ---- 四因子 ----
    "fc_growth": ("Growth", "增长", "Crecimiento", "Croissance", "Wachstum",
                  "成長", "Рост", "Crescimento"),
    "fc_infl": ("Inflation", "通胀", "Inflación", "Inflation", "Inflation",
                "インフレ", "Инфляция", "Inflação"),
    "fc_liq": ("Liquidity", "流动性", "Liquidez", "Liquidité", "Liquidität",
               "流動性", "Ликвидность", "Liquidez"),
    "fc_risk": ("Risk Appetite", "风险偏好", "Apetito de Riesgo",
                "Appétit pour le Risque", "Risikoappetit", "リスク選好",
                "Аппетит к риску", "Apetito de Risco"),
    "fc_strong": ("Strong", "偏强", "Fuerte", "Fort", "Stark", "強め",
                  "Сильный", "Forte"),
    "fc_weak": ("Weak", "偏弱", "Débil", "Faible", "Schwach", "弱め",
                "Слабый", "Fraco"),
    "fc_neu": ("Neutral", "中性", "Neutral", "Neutre", "Neutral", "中立",
               "Нейтральный", "Neutro"),
    "fc_na": ("Insufficient data", "数据不足", "Datos insuficientes",
              "Données insuffisantes", "Unzureichende Daten", "データ不足",
              "Недостаточно данных", "Dados insuficientes"),
    # ---- Regime 象限 ----
    "qd_refl": ("Reflation", "Reflation 再通胀", "Reflación", "Reflation",
                "Reflation", "リフレーション", "Рефляция", "Reflação"),
    "qn_refl": ("Growth and inflation rising together", "增长与通胀同步上行",
                "Crecimiento e inflación suben a la vez",
                "Croissance et inflation en hausse ensemble",
                "Wachstum und Inflation steigen gemeinsam",
                "成長とインフレが同時に上昇",
                "Рост и инфляция растут вместе",
                "Crescimento e inflação subindo juntos"),
    "qd_gold": ("Goldilocks", "Goldilocks 金发姑娘", "Goldilocks", "Goldilocks",
                "Goldilocks", "ゴールディロックス", "Голдилокс", "Goldilocks"),
    "qn_gold": ("Growth up with mild inflation", "增长上行且通胀温和",
                "Crecimiento al alza con inflación moderada",
                "Croissance en hausse, inflation modérée",
                "Wachstum steigt bei moderater Inflation",
                "成長上昇かつインフレは穏やか",
                "Рост растёт при умеренной инфляции",
                "Crescimento em alta com inflação moderada"),
    "qd_stag": ("Stagflation", "Stagflation 滞胀", "Estanflación", "Stagflation",
                "Stagflation", "スタグフレーション", "Стагфляция", "Estagflação"),
    "qn_stag": ("Growth weakening while inflation rises", "增长走弱且通胀上行",
                "Crecimiento débil con inflación al alza",
                "Croissance affaiblie, inflation en hausse",
                "Wachstum schwächt sich, Inflation steigt",
                "成長は減速しインフレは上昇",
                "Рост слабеет, инфляция растёт",
                "Crescimento enfraquecendo com inflação em alta"),
    "qd_defl": ("Deflation / Recession", "Deflation 通缩/衰退",
                "Deflación / Recesión", "Déflation / Récession",
                "Deflation / Rezession", "デフレ・景気後退",
                "Дефляция / Рецессия", "Deflação / Recessão"),
    "qn_defl": ("Growth and inflation both weakening", "增长与通胀同步走弱",
                "Crecimiento e inflación se debilitan",
                "Croissance et inflation faiblissent",
                "Wachstum und Inflation schwächen sich",
                "成長とインフレが同時に減速",
                "Рост и инфляция оба слабеют",
                "Crescimento e inflação enfraquecendo"),
    "qd_irefl": ("Reflation-tilted", "偏 Reflation", "Sesgo reflacionario",
                 "Inclinaison reflation", "Reflationstendenz", "リフレーション寄り",
                 "Смещение к рефляции", "Viés de Reflação"),
    "qn_irefl": ("Inflation rising dominates", "通胀上行主导",
                 "Domina la inflación al alza", "L'inflation en hausse domine",
                 "Steigende Inflation dominiert", "インフレ上昇が主導",
                 "Доминирует рост инфляции", "Inflação em alta domina"),
    "qd_idis": ("Disinflation-tilted", "偏 Disinflation", "Sesgo desinflacionario",
                "Inclinaison désinflation", "Disinflationstendenz",
                "ディスインフレ寄り", "Смещение к дезинфляции",
                "Viés de Desinflação"),
    "qn_idis": ("Inflation falling dominates", "通胀回落主导",
                "Domina la inflación a la baja", "L'inflation en baisse domine",
                "Fallende Inflation dominiert", "インフレ低下が主導",
                "Доминирует снижение инфляции", "Inflação em queda domina"),
    "qd_neu": ("Neutral zone", "中性区间", "Zona neutral", "Zone neutre",
               "Neutrale Zone", "中立ゾーン", "Нейтральная зона", "Zona neutra"),
    "qn_neu": ("No clear direction in growth or inflation", "增长与通胀均无明确方向",
               "Sin dirección clara en crecimiento o inflación",
               "Pas de direction nette croissance/inflation",
               "Keine klare Richtung bei Wachstum oder Inflation",
               "成長・インフレともに明確な方向感なし",
               "Нет ясного направления роста и инфляции",
               "Sem direção clara de crescimento ou inflação"),
    "qd_na": ("Unknown", "未知", "Desconocido", "Inconnu", "Unbekannt", "不明",
              "Неизвестно", "Desconhecido"),
    "qn_na": ("Insufficient data", "数据不足", "Datos insuficientes",
              "Données insuffisantes", "Unzureichende Daten", "データ不足",
              "Недостаточно данных", "Dados insuficientes"),
    # ---- 宏观分歧 ----
    "dv_n_credit": ("Equities vs Credit", "股票 vs 信用", "Acciones vs Crédito",
                    "Actions vs Crédit", "Aktien vs Kredit", "株 vs 信用",
                    "Акции vs Кредит", "Ações vs Crédito"),
    "dv_credit": ("S&P 500 {s}% over 3 months while high-yield bonds {h}% — "
                  "equities and credit markets are diverging; the credit side "
                  "is starting to warn.",
                  "标普500 近3月 {s}% 上行，但高收益债 {h}% 走弱，股市与信用市场出现分歧，信用端已开始预警。",
                  "El S&P 500 {s}% en 3 meses mientras la renta alta {h}% — "
                  "las acciones y el crédito divergen; el lado crediticio ya "
                  "advertía.",
                  "Le S&P 500 {s}% sur 3 mois tandis que la haute rendement "
                  "{h}% — actions et crédit divergent ; le crédit commence à "
                  "alerter.",
                  "S&P 500 {s}% in 3 Monaten, während High-Yield {h}% — Aktien "
                  "und Kreditmärkte divergieren; die Kreditseite warnt bereits.",
                  "S&P500は3ヶ月 {s}% 上昇一方、ハイイールド債は {h}%。株式と信用市場が乖離しており、信用側がすでに警鐘を鳴らしている。",
                  "S&P 500 {s}% за 3 месяца при высокодоходных облигациях {h}% "
                  "— акции и кредитный рынок расходятся; кредитная сторона уже "
                  "предупреждает.",
                  "S&P 500 {s}% em 3 meses enquanto títulos de alto rendimento "
                  "{h}% — ações e crédito divergem; o lado do crédito já alerta."),
    "dv_n_gold": ("Gold vs USD", "黄金 vs 美元", "Oro vs Dólar", "Or vs Dollar",
                  "Gold vs USD", "金 vs 米ドル", "Золото vs Доллар", "Ouro vs Dólar"),
    "dv_gold": ("Gold {g}% and the dollar {d}% over 3 months, rising together — "
                "an unusual mix usually driven by safe-haven demand or "
                "central-bank buying; watch event risk.",
                "黄金近3月 {g}% 与美元 {d}% 同涨，属异常组合，通常由避险需求或央行购金主导，警惕事件风险。",
                "El oro {g}% y el dólar {d}% en 3 meses subiendo juntos — una "
                "combinación inusual impulsada por demanda refugio o compras de "
                "bancos centrales; vigilar el riesgo de eventos.",
                "L'or {g}% et le dollar {d}% sur 3 mois, en hausse ensemble — "
                "un mix inhabituel tiré par la demande refuge ou les achats de "
                "banques centrales ; surveiller le risque événementiel.",
                "Gold {g}% und der Dollar {d}% in 3 Monaten steigen gemeinsam — "
                "eine ungewöhnliche Kombination, meist von Krisenschutz-Nachfrage "
                "oder Zentralbankkäufen getrieben; Ereignisrisiko beachten.",
                "金は3ヶ月 {g}%、米ドルは {d}% と同時に上昇。異例の組合せで、多くは有事の需要や中央銀行の買いが主導。イベントリスクに注意。",
                "Золото {g}% и доллар {d}% за 3 месяца растут вместе — "
                "нетипичная комбинация, обычно движимая спросом на защитные "
                "активы или покупками центробанков; следите за событийными "
                "рисками.",
                "Ouro {g}% e o dólar {d}% em 3 meses subindo juntos — uma "
                "combinação incomum, geralmente impulsionada por demanda de "
                "porto seguro ou compras de bancos centrais; atenção ao risco "
                "de eventos."),
    "dv_n_copper": ("Equities vs Copper", "股票 vs 铜", "Acciones vs Cobre",
                    "Actions vs Cuivre", "Aktien vs Kupfer", "株 vs 銅",
                    "Акции vs Медь", "Ações vs Cobre"),
    "dv_copper": ("Equities {s}% over 3 months but copper {c}% — the "
                  "cyclical/demand side has not confirmed equity optimism; "
                  "growth expectations are diverging.",
                  "股市近3月 {s}% 上行，但铜 {c}% 走弱，周期/需求端未确认股市乐观，增长预期存在分歧。",
                  "Las acciones {s}% en 3 meses pero el cobre {c}% — el lado "
                  "cíclico/demanda no confirma el optimismo bursátil; las "
                  "expectativas de crecimiento divergen.",
                  "Les actions {s}% sur 3 mois mais le cuivre {c}% — le versant "
                  "cyclique/demande ne confirme pas l'optimisme boursier ; les "
                  "anticipations de croissance divergent.",
                  "Aktien {s}% in 3 Monaten, aber Kupfer {c}% — die zyklische "
                  "Nachfrageseite bestätigt die Aktienoptimismus nicht; "
                  "Wachstumserwartungen divergieren.",
                  "株式は3ヶ月 {s}% 上昇したが、銅は {c}% 下落。サイクル／需要側が株式の楽観を裏付けておらず、成長期待に乖離がある。",
                  "Акции {s}% за 3 месяца, но медь {c}% — циклическая/спросовая "
                  "сторона не подтверждает биржевой оптимизм; ожидания роста "
                  "расходятся.",
                  "Ações {s}% em 3 meses, mas o cobre {c}% — o lado cíclico/"
                  "demanda não confirma o otimismo das ações; as expectativas "
                  "de crescimento divergem."),
    "dv_n_style": ("Nasdaq vs S&P", "纳指 vs 标普", "Nasdaq vs S&P",
                   "Nasdaq vs S&P", "Nasdaq vs S&P", "NASDAQ vs S&P500",
                   "Nasdaq vs S&P", "Nasdaq vs S&P"),
    "dv_style": ("Nasdaq vs S&P 500 3-month performance gap {g}pp ({lead}) — "
                 "a significant growth/value style divergence.",
                 "纳指与标普500 近3月涨幅差 {g}pp（{lead}），成长/价值风格显著分化。",
                 "Brecha de 3 meses entre Nasdaq y S&P 500 de {g}pp ({lead}) — "
                 "una divergencia estilística crecimiento/valor significativa.",
                 "Écart de performance sur 3 mois Nasdaq vs S&P 500 de {g}pp "
                 "({lead}) — une divergence de style croissance/valeur "
                 "notable.",
                 "3-Monats-Performanceunterschied Nasdaq vs S&P 500 von {g}pp "
                 "({lead}) — eine deutliche Wachstums/Wert-Stildivergenz.",
                 "NASDAQとS&P500の3ヶ月パフォーマンス差は {g}pp（{lead}）。グロース／バリュー・スタイルの大幅な分化。",
                 "Разрыв доходности Nasdaq и S&P 500 за 3 месяца {g}pp ({lead}) "
                 "— значительная дивергенция стилей рост/стоимость.",
                 "Diferença de performance de 3 meses Nasdaq vs S&P 500 de "
                 "{g}pp ({lead}) — uma divergência significativa de estilo "
                 "crescimento/valor."),
    "dv_lead_ndx": ("Nasdaq leading", "纳指领涨", "Nasdaq lidera", "Nasdaq mène",
                    "Nasdaq führt", "NASDAQが主導", "Nasdaq лидирует",
                    "Nasdaq liderando"),
    "dv_lead_spx": ("Nasdaq lagging", "纳指落后", "Nasdaq rezaga",
                    "Nasdaq à la traîne", "Nasdaq hinkt hinterher",
                    "NASDAQが出遅れ", "Nasdaq отстаёт", "Nasdaq atrasado"),
    "dv_n_oil": ("Oil vs Equities", "油价 vs 股市", "Petróleo vs Acciones",
                 "Pétrole vs Actions", "Öl vs Aktien", "油価 vs 株式",
                 "Нефть vs Акции", "Petróleo vs Ações"),
    "dv_oil": ("Brent {b}% over 3 months while equities keep rising — the "
               "market has not yet priced the drag of energy inflation on "
               "earnings and valuations.",
               "布伦特近3月 {b}% 大涨而股市仍在上行，股市尚未计入能源通胀对盈利与估值的压制。",
               "El Brent {b}% en 3 meses mientras las acciones siguen subiendo "
               "— el mercado aún no descuenta el freno de la inflación "
               "energética sobre beneficios y valoraciones.",
               "Le Brent {b}% sur 3 mois tandis que les actions continuent de "
               "monter — le marché n'a pas encore intégré l'impact de "
               "l'inflation énergétique sur les bénéfices et les valorisations.",
               "Brent {b}% in 3 Monaten, während Aktien weiter steigen — der "
               "Markt hat den Drag der Energieinflation auf Gewinne und "
               "Bewertungen noch nicht eingepreist.",
               "ブレントは3ヶ月 {b}% 大幅高ながら株式も上昇。エネルギーインフレが収益とバリュエーションに与える圧力はまだ織り込まれていない。",
               "Brent {b}% за 3 месяца, а акции продолжают расти — рынок ещё "
               "не заложил давление энергетической инфляции на прибыли и "
               "оценки активов.",
               "Brent {b}% em 3 meses enquanto as ações seguem em alta — o "
               "mercado ainda não precificou o efeito da inflação de energia "
               "sobre lucros e avaliações."),
    # ---- 资产倾向 ----
    "ab_spx": ("S&P 500", "标普500", "S&P 500", "S&P 500", "S&P 500",
               "S&P500", "S&P 500", "S&P 500"),
    "ab_ndx": ("Nasdaq", "纳斯达克", "Nasdaq", "Nasdaq", "Nasdaq", "ナスダック",
               "Nasdaq", "Nasdaq"),
    "ab_gold": ("Gold", "黄金", "Oro", "Or", "Gold", "金", "Золото", "Ouro"),
    "ab_oil": ("Crude Oil", "原油", "Petróleo", "Pétrole", "Rohöl", "原油",
               "Нефть", "Petróleo"),
    "ab_bond": ("US Treasuries (Duration)", "美债(久期)", "Tesoros EE.UU. (Duración)",
                "Treasuries US (Duration)", "US-Staatsanleihen (Duration)",
                "米国債（デュレーション）", "Госдолг США (дюрация)",
                "Treasuries dos EUA (Duração)"),
    "ab_usd": ("US Dollar", "美元", "Dólar EE.UU.", "Dollar US", "US-Dollar",
               "米ドル", "Доллар США", "Dólar dos EUA"),
    "ab_hy": ("High-Yield Credit", "高收益信用", "Crédito de Alto Rendimiento",
              "Crédit à Haut Rendement", "High-Yield-Kredit",
              "ハイイールド・クレジット", "Высокодоходный кредит",
              "Crédito de Alto Rendimento"),
    "bb_pos": ("🟢 Positive", "🟢 正面", "🟢 Positivo", "🟢 Positif",
               "🟢 Positiv", "🟢 ポジティブ", "🟢 Позитивно", "🟢 Positivo"),
    "bb_neg": ("🔴 Cautious", "🔴 谨慎", "🔴 Cautela", "🔴 Prudence",
               "🔴 Vorsichtig", "🔴 慎重", "🔴 Осторожно", "🔴 Cautela"),
    "bb_neu": ("🟡 Neutral", "🟡 中性", "🟡 Neutral", "🟡 Neutre", "🟡 Neutral",
               "🟡 中立", "🟡 Нейтрально", "🟡 Neutro"),
    "bb_watch": ("🟡 Watch", "🟡 关注", "🟡 Vigilar", "🟡 À surveiller",
                 "🟡 Beobachten", "🟡 注視", "🟡 Следить", "🟡 Observar"),
    "ab_spx_p": ("Risk appetite and growth in sync — tailwind",
                 "风险偏好与增长共振，顺风",
                 "Apetito de riesgo y crecimiento en sintonía — viento a favor",
                 "Appétit pour le risque et croissance alignés — vent portant",
                 "Risikoappetit und Wachstum im Einklang — Rückenwind",
                 "リスク選好と成長が同調、追い風",
                 "Аппетит к риску и рост согласованы — попутный ветер",
                 "Apetito de risco e crescimento em sintonia — vento a favor"),
    "ab_spx_n": ("Risk appetite deteriorating — headwind", "风险偏好恶化，逆风",
                 "El apetito de riesgo se deteriora — viento en contra",
                 "Appétit pour le risque en dégradation — vent contraire",
                 "Risikoappetit verschlechtert sich — Gegenwind",
                 "リスク選好が悪化、逆風",
                 "Аппетит к риску ухудшается — встречный ветер",
                 "Apetito de risco deteriorando — vento contra"),
    "ab_ndx_p": ("Growth stocks benefit from growth + risk appetite",
                 "成长股受益于增长+风险偏好",
                 "Las acciones de crecimiento se benefician de crecimiento + "
                 "apetito de riesgo",
                 "Les valeurs de croissance bénéficient de croissance + appétit "
                 "pour le risque",
                 "Wachstumsaktien profitieren von Wachstum + Risikoappetit",
                 "グロース株は成長＋リスク選好の恩恵を受ける",
                 "Акции роста выигрывают от роста и аппетита к риску",
                 "Ações de crescimento se beneficiam de crescimento + apetito "
                 "de risco"),
    "ab_ndx_n": ("Tightening liquidity or worsening risk appetite weigh on "
                 "valuations", "流动性收紧或风险偏好恶化压制估值",
                 "Menor liquidez o apetito de riesgo en deterioro presionan las "
                 "valoraciones",
                 "Le resserrement de la liquidité ou la dégradation de l'appétit "
                 "pour le risque pèsent sur les valorisations",
                 "Straffere Liquidität oder schlechterer Risikoappetit belasten "
                 "Bewertungen",
                 "流動性引き締めやリスク選好の悪化がバリュエーションを圧迫",
                 "Сжатие ликвидности или ухудшение аппетита к риску давят на "
                 "оценки активов",
                 "Liquidez mais apertada ou apetito de risco piorando pressionam "
                 "as avaliações"),
    "ab_gold_p": ("Rising inflation or easy liquidity supports gold",
                  "通胀上行或流动性宽松支撑金价",
                  "Inflación al alza o liquidez abundante apoyan al oro",
                  "Inflation en hausse ou liquidité abondante soutiennent l'or",
                  "Steigende Inflation oder lockere Liquidität stützen Gold",
                  "インフレ上昇や流動性緩和が金を下支え",
                  "Рост инфляции или мягкая ликвидность поддерживают золото",
                  "Inflação em alta ou liquidez abundante sustentam o ouro"),
    "ab_gold_n": ("Rising real rates are unfavorable for gold",
                  "实际利率上行环境对黄金不利",
                  "Las tasas reales al alza son desfavorables para el oro",
                  "La hausse des taux réels est défavorable à l'or",
                  "Steigende Realzinsen sind ungünstig für Gold",
                  "実質金利の上昇は金に不利",
                  "Рост реальных ставок неблагоприятен для золота",
                  "Taxas reais em alta são desfavoráveis para o ouro"),
    "ab_oil_i": ("Inflation trade lifts oil, but policy-tightening risk rises",
                 "通胀交易推升油价，但政策收紧风险上升",
                 "El trade de inflación impulsa el petróleo, pero sube el riesgo "
                 "de endurecimiento de política",
                 "Le trade inflationniste soutient le pétrole, mais le risque de "
                 "resserrement de politique augmente",
                 "Inflationstrade hebt das Öl, aber das Risiko strafferer Politik "
                 "steigt",
                 "インフレ・トレードが油価を押し上げるが、政策引き締めリスクが上昇",
                 "Инфляционная сделка поднимает нефть, но растёт риск ужесточения "
                 "политики",
                 "O trade de inflação impulsiona o petróleo, mas o risco de "
                 "aperto de política sobe"),
    "ab_oil_n": ("Tracks the inflation factor and geopolitics",
                 "跟随通胀因子与地缘局势",
                 "Sigue el factor de inflación y la geopolítica",
                 "Suit le facteur d'inflation et la géopolitique",
                 "Folgt dem Inflationsfaktor und der Geopolitik",
                 "インフレ・ファクターと地政学に連動",
                 "Следует за инфляционным фактором и геополитикой",
                 "Acompanha o fator de inflação e a geopolítica"),
    "ab_bond_i": ("Rising inflation pushes yields up, weighing on prices",
                  "通胀上行推升收益率、压制价格",
                  "La inflación al alza empuja los rendimientos y presiona los "
                  "precios",
                  "L'inflation en hausse pousse les rendements à la hausse et "
                  "pèse sur les prix",
                  "Steigende Inflation treibt die Renditen und drückt die Kurse",
                  "インフレ上昇が利回りを押し上げ、価格を圧迫",
                  "Рост инфляции поднимает доходности, давя на цены",
                  "Inflação em alta empurra os rendimentos para cima, pressionando "
                  "os preços"),
    "ab_bond_d": ("Falling inflation favors duration", "通胀回落利好久期",
                  "La inflación a la baja favorece la duración",
                  "La baisse de l'inflation favorise la duration",
                  "Fallende Inflation begünstigt Duration",
                  "インフレ低下はデュレーションに追い風",
                  "Снижение инфляции благоприятствует дюрации",
                  "Inflação em queda favorece a duração"),
    "ab_usd_p": ("Tighter global USD liquidity supports the dollar",
                 "全球美元流动性收紧支撑美元",
                 "Menor liquidez global en dólares apoya al dólar",
                 "Le resserrement de la liquidité mondiale en dollars soutient "
                 "le dollar",
                 "Straffere globale USD-Liquidität stützt den Dollar",
                 "世界的なドル流動性の引き締めが米ドルを下支え",
                 "Более тесная глобальная долларовая ликвидность поддерживает "
                 "доллар",
                 "Liquidez global em dólar mais apertada sustenta o dólar"),
    "ab_usd_n": ("Easy dollar liquidity — USD softer", "美元流动性宽松，美元偏弱",
                 "Liquidez abundante en dólares — dólar más débil",
                 "Liquidité abondante en dollars — dollar plus faible",
                 "Lockere Dollarliquidität — Dollar schwächer",
                 "ドル流動性が緩和し米ドルは弱含み",
                 "Мягкая долларовая ликвидность — доллар слабее",
                 "Liquidez em dólar abundante — dólar mais fraco"),
    "ab_hy_p": ("Improving risk appetite compresses credit spreads",
                "风险偏好改善压低信用利差",
                "La mejora del apetito de riesgo comprime los diferenciales de "
                "crédito",
                "L'amélioration de l'appétit pour le risque comprime les spreads "
                "de crédit",
                "Verbesserter Risikoappetit verengt Kreditspreads",
                "リスク選好の改善がクレジットスプレッドを縮小",
                "Улучшение аппетита к риску сжимает кредитные спреды",
                "Melhora do apetito de risco comprime os spreads de crédito"),
    "ab_hy_n": ("Risk appetite deteriorating — spreads widening",
                "风险偏好恶化，利差走扩",
                "El apetito de riesgo se deteriora — los diferenciales se "
                "amplían",
                "L'appétit pour le risque se dégrade — les spreads s'élargissent",
                "Risikoappetit verschlechtert sich — Spreads weiten sich",
                "リスク選好の悪化によりスプレッドが拡大",
                "Аппетит к риску ухудшается — спреды расширяются",
                "Apetito de risco deteriorando — spreads se ampliando"),
    # ---- 派生指标 ----
    "dv2_curve": ("10Y-2Y Term Spread", "期限利差 10Y-2Y",
                  "Diferencial 10Y-2Y", "Écart de taux 10Y-2Y",
                  "Zinsdifferenz 10Y-2Y", "期間スプレッド 10Y-2Y",
                  "Спред 10Y-2Y", "Spread 10Y-2Y"),
    "dv2_curve_steep": ("Curve steepening — markets pricing growth and "
                        "inflation recovery", "曲线陡峭化，市场定价增长与通胀回升",
                        "Endurecimiento de la curva — el mercado descuenta "
                        "recuperación de crecimiento e inflación",
                        "Creusement de la courbe — le marché anticipe une "
                        "reprise de la croissance et de l'inflation",
                        "Kurve steilt sich — der Markt preist eine Erholung von "
                        "Wachstum und Inflation ein",
                        "カーブのスティープ化。市場は成長とインフレの回復を織り込み",
                        "Кривая круче — рынок закладывает восстановление роста и "
                        "инфляции",
                        "Curva acentuando — o mercado precifica recuperação de "
                        "crescimento e inflação"),
    "dv2_curve_normal": ("Curve normal to steep — recession signal cleared",
                         "曲线正常偏陡，衰退信号解除",
                         "Curva normal a empinada — señal de recesión despejada",
                         "Courbe normale à pentue — signal de récession levé",
                         "Kurve normal bis steil — Rezessionssignal entschärft",
                         "カーブは正常からやや急。景気後退シグナルは解除",
                         "Кривая от нормальной к крутой — сигнал рецессии снят",
                         "Curva normal a acentuada — sinal de recessão dissipado"),
    "dv2_curve_mild": ("Mildly inverted — early warning", "曲线轻度倒挂，边缘预警",
                       "Levemente invertida — advertencia temprana",
                       "Légèrement inversée — signal d'alerte précoce",
                       "Leicht invertiert — Frühwarnung",
                       "軽度の逆転。早期警戒シグナル",
                       "Слабо инвертирована — раннее предупреждение",
                       "Levemente invertida — alerta antecipada"),
    "dv2_curve_inv": ("Curve inverted — historically leads recessions",
                      "曲线倒挂，历史上常领先经济衰退",
                      "Curva invertida — históricamente anticipa recesiones",
                      "Courbe inversée — précède historiquement les récessions",
                      "Kurve invertiert — geht historisch Rezessionen voraus",
                      "カーブ逆転。歴史的に景気後退に先行",
                      "Кривая инвертирована — исторически предшествует рецессиям",
                      "Curva invertida — historicamente antecede recessões"),
    "dv2_oilspread": ("Brent-WTI Spread", "Brent-WTI 价差", "Diferencial Brent-WTI",
                      "Écart Brent-WTI", "Brent-WTI-Spread", "Brent-WTI スプレッド",
                      "Спред Brent-WTI", "Spread Brent-WTI"),
    "dv2_oil_wide": ("Spread wide — US crude export/refinery factors dominate",
                     "价差偏宽，美国原油出口/炼厂因素主导",
                     "Diferencial amplio — dominan las exportaciones de crudo "
                     "EE.UU. y las refinerías",
                     "Écart large — prédominance des exportations de pétrole "
                     "américain et du raffinage",
                     "Spread weit — US-Rohölexport-/Raffineriefaktoren dominieren",
                     "スプレッド拡大。米原油輸出／精製要因が主導",
                     "Спред широкий — доминируют факторы экспорта нефти США и "
                     "НПЗ",
                     "Spread amplo — fatores de exportação de petróleo e "
                     "refino dos EUA dominam"),
    "dv2_oil_normal": ("Spread within normal range", "价差正常区间",
                       "Diferencial dentro del rango normal",
                       "Écart dans la fourchette normale",
                       "Spread im normalen Bereich", "スプレッドは通常レンジ",
                       "Спред в нормальном диапазоне", "Spread na faixa normal"),
    "dv2_goldoil": ("Gold-Oil Ratio (Gold÷Brent)", "金油比 (Gold÷Brent)",
                    "Ratio Oro-Petróleo (Oro÷Brent)", "Ratio Or-Pétrole (Or÷Brent)",
                    "Gold-Öl-Ratio (Gold÷Brent)", "金油比（金÷ブレント）",
                    "Отношение золото/нефть (Золото÷Brent)",
                    "Razão Ouro-Petróleo (Ouro÷Brent)"),
    "dv2_go_high": (">25 often signals recession/risk-off dominance; <15 "
                    "expansion dominance",
                    ">25 通常预示衰退/避险主导；<15 经济扩张主导",
                    ">25 suele señalar recesión/dominio de aversión al riesgo; "
                    "<15 dominio de expansión",
                    ">25 signale souvent récession/dominance du risk-off ; <15 "
                    "dominance de l'expansion",
                    ">25 signalisiert oft Rezession/Risk-off-Dominanz; <15 "
                    "Expansionsdominanz",
                    ">25 は景気後退／リスク回避優勢を示唆することが多い。<15 は景気拡大優勢",
                    ">25 часто сигнализирует рецессию или доминирование "
                    "защитных настроений; <15 — доминирование экспансии",
                    ">25 costuma sinalizar recessão/domínio de aversão ao risco; "
                    "<15 domínio de expansão"),
    "dv2_go_normal": ("Ratio in normal range", "比值处于常态区间",
                      "Ratio en rango normal", "Ratio dans la fourchette normale",
                      "Ratio im normalen Bereich", "比率は通常レンジ",
                      "Отношение в нормальном диапазоне", "Razão na faixa normal"),
    "dv2_dxy": ("USD Momentum (DXY 3M)", "美元动量 (DXY 3月)",
                "Momentum del dólar (DXY 3M)", "Momentum du dollar (DXY 3M)",
                "USD-Momentum (DXY 3M)", "米ドル・モメンタム（DXY 3ヶ月）",
                "Импульс доллара (DXY 3M)", "Momentum do dólar (DXY 3M)"),
    "dv2_dxy_up": ("Dollar stronger over 3 months — headwind for EM and "
                   "commodities", "美元3月走强，利空新兴市场与大宗",
                   "El dólar se fortalece en 3 meses — viento en contra para "
                   "emergentes y materias primas",
                   "Le dollar se renforce sur 3 mois — frein pour les émergents "
                   "et les matières premières",
                   "Dollar in 3 Monaten stärker — Belastung für Schwellenländer "
                   "und Rohstoffe",
                   "米ドルは3ヶ月で強含み。新興市場とコモディティには逆風",
                   "Доллар окреп за 3 месяца — помеха для развивающихся рынков и "
                   "сырья",
                   "Dólar mais forte em 3 meses — vento contra para emergentes e "
                   "commodities"),
    "dv2_dxy_dn": ("Dollar weaker over 3 months — supportive for risk assets "
                   "and gold", "美元3月走弱，利多风险资产与黄金",
                   "El dólar se debilita en 3 meses — apoyo para activos de "
                   "riesgo y oro",
                   "Le dollar s'affaiblit sur 3 mois — soutien aux actifs "
                   "risqués et à l'or",
                   "Dollar in 3 Monaten schwächer — Stütze für Risiko-Assets und "
                   "Gold",
                   "米ドルは3ヶ月で弱含み。リスク資産と金には追い風",
                   "Доллар ослаб за 3 месяца — поддержка рисковых активов и "
                   "золота",
                   "Dólar mais fraco em 3 meses — apoio para ativos de risco e "
                   "ouro"),
    "dv2_dxy_flat": ("Dollar range-bound — direction pending",
                     "美元区间震荡，方向待选择",
                     "Dólar en rango — dirección pendiente",
                     "Dollar en range — direction en attente",
                     "Dollar seitwärts — Richtung offen",
                     "米ドルはレンジ内。方向待ち",
                     "Доллар в диапазоне — направление ожидается",
                     "Dólar em consolidação — direção pendente"),
    "dv2_credit": ("Credit Conditions (HYG 3M)", "信用状态 (HYG 3月)",
                   "Condiciones de crédito (HYG 3M)", "Conditions de crédit (HYG 3M)",
                   "Kreditkonditionen (HYG 3M)", "信用状況（HYG 3ヶ月）",
                   "Кредитные условия (HYG 3M)", "Condições de crédito (HYG 3M)"),
    "dv2_cr_weak": ("High yield weakening — credit spreads widening, risk "
                    "appetite deteriorating",
                    "高收益债走弱，信用利差走扩，风险偏好恶化",
                    "La renta alta se debilita — los diferenciales de crédito se "
                    "amplían, el apetito de riesgo empeora",
                    "La haute rendement faiblit — les spreads de crédit "
                    "s'élargissent, l'appétit pour le risque se dégrade",
                    "High Yield schwächt sich — Kreditspreads weiten sich, der "
                    "Risikoappetit verschlechtert sich",
                    "ハイイールド債が弱含み。クレジットスプレッドが拡大し、リスク選好が悪化",
                    "Высокодоходный сегмент слабеет — кредитные спреды "
                    "расширяются, аппетит к риску ухудшается",
                    "Alto rendimento enfraquecendo — spreads de crédito se "
                    "ampliando, apetito de risco piorando"),
    "dv2_cr_strong": ("High yield strengthening — credit conditions easy",
                      "高收益债走强，信用环境宽松",
                      "La renta alta se fortalece — condiciones de crédito "
                      "holgadas",
                      "La haute rendement se renforce — conditions de crédit "
                      "souples",
                      "High Yield festigt sich — Kreditkonditionen locker",
                      "ハイイールド債が強含み。信用環境は緩和的",
                      "Высокодоходный сегмент крепнет — кредитные условия мягкие",
                      "Alto rendimento se fortalecendo — condições de crédito "
                      "frouxas"),
    "dv2_cr_calm": ("High yield steady — no stress signals in credit",
                    "高收益债平稳，信用市场无压力信号",
                    "Renta alta estable — sin señales de estrés en el crédito",
                    "Haute rendement stable — aucun signal de stress sur le "
                    "crédit",
                    "High Yield stabil — keine Stresssignale im Kredit",
                    "ハイイールド債は安定。信用市場にストレスの兆候なし",
                    "Высокодоходный сегмент стабилен — признаков стресса нет",
                    "Alto rendimento estável — sem sinais de estresse no crédito"),
})


def tr(key: str, lang: str = "zh", **params) -> str:
    """查 TR 表渲染指定语言的模板；参数值以 "@" 开头时先递归翻译。"""
    vals = TR.get(key)
    if not vals:
        return key
    idx = LANGS.index(lang) if lang in LANGS else 1
    s = vals[idx] if idx < len(vals) else vals[0]
    if params:
        p = {k: (tr(v[1:], lang) if isinstance(v, str) and v.startswith("@") else v)
             for k, v in params.items()}
        return s.format(**p)
    return s


# 指标多语言名称与一句话说明 {col: {lang: (label, desc)}}
IND_L10N = {
    "US10Y": {
        "en": ("US 10Y Treasury Yield", "The anchor of global asset pricing; high levels compress valuations"),
        "zh": ("美国10年期国债收益率", "全球资产定价之锚；高位压制估值"),
        "es": ("Rendimiento del Tesoro EE.UU. 10Y", "El ancla de la valoración de activos globales; niveles altos comprimen valoraciones"),
        "fr": ("Rendement du Trésor américain 10 ans", "L'ancre de la valorisation des actifs mondiaux ; des niveaux élevés compriment les valorisations"),
        "de": ("US-10Y-Staatsanleihenrendite", "Der Anker der globalen Asset-Bewertung; hohe Niveaus drücken Bewertungen"),
        "ja": ("米国10年国債利回り", "世界の資産価格のアンカー。高水準はバリュエーションを圧迫"),
        "ru": ("Доходность 10-летних казначейских облигаций США", "Якорь глобального ценообразования активов; высокие уровни сжимают оценки"),
        "pt": ("Rendimento do Tesouro 10Y dos EUA", "A âncora da precificação de ativos globais; níveis altos comprimem avaliações"),
    },
    "US2Y": {
        "en": ("US 2Y Treasury Yield", "Tracks Fed policy rate expectations"),
        "zh": ("美国2年期国债收益率", "更贴近美联储政策利率预期"),
        "es": ("Rendimiento del Tesoro EE.UU. 2Y", "Refleja las expectativas de la tasa de política de la Fed"),
        "fr": ("Rendement du Trésor américain 2 ans", "Proche des anticipations de taux directeur de la Fed"),
        "de": ("US-2Y-Staatsanleihenrendite", "Nahe an den Fed-Leitzinserwartungen"),
        "ja": ("米国2年国債利回り", "FRB政策金利期待に最も連動"),
        "ru": ("Доходность 2-летних казначейских облигаций США", "Ближе к ожиданиям ставки ФРС"),
        "pt": ("Rendimento do Tesouro 2Y dos EUA", "Acompanha as expectativas da taxa da Fed"),
    },
    "Brent": {
        "en": ("Brent Crude Oil", "Global inflation & geopolitics thermometer"),
        "zh": ("布伦特原油", "全球通胀与地缘温度计"),
        "es": ("Petróleo Brent", "Termómetro de la inflación global y la geopolítica"),
        "fr": ("Pétrole Brent", "Thermomètre de l'inflation mondiale et de la géopolitique"),
        "de": ("Brent-Rohöl", "Thermometer für globale Inflation und Geopolitik"),
        "ja": ("ブレント原油", "世界のインフレと地政学の温度計"),
        "ru": ("Нефть Brent", "Термометр мировой инфляции и геополитики"),
        "pt": ("Petróleo Brent", "Termômetro da inflação global e da geopolítica"),
    },
    "WTI": {
        "en": ("WTI Crude Oil", "US oil price benchmark"),
        "zh": ("WTI原油", "美国油价基准"),
        "es": ("Petróleo WTI", "Referente del precio del petróleo en EE.UU."),
        "fr": ("Pétrole WTI", "Référence du prix du pétrole américain"),
        "de": ("WTI-Rohöl", "US-Ölpreis-Benchmark"),
        "ja": ("WTI原油", "米国の油価ベンチマーク"),
        "ru": ("Нефть WTI", "Американский ориентир цен на нефть"),
        "pt": ("Petróleo WTI", "Referência do preço do petróleo nos EUA"),
    },
    "Gold": {
        "en": ("Gold", "Safe haven + inflation hedge; moves inversely to real rates"),
        "zh": ("黄金", "避险+抗通胀；与实际利率反向"),
        "es": ("Oro", "Refugio + cobertura de inflación; inverso a las tasas reales"),
        "fr": ("Or", "Valeur refuge + couverture inflation ; inverse aux taux réels"),
        "de": ("Gold", "Krisenschutz + Inflationsschutz; entgegengesetzt zu Realzinsen"),
        "ja": ("金", "安全資産＋インフレヘッジ。実質金利と逆相関"),
        "ru": ("Золото", "Защитный актив + хедж инфляции; обратно связано с реальными ставками"),
        "pt": ("Ouro", "Porto seguro + hedge de inflação; inverso às taxas reais"),
    },
    "Copper": {
        "en": ("COMEX Copper", "Global growth leading indicator; rising copper = recovering demand"),
        "zh": ("COMEX铜", "全球经济领先指标；铜涨=需求回暖"),
        "es": ("Cobre COMEX", "Indicador líder del crecimiento global; cobre al alza = demanda en recuperación"),
        "fr": ("Cuivre COMEX", "Indicateur avancé de la croissance mondiale ; cuivre en hausse = demande en rétablissement"),
        "de": ("COMEX-Kupfer", "Frühindikator für globales Wachstum; steigendes Kupfer = Erholung der Nachfrage"),
        "ja": ("COMEX銅", "世界経済の先行指標。銅高＝需要回復"),
        "ru": ("Медь COMEX", "Опережающий индикатор мирового роста; рост меди = восстановление спроса"),
        "pt": ("Cobre COMEX", "Indicador líder do crescimento global; cobre subindo = demanda em recuperação"),
    },
    "DXY": {
        "en": ("US Dollar Index (DXY)", "USD strength; strong → pressure on EM markets"),
        "zh": ("美元指数 DXY", "美元强弱；强→新兴市场承压"),
        "es": ("Índice del Dólar (DXY)", "Fuerza del dólar; fuerte → presión sobre emergentes"),
        "fr": ("Indice Dollar (DXY)", "Force du dollar ; fort → pression sur les émergents"),
        "de": ("US-Dollar-Index (DXY)", "USD-Stärke; stark → Druck auf Schwellenländer"),
        "ja": ("米ドル指数（DXY）", "米ドルの強弱。強い→新興市場に圧力"),
        "ru": ("Индекс доллара США (DXY)", "Сила доллара; сильный → давление на развивающиеся рынки"),
        "pt": ("Índice do Dólar (DXY)", "Força do dólar; forte → pressão sobre emergentes"),
    },
    "USDCNY": {
        "en": ("USD/CNY", "Affects FX conversion of foreign assets and capital flows"),
        "zh": ("美元兑人民币", "影响海外资产换算与资金流向"),
        "es": ("USD/CNY", "Afecta la conversión de activos extranjeros y los flujos de capital"),
        "fr": ("USD/CNY", "Affecte la conversion des actifs étrangers et les flux de capitaux"),
        "de": ("USD/CNY", "Beeinflusst Umrechnung ausländischer Anlagen und Kapitalflüsse"),
        "ja": ("米ドル/人民元", "海外資産の換算と資金流動に影響"),
        "ru": ("USD/CNY", "Влияет на пересчёт иностранных активов и потоки капитала"),
        "pt": ("USD/CNY", "Afeta a conversão de ativos externos e os fluxos de capital"),
    },
    "USDJPY": {
        "en": ("USD/JPY", "Carry-trade bellwether; sharp moves above 160 are risky"),
        "zh": ("美元兑日元", "套息交易风向标；急升破160有风险"),
        "es": ("USD/JPY", "Brújula del carry trade; subidas bruscas por encima de 160 son arriesgadas"),
        "fr": ("USD/JPY", "Baromètre du carry trade ; des hausses brusques au-dessus de 160 sont risquées"),
        "de": ("USD/JPY", "Carry-Trade-Barometer; schnelle Anstiege über 160 sind riskant"),
        "ja": ("米ドル/円", "キャリートレードの指標。160急突破はリスク"),
        "ru": ("USD/JPY", "Индикатор кэрри-трейда; резкие движения выше 160 рискованны"),
        "pt": ("USD/JPY", "Termômetro do carry trade; altas bruscas acima de 160 são arriscadas"),
    },
    "VIX": {
        "en": ("VIX Fear Index", "<15 calm, 20–30 tense, >30 panic"),
        "zh": ("VIX恐慌指数", "<15平静 20-30紧张 >30恐慌"),
        "es": ("Índice del Miedo VIX", "<15 calma, 20–30 tenso, >30 pánico"),
        "fr": ("Indice de peur VIX", "<15 calme, 20–30 tendu, >30 panique"),
        "de": ("VIX Fear Index", "<15 ruhig, 20–30 angespannt, >30 Panik"),
        "ja": ("VIX恐怖指数", "<15 平静、20-30 緊張、>30 パニック"),
        "ru": ("Индекс страха VIX", "<15 спокойно, 20–30 напряжённо, >30 паника"),
        "pt": ("Índice do Medo VIX", "<15 calmo, 20–30 tenso, >30 pânico"),
    },
    "SPX": {
        "en": ("S&P 500", "Global risk-asset bellwether"),
        "zh": ("标普500", "全球风险资产风向标"),
        "es": ("S&P 500", "Referente global de los activos de riesgo"),
        "fr": ("S&P 500", "Baromètre mondial des actifs risqués"),
        "de": ("S&P 500", "Globaler Kompass für Risiko-Assets"),
        "ja": ("S&P500", "世界のリスク資産の風見鶏"),
        "ru": ("S&P 500", "Мировой барометр рисковых активов"),
        "pt": ("S&P 500", "Referência global dos ativos de risco"),
    },
    "NDX": {
        "en": ("Nasdaq Composite", "Growth-stock risk appetite"),
        "zh": ("纳斯达克综合", "成长股风险偏好"),
        "es": ("Nasdaq Compuesto", "Apetito de riesgo de las acciones de crecimiento"),
        "fr": ("Nasdaq Composite", "Appétit pour le risque des valeurs de croissance"),
        "de": ("Nasdaq Composite", "Risikoappetit von Wachstumsaktien"),
        "ja": ("ナスダック総合", "グロース株のリスク選好"),
        "ru": ("Nasdaq Composite", "Аппетит к риску акций роста"),
        "pt": ("Nasdaq Composite", "Apetito de risco das ações de crescimento"),
    },
    "HYG": {
        "en": ("High Yield Bond ETF (HYG)", "Credit spread proxy; big drops = credit stress"),
        "zh": ("高收益债ETF(HYG)", "信用利差代理；大跌=信用压力"),
        "es": ("ETF de Renta Alta (HYG)", "Proxy del diferencial de crédito; caídas fuertes = estrés crediticio"),
        "fr": ("ETF Haut Rendement (HYG)", "Proxy des spreads de crédit ; fortes baisses = stress crédit"),
        "de": ("High-Yield-ETF (HYG)", "Kreditspread-Proxy; starke Rückgänge = Kreditstress"),
        "ja": ("ハイイールド債ETF（HYG）", "クレジットスプレッドの代理指標。急落＝信用ストレス"),
        "ru": ("ETF высокодоходных облигаций (HYG)", "Прокси кредитного спреда; резкие падения = кредитный стресс"),
        "pt": ("ETF de Alto Rendimento (HYG)", "Proxy do spread de crédito; quedas fortes = estresse de crédito"),
    },
    "CPI": {
        "en": ("US CPI", "Headline inflation; Fed target 2%"),
        "zh": ("美国CPI(城市消费者)", "通胀总口径；美联储目标 2%"),
        "es": ("IPC de EE.UU.", "Inflación general; objetivo de la Fed 2%"),
        "fr": ("IPC des États-Unis", "Inflation globale ; objectif de la Fed 2 %"),
        "de": ("US-Verbraucherpreisindex", "Gesamtinflation; Ziel der Fed 2 %"),
        "ja": ("米国CPI", "総合インフレ。FRBの目標は2%"),
        "ru": ("ИПЦ США", "Общая инфляция; цель ФРС 2%"),
        "pt": ("IPC dos EUA", "Inflação cheia; meta do Fed 2%"),
    },
    "PCE": {
        "en": ("US PCE Price Index", "The inflation gauge the Fed actually targets"),
        "zh": ("美国PCE物价指数", "美联储决策最看重的通胀口径"),
        "es": ("Índice de precios PCE de EE.UU.", "La medida de inflación que sigue la Fed"),
        "fr": ("Indice des prix PCE des États-Unis", "La mesure d'inflation suivie par la Fed"),
        "de": ("US-PCE-Preisindex", "Das Inflationsmaß, auf das die Fed schaut"),
        "ja": ("米国PCE物価指数", "FRBが最も重視するインフレ指標"),
        "ru": ("Индекс цен PCE США", "Инфляционный ориентир ФРС"),
        "pt": ("Índice de preços PCE dos EUA", "A inflação que o Fed acompanha"),
    },
    "PPI": {
        "en": ("US PPI (All Commodities)", "Upstream inflation; usually leads CPI by 1-3 months"),
        "zh": ("美国PPI(全部商品)", "上游通胀，通常领先CPI 1-3个月"),
        "es": ("IPP de EE.UU. (todos los bienes)", "Inflación aguas arriba; suele adelantar al IPC 1-3 meses"),
        "fr": ("IPP des États-Unis (tous biens)", "Inflation en amont ; devance l'IPC de 1 à 3 mois"),
        "de": ("US-Erzeugerpreise (alle Güter)", "Vorgelagerte Inflation; läuft dem CPI 1-3 Monate voraus"),
        "ja": ("米国PPI（全商品）", "川上のインフレ。通常CPIに1〜3か月先行"),
        "ru": ("ИЦП США (все товары)", "Инфляция на входе; обычно опережает ИПЦ на 1-3 месяца"),
        "pt": ("IPP dos EUA (todos os bens)", "Inflação a montante; antecipa o IPC em 1-3 meses"),
    },
    "NFP": {
        "en": ("US Nonfarm Payrolls", "Total employment (thousands); monthly change shown in notes"),
        "zh": ("美国非农就业人数", "就业总规模（千人）；月度新增见备注"),
        "es": ("Nóminas no agrícolas de EE.UU.", "Empleo total (miles); la variación mensual figura en notas"),
        "fr": ("Emplois non agricoles des États-Unis", "Emploi total (milliers) ; variation mensuelle en note"),
        "de": ("US-Arbeitsplätze außerhalb der Landwirtschaft", "Gesamtbeschäftigung (Tausend); Monatsveränderung in den Notizen"),
        "ja": ("米国非農業部門雇用者数", "雇用総数（千人）。月間増減は備考欄"),
        "ru": ("Число занятых в несельском хозяйстве США", "Общая занятость (тыс.); месячное изменение в примечании"),
        "pt": ("Payroll dos EUA", "Emprego total (milhares); variação mensal nas notas"),
    },
    "PMI": {
        "en": ("Manufacturing Index (NY Fed)", "ISM PMI has no free history; NY Fed Empire State used instead; above 0 = expansion"),
        "zh": ("制造业景气指数(纽约联储)", "ISM PMI 无免费历史源，改用纽约联储制造业指数；0上方=扩张"),
        "es": ("Índice manufacturero (Fed de NY)", "El ISM PMI no tiene historial gratuito; se usa el Empire State; >0 = expansión"),
        "fr": ("Indice manufacturier (Fed de NY)", "L'ISM PMI n'a pas d'historique gratuit ; Empire State utilisé ; >0 = expansion"),
        "de": ("Industrieindex (NY Fed)", "ISM PMI ohne freie Historie; stattdessen NY-Fed-Empire-State; >0 = Expansion"),
        "ja": ("製造業景気指数（NY連銀）", "ISM PMIは無料の履歴がないためNY連銀指数を代替。0超＝拡大"),
        "ru": ("Индекс производства (ФРБ Нью-Йорка)", "У ISM PMI нет бесплатной истории; используется Empire State; >0 = рост"),
        "pt": ("Índice industrial (Fed de NY)", "ISM PMI não tem histórico gratuito; usa-se Empire State; >0 = expansão"),
    },
    "CAPE": {
        "en": ("Shiller CAPE", "S&P 500 cyclically adjusted P/E (10Y); above 30 = expensive"),
        "zh": ("席勒CAPE(周期调整PE)", "标普10年周期调整市盈率；>30 偏贵"),
        "es": ("CAPE de Shiller", "PER ajustado cíclicamente del S&P 500 (10 años); >30 = caro"),
        "fr": ("CAPE de Shiller", "PER ajusté du cycle du S&P 500 (10 ans) ; >30 = cher"),
        "de": ("Shiller-CAPE", "Zyklisch adjustiertes KGV des S&P 500 (10J); >30 = teuer"),
        "ja": ("シラーCAPE", "S&P500の10年調整後PER。30超＝割高"),
        "ru": ("CAPE Шиллера", "Циклически скорректированный P/E S&P 500 (10 лет); >30 = дорого"),
        "pt": ("CAPE de Shiller", "P/L ajustado ao ciclo do S&P 500 (10 anos); >30 = caro"),
    },
}

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": UA})


# ----------------------------------------------------------------------------
# 日志
# ----------------------------------------------------------------------------
def log(msg: str) -> None:
    line = f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
        text = LOG.read_text(encoding="utf-8")
        if text.count("\n") > 200:
            LOG.write_text("\n".join(text.splitlines()[-100:]) + "\n", encoding="utf-8")
    except OSError:
        pass


# ----------------------------------------------------------------------------
# 数据抓取（Yahoo Finance 图表接口）
# ----------------------------------------------------------------------------
def fetch_yahoo(symbol: str, range_: str = None) -> dict:
    """返回 {日期字符串: 收盘值}；失败抛异常。
    range_=None 时抓取全量历史（1980 年起，各指标自可得日期起）。
    注意：不能用 range=max —— Yahoo 会静默降采样为月度数据；
    必须用 period1/period2 才能拿到真正的日线。"""
    if range_ is None:
        params = {"period1": "315532800",        # 1980-01-01
                  "period2": str(int(time.time())),
                  "interval": "1d"}
    else:
        params = {"range": range_, "interval": "1d"}
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{requests.utils.quote(symbol)}")
    last_err = None
    for attempt in range(3):
        try:
            r = SESSION.get(url, params=params, timeout=30)
            if r.status_code == 429:          # 限频，退避重试
                time.sleep(3 * (attempt + 1))
                continue
            r.raise_for_status()
            data = r.json()["chart"]["result"][0]
            ts = data.get("timestamp") or []
            closes = (data.get("indicators", {}).get("quote", [{}])[0]
                      .get("close") or [])
            out = {}
            for t, c in zip(ts, closes):
                if c is None:
                    continue
                tz_off = data.get("meta", {}).get("gmtoffset", -14400)
                d = dt.datetime.fromtimestamp(t + tz_off, dt.timezone.utc)
                out[d.strftime("%Y-%m-%d")] = float(c)
            return out
        except Exception as e:                # 网络抖动重试
            last_err = e
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Yahoo 抓取失败 {symbol}: {last_err}")


def fetch_sina_usdcny() -> dict:
    """备用数据源：新浪财经 美元兑人民币（仅当日值）。
    周末汇市休市，报价是周五收盘的陈旧值，不应创建新日期行。"""
    today = dt.date.today()
    if today.weekday() >= 5:            # 周六/周日不采当日值
        log("  今日为周末，跳过新浪当日汇率（避免生成非交易日脏行）")
        return {}
    r = SESSION.get("https://hq.sinajs.cn/list=fx_susdcny",
                    headers={"Referer": "https://finance.sina.com.cn"},
                    timeout=15)
    r.raise_for_status()
    text = r.content.decode("gbk", errors="ignore")
    parts = text.split('"')[1].split(",")
    return {today.strftime("%Y-%m-%d"): float(parts[3])}


# FRED 圣路易斯联储免费数据源（无需 API Key）：为历史较短的指标补长历史
FRED_MAP = {
    "US10Y": "DGS10",          # 10年期国债收益率
    "US2Y":  "DGS2",           # 2年期国债收益率，1976 年起
    "Brent": "DCOILBRENTEU",   # 布伦特现货，1987 年起
    "WTI":   "DCOILWTICO",     # WTI 现货，1986 年起
    "VIX":   "VIXCLS",         # VIX 恐慌指数，1990 年起
}


_HTTP_MODE = ["auto"]     # auto：requests 优先；curl：requests 失败过，直接走 curl


def http_get_bytes(url: str, timeout: int = 60) -> bytes:
    """取回远端内容：先 requests，失败自动回退到系统 curl。

    实测某些站点（FRED、耶鲁 Shiller）对 Python requests 的连接会读超时或被重置，
    而 curl 同样网络下 4 秒就能取回；Windows 10+ / macOS 都自带 curl，
    因此保留 curl 作为兜底通道，保证定时任务环境也能稳定取数。
    """
    last_err = None
    if _HTTP_MODE[0] != "curl":        # requests 失败过一次后，本轮直接走 curl
        try:
            r = SESSION.get(url, timeout=min(timeout, 12),
                            headers={"User-Agent": UA})
            r.raise_for_status()
            if r.content:
                return r.content
            raise RuntimeError("空响应")
        except Exception as e:
            last_err = e
            _HTTP_MODE[0] = "curl"
    exe = shutil.which("curl")
    if exe:
        try:
            # 单次尝试上限 m_arg：FRED 小文件 30s 足够，CAPE 大文件(1.6MB)随 timeout 放宽。
            # 注意：本机自带 curl 较旧，--retry/--retry-all-errors 在连接被重置(rc=56)时
            # 会反复重试并叠加超时、把整轮更新拖死，故这里只用单次直连 + -m 上限，
            # 真正的重试由 fetch_fred 自己的循环负责（每次换新连接，避开坏连接）。
            m_arg = max(30, timeout)
            p = subprocess.run([exe, "-sS", "-L", "--noproxy", "*",
                                "--connect-timeout", "15",
                                "-m", str(m_arg), "-A", UA, url],
                               capture_output=True, timeout=m_arg + 20)
            if p.returncode == 0 and p.stdout:
                return p.stdout
            last_err = (f"curl rc={p.returncode} "
                        f"{p.stderr[:120].decode('utf-8', 'ignore')}")
        except Exception as e:
            last_err = e
    raise RuntimeError(f"HTTP 获取失败 {url}: {last_err}")


def fetch_fred(series_id: str, cache_key: str = None,
               max_age_days: int = 3) -> dict:
    """从 FRED 下载 CSV，返回 {日期字符串: 值}（日频或月频）。

    带本地缓存：FRED 的利率/宏观序列变化很慢，缓存 3 天内有效，日常运行直接读缓存，
    避免反复联网（FRED 对该出口偶发连接重置/限频，且系统 curl 较旧不支持
    --retry-all-errors）。缓存键用 "fl:<series_id>"，与历史补齐用的 "col" 键分开，
    互不覆盖。缓存成功落地后，即使某天 FRED 临时不可达，也能用上次的成功数据。
    """
    url = (f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}")
    ck = cache_key or f"fl:{series_id}"
    cache = load_fred_cache()
    ent = cache.get(ck)
    if isinstance(ent, dict) and "ts" in ent and "data" in ent and ent["data"]:
        try:
            age = (dt.date.today() - dt.date.fromisoformat(ent["ts"])).days
        except Exception:
            age = 99
        if age <= max_age_days:
            return {str(k): float(v) for k, v in ent["data"].items()}
    time.sleep(2)                    # 轻微限速：避免连续请求触发 FRED 频率限制/连接重置
    last_err = None
    # 注意：本环境里 curl 子进程的 Schannel TLS 访问 FRED 会被对端重置(rc=56)，
    # 而 Python requests 经由沙箱代理(HTTPS_PROXY)可取到数据。但代理偶发断开连接，
    # 复用长连接池(SESSION)容易踩到「死连接」而报 ProxyError；故每次尝试用全新的
    # requests.get 建立新连接，并配合下面的多次重试，避开偶发的代理抖动。
    for attempt in range(4):          # FRED 偶发代理抖动，多试几次
        try:
            r = requests.get(url, timeout=30, headers={"User-Agent": UA})
            r.raise_for_status()
            text = r.text
            out = {}
            for line in text.strip().splitlines()[1:]:
                parts = line.split(",")
                if len(parts) == 2 and parts[1] not in ("", "."):
                    try:
                        out[parts[0]] = float(parts[1])
                    except ValueError:
                        pass
            if out:
                items = sorted(out.items())[-8000:]   # 仅缓存最近约 30 年，控文件体积
                cache[ck] = {"ts": dt.date.today().isoformat(),
                             "data": {k: out[k] for k, _ in items}}
                save_fred_cache(cache)
                return out
            raise RuntimeError("空数据")
        except Exception as e:
            last_err = e
            time.sleep(2)
    # 实时抓取失败：若本地有（可能过期的）缓存，降级使用，保证每日运行不丢数据；
    # 仅当完全无缓存时才报错。熔断逻辑仅作为历史补齐的全局冷却判断保留。
    if isinstance(ent, dict) and ent.get("data"):
        log(f"  [FRED] {series_id} 实时抓取失败，降级使用本地缓存（{ent.get('ts')}）")
        return {str(k): float(v) for k, v in ent["data"].items()}
    raise RuntimeError(f"FRED 抓取失败 {series_id}: {last_err}")


def month_end_date(ym: str) -> str:
    """'2026-08' → 该月最后一个工作日（周末回退到周五）。

    月度数据若落在 1 号，可能被「剔除周末脏行」逻辑误删，故统一挪到月末工作日。
    """
    y, m = int(ym[:4]), int(ym[5:7])
    last = calendar.monthrange(y, m)[1]
    d = dt.date(y, m, last)
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d.isoformat()


def to_month_end(series: dict) -> dict:
    """把 YYYY-MM(-01) 形式的月度键改成该月最后一个工作日。"""
    out = {}
    for d, v in series.items():
        try:
            out[month_end_date(d[:7])] = v
        except (ValueError, IndexError):
            out[d] = v
    return out


def fetch_cape() -> dict:
    """解析 Shiller 官方 ie_data.xls，返回 {YYYY-MM-01: CAPE}。

    Data 工作表：A 列日期（1871.01 这种「年.月」写法），M 列（索引 12）是 CAPE。
    官方文件停更于 2023-09，之后的数据从 cape_manual.csv 合并进来。

    取数策略：耶鲁 xls 约 1.6MB，弱网出口下下载极慢（~7KB/s，需 3+ 分钟，
    远超常规超时），而 CAPE 属历史序列、官方已停更，无需每日刷新。因此优先
    复用本地缓存 .cape.xls（完整约 1.6MB）；仅当缓存缺失或损坏时才联网下载
    并写回缓存，避免每次运行都卡在慢速下载上。
    """
    try:
        import xlrd
    except ImportError:
        raise RuntimeError("缺少 xlrd，无法解析 Shiller 数据"
                           "（pip install xlrd 到 .pylibs）")
    cached = BASE / ".cape.xls"
    src = None
    if cached.exists() and cached.stat().st_size >= 1_500_000:
        src = cached
        log("  CAPE 使用本地缓存 .cape.xls（跳过联网下载）")
    else:
        raw, last_err = None, None
        for attempt in range(2):
            try:
                raw = http_get_bytes(CAPE_URL, timeout=300)
                if len(raw) < 1_500_000:       # 完整文件约 1.6MB，太小即截断
                    raise RuntimeError(f"文件不完整（{len(raw)} 字节）")
                break
            except Exception as e:
                last_err = e
                raw = None
        if raw is None:
            raise RuntimeError(f"Shiller 数据下载失败: {last_err}")
        cached.write_bytes(raw)
        (BASE / ".cape_download.xls").write_bytes(raw)
        src = cached
        log(f"  CAPE 联网下载完成（{len(raw)//1024}KB，已缓存）")
    try:
        wb = xlrd.open_workbook(str(src))
        sh = wb.sheet_by_name("Data")
    except Exception as e:
        raise RuntimeError(f"Shiller 数据解析失败: {e}")
    out = {}
    for ri in range(8, sh.nrows):
        try:
            dv, cv = sh.cell_value(ri, 0), sh.cell_value(ri, 12)
        except IndexError:
            break
        if not isinstance(dv, float) or not isinstance(cv, float):
            continue
        year = int(dv)
        month = int(round((dv - year) * 100))
        if not 1 <= month <= 12:
            continue
        out[f"{year:04d}-{month:02d}-01"] = round(float(cv), 3)
    # 官方停更之后的数据：cape_manual.csv 手工补充
    if CAPE_MANUAL.exists():
        n = 0
        try:
            for line in CAPE_MANUAL.read_text(encoding="utf-8-sig").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.replace(",", " ").split()
                if len(parts) < 2:
                    continue
                out[f"{parts[0][:7]}-01"] = round(float(parts[1]), 3)
                n += 1
            if n:
                log(f"  CAPE 手工补充 {n} 个月（cape_manual.csv）")
        except Exception as e:
            log(f"  [警告] cape_manual.csv 读取失败：{e}")
    if not out:
        raise RuntimeError("Shiller 数据为空")
    return out


# FRED 熔断开关：[0]=本次运行是否已判定不可用
FRED_DOWN = [False]
FRED_COOLDOWN_DAYS = 3


FRED_CACHE = BASE / "fred_cache.json"


def load_fred_cache() -> dict:
    """读取本地 FRED 缓存 {列名: {日期: 值}}。"""
    if FRED_CACHE.exists():
        try:
            return json.loads(FRED_CACHE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_fred_cache(cache: dict) -> None:
    try:
        FRED_CACHE.write_text(json.dumps(cache), encoding="utf-8")
    except Exception as e:
        log(f"  [FRED] 缓存保存失败: {e}")


def backfill_from_fred(fetched: dict, errors: dict) -> None:
    """用 FRED 为历史短的指标补更早的历史（只补 Yahoo 未覆盖的更早日期）。
    成功后写入本地缓存 fred_cache.json —— 更早历史是静态数据，缓存一次永久有效，
    之后每日运行直接读缓存，不再访问 FRED（避免网络超时拖慢任务）。"""
    cache = load_fred_cache()
    for col, sid in FRED_MAP.items():
        # Yahoo 已给出足够长的历史时不再访问 FRED（FRED 偶发超时，省下等待）
        if len(fetched.get(col) or {}) >= FRED_SKIP_DAYS:
            continue
        if FRED_DOWN[0]:
            log(f"  [FRED] 已熔断，跳过 {col} 的历史补齐")
            continue
        need_fetch = col not in cache or not cache[col]
        if need_fetch:
            try:
                cache[col] = fetch_fred(sid)
                log(f"  [FRED] {col}: 下载成功 {len(cache[col])} 天")
                save_fred_cache(cache)
            except Exception as e:
                log(f"  [FRED] {col} 下载失败（不影响本次更新）: {e}")
                continue
        fred = cache[col]
        if col in fetched and fetched[col]:
            earliest = min(fetched[col])            # Yahoo 覆盖的最早日期
            older = {d: v for d, v in fred.items() if d < earliest}
            if older:
                fetched[col].update(older)
                log(f"  [FRED] {col}: 补充 {len(older)} 天更早历史"
                    f"（{min(older)} 起，共 {len(fetched[col])} 天）")
        elif col not in fetched:
            # Yahoo 完全失败时，用 FRED 全序列兜底（晚一天，可接受）
            fetched[col] = dict(fred)
            errors.pop(col, None)
            log(f"  [FRED] {col}: Yahoo 失败，改用 FRED 全序列共 {len(fred)} 天")


# ----------------------------------------------------------------------------
# 新浪财经备用源（中国大陆可直连；Yahoo 自 2021-11 起已在中国大陆停止服务）
# ----------------------------------------------------------------------------
# 列名 -> (类型, 新浪代码, 缩放系数)
#   gb  : 美股/ETF/指数  值=字段1  时间=字段3（北京时间）
#   hf  : 外盘期货       值=字段7（昨结算） 日期=字段11
#   fx  : 外汇/美元指数  值=字段3  日期=最后一个字段
#   znb : 全球指数       值=字段1  日期=字段6
SINA_SPEC = {
    "Brent":  ("hf",  "hf_OIL",    1.0),
    "WTI":    ("hf",  "hf_CL",     1.0),
    "Gold":   ("hf",  "hf_GC",     1.0),
    "Copper": ("hf",  "hf_HG",     0.01),   # 美分/磅 → 美元/磅
    "DXY":    ("fx",  "DINIW",     1.0),
    "USDCNY": ("fx",  "fx_susdcny", 1.0),
    "USDJPY": ("fx",  "fx_susdjpy", 1.0),
    "VIX":    ("znb", "znb_VIX",   1.0),
    "SPX":    ("gb",  "gb_$inx",   1.0),
    "NDX":    ("gb",  "gb_$ixic",  1.0),
    "HYG":    ("gb",  "gb_hyg",    1.0),
}
SINA_URL = "https://hq.sinajs.cn/list="


def _prev_weekday(d: dt.date) -> dt.date:
    """向前回退到最近的工作日（周一→上周五）。"""
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def _sina_trade_date(kind: str, raw_date: str, walk_back: bool) -> str | None:
    """把新浪给的时间换算成「美盘交易日」。
    gb/hf/znb 给的是北京时间，美盘收盘在北京次日凌晨，故需减 1 天。
    落在周末时：hf 回退到周五（昨结算本就属于上一场），其余直接跳过，
    避免用假日实时价覆盖掉上一个交易日的收盘价。"""
    try:
        d = dt.date.fromisoformat(raw_date[:10])
    except ValueError:
        return None
    if kind in ("gb", "hf", "znb"):
        d -= dt.timedelta(days=1)
        if d.weekday() >= 5:
            if walk_back:
                d = _prev_weekday(d)
            else:
                return None
    return d.isoformat()


def fetch_sina(col: str) -> dict:
    """取新浪单个指标的最新收盘，返回 {日期: 值}（通常只有 1 个点）。"""
    kind, code, scale = SINA_SPEC[col]
    r = SESSION.get(SINA_URL + code, timeout=15,
                    headers={"Referer": "https://finance.sina.com.cn",
                             "User-Agent": UA})
    r.raise_for_status()
    text = r.content.decode("gbk", errors="ignore")
    body = text.split('"')[1] if '"' in text else ""
    parts = body.split(",")
    if not body or len(parts) < 4:
        raise RuntimeError(f"新浪返回空数据 {code}")

    if kind == "gb":                       # 值=1，时间=3
        value, raw_date = parts[1], parts[3]
        walk = False
    elif kind == "hf":                     # 值=7（昨结算），日期=12
        value, raw_date = parts[7], parts[12]
        walk = True
    elif kind == "fx":                     # 值=3，日期=末字段
        value, raw_date = parts[3], parts[-1]
        walk = False
    else:                                  # znb：值=1，日期=6
        value, raw_date = parts[1], parts[6]
        walk = False

    try:
        v = float(value) * scale
    except ValueError:
        raise RuntimeError(f"新浪数值解析失败 {code}: {value!r}")
    if v <= 0:
        raise RuntimeError(f"新浪数值异常 {code}: {value!r}")

    date = _sina_trade_date(kind, raw_date.strip(), walk)
    if not date:
        raise RuntimeError(f"新浪日期非交易日 {code}: {raw_date!r}")
    return {date: v}


# Yahoo 一旦被判定为不可达就整体跳过，避免 13 个指标各等 3 次超时
YAHOO_BLOCKED = False


def fetch_all() -> tuple:
    """返回 ({列名: {日期: 值}}, {列名: 错误})。
    每个指标按 Yahoo → 新浪 → FRED 的顺序降级取数。"""
    global YAHOO_BLOCKED
    # 旧的「FRED 冷却熔断」已移除：fetch_fred 每次都尝试并带本地缓存/过期降级，
    # 单序列失败不再拖垮全部，也不会写入 __fred_down__ 标记。
    result, errors = {}, {}
    for col, sym, *_ in INDICATORS:
        last_e = ""
        ok = False

        # 0) 月度宏观基本面：Yahoo 没有这些序列，直接走 FRED
        if col in MACRO_FRED:
            if FRED_DOWN[0]:
                errors[col] = "FRED 冷却中，暂不可用"
                log(f"  跳过 {col:7s}：FRED 冷却中")
                continue
            try:
                result[col] = to_month_end(fetch_fred(MACRO_FRED[col]))
                log(f"  抓取成功 {col:7s} (FRED/{MACRO_FRED[col]}) "
                    f"共 {len(result[col])} 月")
                ok = True
            except Exception as e:
                last_e = e
            if ok:
                continue
            errors[col] = str(last_e)
            log(f"  抓取失败 {col:7s} (FRED/{MACRO_FRED[col]}): {last_e}")
            continue

        # 0b) CAPE：耶鲁 Shiller 官方文件（+ cape_manual.csv 手工补充）
        if col == "CAPE":
            try:
                result[col] = to_month_end(fetch_cape())
                _k = sorted(result[col])[-1]
                log(f"  抓取成功 CAPE   (Shiller) 共 {len(result[col])} 月"
                    f"，最新 {_k[:7]} = {result[col][_k]}")
                ok = True
            except Exception as e:
                last_e = e
            if ok:
                continue
            errors[col] = str(last_e)
            log(f"  抓取失败 CAPE   (Shiller): {last_e}")
            continue

        # 1) Yahoo（海外网络 / 代理可用时最完整）
        if not YAHOO_BLOCKED:
            for sym_try in ([sym, "HYG"] if col == "HYG" else [sym]):
                try:
                    result[col] = fetch_yahoo(sym_try)
                    log(f"  抓取成功 {col:7s} ({sym_try}) 共 {len(result[col])} 天")
                    ok = True
                    break
                except Exception as e:
                    last_e = e
                    if "403" in str(e) or "Forbidden" in str(e):
                        YAHOO_BLOCKED = True
                        log("  Yahoo 返回 403（中国大陆已不可用），"
                            "后续指标改用新浪/FRED")
                        break
        if ok:
            continue

        # 2) 新浪财经（当日收盘）
        if col in SINA_SPEC:
            try:
                result[col] = fetch_sina(col)
                d = next(iter(result[col]))
                log(f"  新浪备用源 {col:7s} ({SINA_SPEC[col][1]}) "
                    f"{d} = {result[col][d]}")
                ok = True
            except Exception as e:
                last_e = f"{last_e} / 新浪失败: {e}" if last_e else e
        if ok:
            continue

        # 3) FRED 全序列兜底
        if col in FRED_MAP and not FRED_DOWN[0]:
            try:
                result[col] = fetch_fred(FRED_MAP[col])
                log(f"  FRED 兜底 {col:7s} ({FRED_MAP[col]}) "
                    f"共 {len(result[col])} 天")
                ok = True
            except Exception as e:
                last_e = f"{last_e} / FRED 失败: {e}" if last_e else e
        if ok:
            continue

        errors[col] = str(last_e)
        log(f"  抓取失败 {col:7s} ({sym}): {last_e}")
    if not result:
        raise RuntimeError("所有指标抓取失败，请检查网络")
    return result, errors


# ----------------------------------------------------------------------------
# 数据质量层：异常值检测与剔除（P0）
# ----------------------------------------------------------------------------
def clean_series(col: str, series: dict) -> tuple:
    """异常值剔除，返回 (清洗后序列, 异常点列表)。
    规则1：数值超出合理区间 → 硬剔除（拦截列错位级别的脏数据）；
    规则2：与前一日 AND 后一日相比均为 >SPIKE_PCT% 的 V 型跳变 → 剔除
           （真实行情多为单边持续走势，双向 V 型跳变基本是坏点）。"""
    if len(series) < 3:
        return series, []
    lo, hi = VALID_RANGE.get(col, (None, None))
    anomalies = []
    dates = sorted(series)
    cleaned = {}
    prev = None
    for i, d in enumerate(dates):
        v = series[d]
        nxt = next((series[dates[j]] for j in range(i + 1, len(dates))
                    if series[dates[j]] is not None), None)
    bad = False
    if lo is not None and not (lo <= v <= hi):
        bad = True
    elif prev is not None and prev != 0 and col not in MONTHLY_COLS:
        # 月度宏观/估值序列（CPI/PCE/PPI/非农/CAPE/PMI）的环比跳变是真实经济
        # 波动（尤其 PMI 类调查指数月度大起大落很常见），不应套用日频的 V 型尖刺
        # 剔除逻辑，否则会误删大量合法点。月度序列仅做硬区间(VALID_RANGE)校验。
        chg = abs((v - prev) / prev * 100)
        fwd_bad = (nxt is not None and nxt != 0
                   and abs((nxt - v) / v * 100) > SPIKE_PCT)
        if chg > SPIKE_PCT and fwd_bad:
            bad = True
        if bad:
            anomalies.append((d, prev, v))
        else:
            cleaned[d] = v
            prev = v
    return cleaned, anomalies


def quality_layer(fetched: dict) -> dict:
    """对所有抓取序列执行异常检测；返回 {列名: 异常点列表}。"""
    qa = {}
    for col, series in fetched.items():
        cleaned, anomalies = clean_series(col, series)
        if anomalies:
            log(f"  [数据质量] {col}: 剔除 {len(anomalies)} 个异常点 "
                f"{[a[0] for a in anomalies]}")
            fetched[col] = cleaned
            qa[col] = [{"date": a[0], "prev": round(a[1], 4) if a[1] else a[1],
                        "value": a[2]} for a in anomalies]
    return qa


# ----------------------------------------------------------------------------
# 通用计算
# ----------------------------------------------------------------------------
def value_on_or_before(dates, merged, col, ref_date):
    """ref_date 当日或之前最近一个非空值。"""
    best = None
    for d in dates:
        if d <= ref_date:
            v = merged.get(d, {}).get(col)
            if v is not None:
                best = v
        else:
            break
    return best


def chg_pct(cur, old):
    if cur is None or old in (None, 0):
        return None
    return (cur - old) / old * 100


def month_ref(latest_d, days):
    return (dt.date.fromisoformat(latest_d) - dt.timedelta(days=days)).isoformat()


def rating_for(col, v, chg3, yoy=None):
    """返回 (状态文本, 十六进制颜色)。颜色表示压力/关注度而非好坏。"""
    if v is None:
        return ("-", "#888888")
    GREEN, YELLOW, RED = "#008000", "#BF8F00", "#C00000"
    # 通胀类：看同比（月频指标的自然日环比没有意义）
    if col in ("CPI", "PCE", "PPI"):
        if yoy is None:
            return ("-", "#888888")
        if yoy < 2:
            return ("低于目标/温和", GREEN)
        if yoy < 3.5:
            return ("略高于目标", YELLOW)
        return ("通胀偏高", RED)
    if col == "NFP":
        if chg3 is None:
            return ("-", "#888888")
        if chg3 > 0.4:
            return ("就业稳步增长", GREEN)
        if chg3 < 0:
            return ("就业收缩", RED)
        return ("增长放缓", YELLOW)
    if col == "PMI":
        if v > 10:
            return ("强劲扩张", GREEN)
        if v > 0:
            return ("温和扩张", YELLOW)
        if v > -10:
            return ("温和收缩", YELLOW)
        return ("显著收缩", RED)
    if col == "CAPE":
        if v < 18:
            return ("估值偏低", GREEN)
        if v < 26:
            return ("估值中性", YELLOW)
        if v < 32:
            return ("估值偏高", RED)
        return ("估值极贵", RED)
    if col in ("US10Y", "US2Y"):
        if v < 2.5: return ("低位/宽松", GREEN)
        if v < 4.5: return ("中性", YELLOW)
        return ("偏高/紧缩", RED)
    if col in ("Brent", "WTI"):
        if v < 60: return ("偏低", GREEN)
        if v < 90: return ("中性", YELLOW)
        return ("偏高", RED)
    if col == "DXY":
        if v < 90: return ("偏弱", GREEN)
        if v < 105: return ("中性", YELLOW)
        return ("偏强", RED)
    if col == "USDCNY":
        if v < 6.8: return ("人民币偏强", GREEN)
        if v < 7.2: return ("区间震荡", YELLOW)
        return ("人民币偏弱", RED)
    if col == "USDJPY":
        if v < 130: return ("低位", GREEN)
        if v < 155: return ("中性", YELLOW)
        return ("高位/干预风险", RED)
    if col == "VIX":
        if v < 15: return ("平静", GREEN)
        if v < 20: return ("正常", YELLOW)
        if v < 30: return ("紧张", RED)
        return ("恐慌", RED)
    if col in ("SPX", "NDX"):
        if chg3 is None: return ("-", "#888888")
        if chg3 > 3: return ("上行趋势", RED)      # 涨=红（中国惯例）
        if chg3 < -5: return ("回调", GREEN)
        return ("横盘", YELLOW)
    if col == "HYG":
        if chg3 is None: return ("-", "#888888")
        if chg3 < -3: return ("利差走扩", RED)
        if chg3 > 1: return ("信用宽松", GREEN)
        return ("平稳", YELLOW)
    if col == "Gold":
        if chg3 is None: return ("-", "#888888")
        if chg3 > 5: return ("强势上行", RED)
        if chg3 < -5: return ("走弱", GREEN)
        return ("高位震荡", YELLOW)
    return ("-", "#888888")


def compute_derived(dates, merged):
    """派生指标：返回 {key: {label, label_key, value, value_txt, note, note_key}}。
    label/note 为中文（Excel 用），label_key/note_key 为 i18n 键（HTML 用）。"""
    latest_d = dates[-1]

    def lv(col):
        return value_on_or_before(dates, merged, col, latest_d)

    def c(col, days):
        v = value_on_or_before(dates, merged, col, latest_d)
        o = value_on_or_before(dates, merged, col, month_ref(latest_d, days))
        return chg_pct(v, o)

    out = {}
    y10, y2 = lv("US10Y"), lv("US2Y")
    if y10 is not None and y2 is not None:
        spread = y10 - y2
        if spread > 0.5:
            note, nk = "曲线陡峭化，市场定价增长与通胀回升", "dv2_curve_steep"
        elif spread >= 0:
            note, nk = "曲线正常偏陡，衰退信号解除", "dv2_curve_normal"
        elif spread > -0.3:
            note, nk = "曲线轻度倒挂，边缘预警", "dv2_curve_mild"
        else:
            note, nk = "曲线倒挂，历史上常领先经济衰退", "dv2_curve_inv"
        out["curve"] = {"label": "期限利差 10Y-2Y", "label_key": "dv2_curve",
                        "value": round(spread, 3),
                        "value_txt": f"{spread:+.3f}%",
                        "note": note, "note_key": nk}
    brent, wti = lv("Brent"), lv("WTI")
    if brent is not None and wti is not None:
        sp = brent - wti
        wide = sp > 5
        out["oilspread"] = {"label": "Brent-WTI 价差",
                            "label_key": "dv2_oilspread",
                            "value": round(sp, 2),
                            "value_txt": f"${sp:.2f}",
                            "note": ("价差偏宽，美国原油出口/炼厂因素主导"
                                     if wide else "价差正常区间"),
                            "note_key": ("dv2_oil_wide" if wide
                                         else "dv2_oil_normal")}
    gold = lv("Gold")
    if gold is not None and brent is not None:
        ratio = gold / brent
        hi = ratio > 25
        out["goldoil"] = {"label": "金油比 (Gold÷Brent)",
                          "label_key": "dv2_goldoil",
                          "value": round(ratio, 1),
                          "value_txt": f"{ratio:.1f}",
                          "note": (">25 通常预示衰退/避险主导；<15 经济扩张主导"
                                   if hi else "比值处于常态区间"),
                          "note_key": ("dv2_go_high" if hi
                                       else "dv2_go_normal")}
    d3 = c("DXY", 90)
    if d3 is not None:
        txt = f"{d3:+.1f}%"
        if d3 > 2:
            note, nk = "美元3月走强，利空新兴市场与大宗", "dv2_dxy_up"
        elif d3 < -2:
            note, nk = "美元3月走弱，利多风险资产与黄金", "dv2_dxy_dn"
        else:
            note, nk = "美元区间震荡，方向待选择", "dv2_dxy_flat"
        out["dxymom"] = {"label": "美元动量 (DXY 3月)", "label_key": "dv2_dxy",
                         "value": round(d3, 2),
                         "value_txt": txt, "note": note, "note_key": nk}
    h3 = c("HYG", 90)
    if h3 is not None:
        if h3 < -3:
            note, nk = "高收益债走弱，信用利差走扩，风险偏好恶化", "dv2_cr_weak"
        elif h3 > 0:
            note, nk = "高收益债走强，信用环境宽松", "dv2_cr_strong"
        else:
            note, nk = "高收益债平稳，信用市场无压力信号", "dv2_cr_calm"
        out["credit"] = {"label": "信用状态 (HYG 3月)", "label_key": "dv2_credit",
                         "value": round(h3, 2),
                         "value_txt": f"{h3:+.1f}%", "note": note, "note_key": nk}
    return out


def build_regime(dates, merged):
    """规则引擎：宏观状态 5 维度 + 综合判断 + 分项解读段落。"""
    latest_d = dates[-1]

    def lv(col):
        return value_on_or_before(dates, merged, col, latest_d)

    def c(col, days):
        v = value_on_or_before(dates, merged, col, latest_d)
        o = value_on_or_before(dates, merged, col, month_ref(latest_d, days))
        return chg_pct(v, o)

    vix = lv("VIX"); spx_c3 = c("SPX", 90); spx_c1 = c("SPX", 30)
    ndx_c3 = c("NDX", 90); hyg_c3 = c("HYG", 90)
    y10 = lv("US10Y"); y10_c1 = c("US10Y", 30)
    dxy = lv("DXY"); dxy_c1 = c("DXY", 30)
    oil_c1 = c("Brent", 30); oil_c3 = c("Brent", 90)
    gold_c1 = c("Gold", 30); gold_c3 = c("Gold", 90)
    cny_c1 = c("USDCNY", 30)
    curve = (y10 - lv("US2Y")) if (y10 is not None and lv("US2Y") is not None) else None

    # --- Risk-On / Off ---
    score = 0
    if spx_c3 is not None:
        if spx_c3 > 2: score += 1
        if spx_c3 < -5: score -= 1
    if vix is not None:
        if vix < 15: score += 1
        elif vix > 25: score -= 1
    if ndx_c3 is not None and ndx_c3 > 3: score += 1
    if hyg_c3 is not None and hyg_c3 < -3: score -= 1
    if score >= 2:
        risk = ("risk_on", "green")
    elif score <= -1:
        risk = ("risk_off", "red")
    else:
        risk = ("risk_neu", "yellow")

    # --- 通胀压力 ---
    if (oil_c1 or 0) > 10 or (oil_c3 or 0) > 15:
        inflation = ("infl_up", "red")
    elif (oil_c3 or 0) < -5 and (gold_c3 or 0) < 0:
        inflation = ("infl_ease", "green")
    else:
        inflation = ("infl_mid", "yellow")

    # --- 利率压力 ---
    if y10 is None:
        rates = ("na_txt", "yellow")
    elif y10 >= 4.5:
        rates = ("rates_high", "red")
    elif y10 >= 4.0:
        rates = ("rates_mh", "yellow")
    else:
        rates = ("rates_ok", "green")

    # --- 美元 ---
    if dxy is None:
        usd = ("na_txt", "yellow")
    elif dxy >= 105 or (dxy_c1 or 0) > 1:
        usd = ("usd_strong", "red")
    elif dxy <= 95 or (dxy_c1 or 0) < -1:
        usd = ("usd_weak", "green")
    else:
        usd = ("usd_mid", "yellow")

    # --- 波动率 ---
    if vix is None:
        vol = ("na_txt", "yellow")
    elif vix < 15:
        vol = ("vol_low", "green")
    elif vix < 20:
        vol = ("vol_norm", "yellow")
    elif vix < 30:
        vol = ("vol_tense", "red")
    else:
        vol = ("vol_panic", "red")

    def item(label_key, level, state_key, **params):
        return {"key": label_key, "label_key": label_key,
                "label": tr(label_key, "zh"),
                "level": level, "state_key": state_key, "params": params,
                "text": tr(state_key, "zh", **params)}

    items = [
        item("lbl_risk", risk[1], risk[0]),
        item("lbl_infl", inflation[1], inflation[0]),
        item("lbl_rates", rates[1], rates[0],
             **({"y": f"{y10:.2f}"} if y10 is not None else {})),
        item("lbl_usd", usd[1], usd[0]),
        item("lbl_vol", vol[1], vol[0],
             **({"v": f"{vix:.1f}"} if vix is not None else {})),
    ]

    # --- 综合判断（bits 为 i18n 键，中文由 tr 渲染，供 Excel 使用） ---
    bits = []
    bits.append({"risk_on": "sb_risk_s", "risk_off": "sb_risk_w",
                 "risk_neu": "sb_risk_n"}[risk[0]])
    bits.append("sb_rate_hi" if rates[0] in ("rates_high", "rates_mh")
                else "sb_rate_ok")
    if inflation[0] == "infl_up":
        bits.append("sb_infl_up")
    elif inflation[0] == "infl_ease":
        bits.append("sb_infl_ease")
    if usd[0] == "usd_weak":
        bits.append("sb_usd_weak")
    elif usd[0] == "usd_strong":
        bits.append("sb_usd_strong")
    if vol[0] == "vol_low":
        bits.append("sb_vol_low")
    summary = (tr("sum_prefix", "zh")
               + " + ".join(tr(b, "zh") for b in bits)
               + tr("sum_suffix", "zh"))

    # --- 分项解读（语言中立的骨架：标题键 + 分段[键,参数] 列表，JS 端渲染） ---
    ins = []
    if y10 is not None:
        segs = []
        if y10_c1 is not None:
            segs.append(("ir_s1", {"y": f"{y10:.2f}", "c": f"{y10_c1:+.2f}"}))
        else:
            segs.append(("ir_s1b", {"y": f"{y10:.2f}"}))
        segs.append(("ir_hi" if y10 >= 4.5 else "ir_mid", {}))
        if curve is not None:
            segs.append(("ir_curve", {"c": f"{curve:+.2f}"}))
            segs.append(("ir_curve_inv" if curve < 0 else "ir_curve_ok", {}))
        ins.append({"t": "it_rates", "segs": segs})
    if dxy is not None:
        segs = [("iu_s1", {"d": f"{dxy:.2f}", "c": f"{dxy_c1:+.2f}"})
                if dxy_c1 is not None else ("iu_s1b", {"d": f"{dxy:.2f}"})]
        if (dxy_c1 or 0) < -0.5:
            segs.append(("iu_weak", {}))
        elif (dxy_c1 or 0) > 0.5:
            segs.append(("iu_strong", {}))
        else:
            segs.append(("iu_mid", {}))
        if cny_c1 is not None:
            segs.append(("iu_cny_dn" if cny_c1 < 0 else "iu_cny_up",
                         {"c": f"{abs(cny_c1):.2f}"}))
        ins.append({"t": "it_usd", "segs": segs})
    if vix is not None and spx_c3 is not None:
        lvl = ("iv_s1_low" if vix < 15
               else "iv_s1_norm" if vix < 20 else "iv_s1_high")
        segs = [(lvl, {"v": f"{vix:.1f}", "s": f"{spx_c3:+.1f}"})]
        if vix < 15 and spx_c3 > 0:
            segs.append(("iv_low_up", {}))
        elif vix >= 20:
            segs.append(("iv_high", {}))
        else:
            segs.append(("iv_mid", {}))
        if hyg_c3 is not None and hyg_c3 < -3:
            segs.append(("iv_hyg", {"h": f"{hyg_c3:+.1f}"}))
        ins.append({"t": "it_risk", "segs": segs})
    if oil_c1 is not None:
        segs = [("io_s1", {"c1": f"{oil_c1:+.1f}", "c3": f"{oil_c3:+.1f}"})]
        if oil_c1 > 8:
            segs.append(("io_hot", {}))
        elif (oil_c3 or 0) > 10:
            segs.append(("io_cool", {}))
        else:
            segs.append(("io_flat", {}))
        ins.append({"t": "it_infl", "segs": segs})
    if gold_c1 is not None:
        segs = [("ig_s1", {"c1": f"{gold_c1:+.1f}", "c3": f"{gold_c3:+.1f}"})]
        if gold_c1 > 0 and (dxy_c1 or 0) < 0:
            segs.append(("ig_up_wd", {}))
        elif gold_c1 > 0 and (dxy_c1 or 0) > 0:
            segs.append(("ig_up_sd", {}))
        elif gold_c1 < 0:
            segs.append(("ig_dn", {}))
        else:
            segs.append(("ig_flat", {}))
        ins.append({"t": "it_gold", "segs": segs})

    return {
        "items": items,
        "summary": summary,
        "sum_bits": bits,     # i18n 键列表（HTML 端按语言渲染）
        "insights": ins,      # 骨架 [{t, segs}, ...]（HTML 端按语言渲染）
        "time": f"{dt.datetime.now():%Y-%m-%d %H:%M}",
    }


# ----------------------------------------------------------------------------
# 宏观因子引擎（四因子 + Regime 象限 + 分歧检测 + 资产倾向 + 因子历史）
# ----------------------------------------------------------------------------
def build_series(dates, merged, col):
    """按 dates 顺序返回列值列表（缺失为 None）。"""
    return [merged.get(d, {}).get(col) for d in dates]


def mom_zscore(dates, merged, col, window=63):
    """近 window 个交易日的动量（%变化）相对全历史同一窗口动量分布的 Z-Score。"""
    vals = build_series(dates, merged, col)
    rets = []
    for i in range(window, len(vals)):
        a, b = vals[i], vals[i - window]
        if a is not None and b not in (None, 0):
            rets.append((a - b) / b * 100)
    if len(rets) < 60:
        return None
    cur = rets[-1]
    m = sum(rets) / len(rets)
    sd = (sum((x - m) ** 2 for x in rets) / len(rets)) ** 0.5
    if sd == 0:
        return None
    return (cur - m) / sd


def level_pct_rank(dates, merged, col):
    """当前水平在全历史中的百分位（0-100）。"""
    vals = [v for v in build_series(dates, merged, col) if v is not None]
    if len(vals) < 60:
        return None
    cur = vals[-1]
    return sum(1 for v in vals if v <= cur) / len(vals) * 100


def chg_over_trading_days(dates, merged, col, window=63):
    """近 window 个交易日的 %变化。"""
    vals = build_series(dates, merged, col)
    a, b = None, None
    for i in range(len(vals) - 1, -1, -1):
        if vals[i] is not None:
            a = vals[i]
            break
    if a is None:
        return None
    count = 0
    for i in range(len(vals) - 1, -1, -1):
        if vals[i] is not None:
            count += 1
            if count == window + 1:
                b = vals[i]
                break
    if b in (None, 0):
        return None
    return (a - b) / b * 100


def build_factors(dates, merged):
    """四因子评分：Growth / Inflation / Liquidity / Risk + Regime 象限。
    输出同时含中文字段（Excel 用）与 i18n 键（HTML 用）。"""
    # 各因子成分（列名, 方向 +1 正贡献 / -1 反向）
    comp = {
        "Growth":     [("Copper", 1), ("SPX", 1), ("NDX", 1)],
        "Inflation":  [("Brent", 1), ("Gold", 1)],
        "Liquidity":  [("US10Y", -1), ("US2Y", -1), ("DXY", -1)],
        "Risk":       [("SPX", 1), ("NDX", 1), ("VIX", -1), ("HYG", 1)],
    }
    label_key = {"Growth": "fc_growth", "Inflation": "fc_infl",
                 "Liquidity": "fc_liq", "Risk": "fc_risk"}
    factors = []
    fscore = {}
    for key, members in comp.items():
        zs = []
        for col, sign in members:
            z = mom_zscore(dates, merged, col)
            if z is not None:
                zs.append(sign * z)
        score = sum(zs) / len(zs) if zs else None
        fscore[key] = score
        if score is None:
            arrow, skey = "→", "fc_na"
        elif score > 0.5:
            arrow, skey = "↑", "fc_strong"
        elif score < -0.5:
            arrow, skey = "↓", "fc_weak"
        else:
            arrow, skey = "→", "fc_neu"
        factors.append({"key": key, "label": tr(label_key[key], "zh"),
                        "label_key": label_key[key], "arrow": arrow,
                        "score": round(score, 2) if score is not None else None,
                        "state_key": skey, "text": tr(skey, "zh")})
    # Regime 象限（Growth × Inflation）
    g = fscore.get("Growth")
    i = fscore.get("Inflation")
    if g is None or i is None:
        qk, nk = "qd_na", "qn_na"
    elif g > 0.5 and i > 0.5:
        qk, nk = "qd_refl", "qn_refl"
    elif g > 0.5 and i < -0.5:
        qk, nk = "qd_gold", "qn_gold"
    elif g < -0.5 and i > 0.5:
        qk, nk = "qd_stag", "qn_stag"
    elif g < -0.5 and i < -0.5:
        qk, nk = "qd_defl", "qn_defl"
    elif i > 0.5:
        qk, nk = "qd_irefl", "qn_irefl"
    elif i < -0.5:
        qk, nk = "qd_idis", "qn_idis"
    else:
        qk, nk = "qd_neu", "qn_neu"
    return {"factors": factors, "quadrant": tr(qk, "zh"),
            "quadrant_note": tr(nk, "zh"),
            "quadrant_key": qk, "quadrant_note_key": nk}


def build_divergences(dates, merged):
    """宏观分歧检测器。输出 i18n 键 + 预格式化参数（HTML 端按语言渲染）。"""
    def c3(col):
        return chg_over_trading_days(dates, merged, col)

    divs = []
    spx, ndx = c3("SPX"), c3("NDX")
    hyg, dxy = c3("HYG"), c3("DXY")
    gold, copper, brent = c3("Gold"), c3("Copper"), c3("Brent")
    if spx is not None and hyg is not None:
        if spx > 3 and hyg < -3:
            divs.append({"name_key": "dv_n_credit", "text_key": "dv_credit",
                         "params": {"s": f"{spx:+.1f}", "h": f"{hyg:+.1f}"}})
    if gold is not None and dxy is not None:
        if gold > 5 and dxy > 2:
            divs.append({"name_key": "dv_n_gold", "text_key": "dv_gold",
                         "params": {"g": f"{gold:+.1f}", "d": f"{dxy:+.1f}"}})
    if spx is not None and copper is not None:
        if spx > 3 and copper < -5:
            divs.append({"name_key": "dv_n_copper", "text_key": "dv_copper",
                         "params": {"s": f"{spx:+.1f}", "c": f"{copper:+.1f}"}})
    if ndx is not None and spx is not None:
        if abs(ndx - spx) > 8:
            lead = "@dv_lead_ndx" if ndx > spx else "@dv_lead_spx"
            divs.append({"name_key": "dv_n_style", "text_key": "dv_style",
                         "params": {"g": f"{ndx - spx:+.1f}", "lead": lead}})
    if brent is not None and spx is not None:
        if brent > 15 and spx > 3:
            divs.append({"name_key": "dv_n_oil", "text_key": "dv_oil",
                         "params": {"b": f"{brent:+.1f}"}})
    return divs


def build_asset_bias(factors):
    """宏观状态 → 资产倾向（非买卖信号）。输出 i18n 键（HTML 端按语言渲染）。"""
    f = {x["key"]: x["arrow"] for x in factors["factors"]}
    g, i, lq, r = f.get("Growth", "→"), f.get("Inflation", "→"), \
        f.get("Liquidity", "→"), f.get("Risk", "→")
    def bias(asset, cond_pos, cond_neg, note_p, note_n):
        if cond_pos: return (asset, "bb_pos", note_p)
        if cond_neg: return (asset, "bb_neg", note_n)
        return (asset, "bb_neu", None)
    out = []
    out.append(bias("ab_spx", r == "↑" and g == "↑", r == "↓",
                    "ab_spx_p", "ab_spx_n"))
    out.append(bias("ab_ndx", r == "↑" and g == "↑", r == "↓" or lq == "↓",
                    "ab_ndx_p", "ab_ndx_n"))
    out.append(bias("ab_gold", i == "↑" or lq == "↑", i == "↓" and lq == "↓",
                    "ab_gold_p", "ab_gold_n"))
    # 原油：通胀上行时「关注」，否则「中性」
    if i == "↑":
        out.append(("ab_oil", "bb_watch", "ab_oil_i"))
    else:
        out.append(("ab_oil", "bb_neu", "ab_oil_n"))
    # 美债久期：通胀上行使收益率上行压制价格，通胀回落利好久期
    if i == "↑":
        out.append(("ab_bond", "bb_neg", "ab_bond_i"))
    elif i == "↓":
        out.append(("ab_bond", "bb_pos", "ab_bond_d"))
    else:
        out.append(("ab_bond", "bb_neu", None))
    out.append(bias("ab_usd", lq == "↓", lq == "↑", "ab_usd_p", "ab_usd_n"))
    out.append(bias("ab_hy", r == "↑", r == "↓", "ab_hy_p", "ab_hy_n"))
    return [{"asset_key": a, "bias_key": bs, "note_key": n}
            for a, bs, n in out]


def build_regime_history(dates, merged, months=6):
    """过去 N 个月的 Growth / Inflation 因子状态（90日动量符号简化判定）。"""
    out = []
    today = dt.date.fromisoformat(dates[-1])
    # 每月取月末最后一个交易日
    month_last = {}
    for d in dates:
        month_last[d[:7]] = d          # 日期升序，后写覆盖 → 每月最后一天
    keys = sorted(month_last)[-months:]
    for mk in keys:
        d = month_last[mk]
        di = dates.index(d)
        sub_dates = dates[:di + 1]
        def m3(col):
            v = chg_over_trading_days(sub_dates, merged, col)
            return v
        cyc = [m3("SPX"), m3("NDX"), m3("Copper")]
        cyc = [x for x in cyc if x is not None]
        inf = [m3("Brent"), m3("Gold")]
        inf = [x for x in inf if x is not None]
        g = sum(cyc) / len(cyc) if cyc else 0
        i = sum(inf) / len(inf) if inf else 0
        ga = "↑" if g > 1.5 else ("↓" if g < -1.5 else "→")
        ia = "↑" if i > 1.5 else ("↓" if i < -1.5 else "→")
        out.append({"month": mk, "g": ga, "i": ia})
    return out


def build_zscores(dates, merged):
    """各指标动量 Z-Score + 历史百分位。"""
    labels = {i[0]: i[2] for i in INDICATORS}
    cols = ["US10Y", "US2Y", "Gold", "Copper", "Brent", "DXY", "VIX",
            "SPX", "NDX", "HYG"]
    out = []
    for col in cols:
        z = mom_zscore(dates, merged, col)
        p = level_pct_rank(dates, merged, col)
        if z is None and p is None:
            continue
        out.append({"col": col, "label": labels.get(col, col),
                    "z": round(z, 2) if z is not None else None,
                    "pct": round(p) if p is not None else None})
    return out


# ----------------------------------------------------------------------------
# Excel 更新
# ----------------------------------------------------------------------------
def d2s(v) -> str:
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def load_history(ws) -> dict:
    header = [c.value for c in ws[3]]
    cols = {get_column_letter(i + 1): name
            for i, name in enumerate(header) if name}
    data = {}
    for row in ws.iter_rows(min_row=4):
        d = d2s(row[0].value) if row[0].value else None
        if not d:
            continue
        rec = data.setdefault(d, {})
        for cell in row[1:]:
            name = cols.get(get_column_letter(cell.column))
            if name and cell.value is not None:
                rec[name] = float(cell.value)
    return data


def merge_history(ws, fetched: dict) -> tuple:
    old = load_history(ws)
    merged = {d: dict(rec) for d, rec in old.items()}
    added, updated = 0, 0
    for col, series in fetched.items():
        for d, v in series.items():
            if d not in merged:
                merged[d] = {col: v}
                added += 1
            else:
                if merged[d].get(col) != v:
                    updated += 1
                merged[d][col] = v
    dates = sorted(merged)
    # 剔除周末脏行：所有数据源均为交易日口径，周末行必为陈旧/错位数据
    weekend = [d for d in dates
               if dt.datetime.strptime(d, "%Y-%m-%d").weekday() >= 5]
    if weekend:
        for d in weekend:
            merged.pop(d, None)
        dates = sorted(merged)
        log(f"  剔除 {len(weekend)} 个周末非交易日脏行: {weekend}")
    # 数据质量：清洗历史遗留异常点（抓取覆盖不到的旧脏数据）
    removed = 0
    for col in VALID_RANGE:
        series = {d: rec.get(col) for d, rec in merged.items()
                  if rec.get(col) is not None}
        _, anomalies = clean_series(col, series)
        for d, _prev, _v in anomalies:
            if d in merged and col in merged[d]:
                del merged[d][col]
                removed += 1
    if removed:
        dates = sorted(merged)
        log(f"  [数据质量] 历史遗留异常点剔除 {removed} 个值")
    if len(dates) > HIST_KEEP:
        dates = dates[-HIST_KEEP:]
        log(f"  历史数据超过 {HIST_KEEP} 行，截断保留最新 {HIST_KEEP} 行")
    if ws.max_row > 3:
        ws.delete_rows(4, ws.max_row - 3)
    header = [i[0] for i in INDICATORS]
    # 重写表头行，确保列名与数据写入顺序严格一致（防止旧表头错位）
    ws.cell(row=3, column=1, value="Date")
    for j, name in enumerate(header, start=2):
        ws.cell(row=3, column=j, value=name)
    for i, d in enumerate(dates):
        r = 4 + i
        ws.cell(row=r, column=1, value=d)
        rec = merged[d]
        for j, name in enumerate(header, start=2):
            if name in rec:
                ws.cell(row=r, column=j, value=rec[name])
    ws.column_dimensions["A"].width = 12
    for j in range(2, 1 + len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 10
    return added, updated, dates, merged


def series_pairs(dates: list, merged: dict, col: str) -> list:
    """取某指标的有效 (日期, 值) 序列（升序，跳过空值）。"""
    out = []
    for d in dates:
        v = (merged.get(d) or {}).get(col)
        if isinstance(v, (int, float)):
            out.append((d, float(v)))
    return out


def recent_swings(pairs: list, pct: float = 0.05) -> tuple:
    """用 ZigZag（百分比反转确认）找最近一次的摆动高点与低点。

    返回 ((高点日期, 高点值), (低点日期, 低点值))，找不到则为 None。
    最近出现的那个摆动点决定当前阶段：
      最后是低点 → 上行段，参考「自低点涨幅」；
      最后是高点 → 下行段，参考「自高点跌幅」。
    """
    if len(pairs) < 5:
        return None, None
    pivots = []          # (日期, 值, 'H'/'L')
    trend = 1            # 1=上升段，-1=下降段
    ext = 0              # 当前段内极值下标
    for i in range(1, len(pairs)):
        v, ev = pairs[i][1], pairs[ext][1]
        if ev <= 0:      # 防御：油价曾出现负值，跳过比例判断
            ext = i
            continue
        if trend == 1:
            if v >= ev:
                ext = i                       # 创新高，继续跟踪
            elif v <= ev * (1 - pct):         # 回落达到阈值 → 确认前高
                pivots.append((pairs[ext][0], ev, "H"))
                trend, ext = -1, i
        else:
            if v <= ev:
                ext = i                       # 创新低，继续跟踪
            elif v >= ev * (1 + pct):         # 反弹达到阈值 → 确认前低
                pivots.append((pairs[ext][0], ev, "L"))
                trend, ext = 1, i
    hi = next((p for p in reversed(pivots) if p[2] == "H"), None)
    lo = next((p for p in reversed(pivots) if p[2] == "L"), None)
    return (hi[0], hi[1]) if hi else None, (lo[0], lo[1]) if lo else None


def swing_phase(hi, lo) -> str:
    """当前阶段：'up' 自低点反弹 / 'down' 自高点回落 / '-' 无法判断。"""
    if hi and lo:
        return "up" if lo[0] > hi[0] else "down"
    if hi:
        return "down"
    if lo:
        return "up"
    return "-"


def fmt_swing(pair, prec: int) -> str:
    """把摆动点格式化成『值｜月-日』，如 7,850.20｜08-14。"""
    if not pair:
        return "-"
    d, v = pair
    return f"{v:,.{prec}f}｜{d[2:]}"      # 带年份的 YY-MM-DD，避免跨年误读


def refresh_snapshot(ws, dates: list, merged: dict) -> None:
    """刷新最新看板：13 指标（值/近1月/近3月/前期高低点/状态/备注）
    + 派生指标 + 宏观状态。"""
    latest = dates[-1]
    m1, m3 = month_ref(latest, 30), month_ref(latest, 90)
    y1 = month_ref(latest, 365)
    labels = {ind[2]: ind for ind in INDICATORS}

    ws["A1"] = (f"关键宏观指标最新看板  |  更新时间："
                f"{dt.datetime.now():%Y-%m-%d %H:%M}")
    # 图例：说明前期高低点怎么读（原第 2 行为空行，直接复用）
    lg = ws.cell(row=2, column=1,
                 value="前期高点/低点＝最近一次 ZigZag 摆动转折；"
                       "黄底加粗那格是当前所处阶段"
                       "（上行看『距低点』涨幅，下行看『距高点』跌幅）")
    lg.font = Font(size=9, color="808080", italic=True)
    # 列：1指标 2最新值 3同比 4近1月(月度指标为环比) 5近3月
    #     6前期高点 7距高点 8前期低点 9距低点 10状态 11备注
    C_ST, C_NOTE = 10, 11
    hdr = ["指标", "最新值", "同比", "近1月变化", "近3月变化",
           "前期高点", "距高点", "前期低点", "距低点", "状态/区间", "备注"]
    for j, h in enumerate(hdr, start=1):
        c = ws.cell(row=3, column=j, value=h)
        c.font = Font(bold=True)

    # 旧数据区清理（含旧备注列残留）
    if ws.max_row > 3:
        ws.delete_rows(4, ws.max_row - 3)

    for i, (col, sym, label, note, prec, color, _grp) in enumerate(INDICATORS):
        r = 4 + i
        cur = value_on_or_before(dates, merged, col, latest)
        if cur is None:
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=C_ST, value="-")
            ws.cell(row=r, column=C_NOTE, value=note)
            continue
        if col in MONTHLY_COLS:
            # 月度指标：按「数据点个数」回看（1/3/12 个月前），而非自然日
            vals = [p[1] for p in series_pairs(dates, merged, col)]
            cur_d = series_pairs(dates, merged, col)[-1][0]
            prev1 = vals[-2] if len(vals) >= 2 else None
            prev3 = vals[-4] if len(vals) >= 4 else None
            base_y = vals[-13] if len(vals) >= 13 else None
        else:
            cur_d = latest
            prev1 = value_on_or_before(dates, merged, col, m1)
            prev3 = value_on_or_before(dates, merged, col, m3)
            base_y = value_on_or_before(dates, merged, col, y1)
        p1, p3 = chg_pct(cur, prev1), chg_pct(cur, prev3)
        py = chg_pct(cur, base_y)
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=round(cur, prec))
        ws.cell(row=r, column=3, value=f"{py:+.2f}%" if py is not None else "-")
        ws.cell(row=r, column=4,
                value=(f"环比{p1:+.2f}%" if col in MONTHLY_COLS
                       else f"{p1:+.2f}%") if p1 is not None else "-")
        ws.cell(row=r, column=5, value=f"{p3:+.2f}%" if p3 is not None else "-")
        st_txt, st_color = rating_for(col, cur, p3, py)
        c5 = ws.cell(row=r, column=C_ST, value=st_txt)
        c5.font = Font(color=st_color.lstrip("#"))
        # 非农额外标注月度新增；月度指标统一标注数据月份（发布有滞后）
        note_txt = note
        if col == "NFP" and prev1 is not None:
            note_txt = f"{note}；最新月度新增 {(cur - prev1) / 10:+.1f} 万人"
        if col in MONTHLY_COLS:
            note_txt = f"{note_txt}｜数据月份 {cur_d[:7]}"
        ws.cell(row=r, column=C_NOTE, value=note_txt)
        # 涨红跌绿
        for cell, p in ((ws.cell(row=r, column=3), py),
                        (ws.cell(row=r, column=4), p1),
                        (ws.cell(row=r, column=5), p3)):
            if p is not None:
                cell.font = Font(color="C00000" if p > 0 else "008000")

        # ---- 前期摆动高点 / 低点 ----
        pairs = series_pairs(dates, merged, col)
        hi, lo = recent_swings(pairs, SWING_PCT.get(col, SWING_PCT_DEFAULT))
        phase = swing_phase(hi, lo)
        ws.cell(row=r, column=6, value=fmt_swing(hi, prec))
        ws.cell(row=r, column=8, value=fmt_swing(lo, prec))
        ph = chg_pct(cur, hi[1]) if hi else None
        pl = chg_pct(cur, lo[1]) if lo else None
        c_hi = ws.cell(row=r, column=7,
                       value=f"{ph:+.2f}%" if ph is not None else "-")
        c_lo = ws.cell(row=r, column=9,
                       value=f"{pl:+.2f}%" if pl is not None else "-")
        # 当前所处阶段的那一格加粗+浅底，一眼看出该参照高点还是低点
        for cell, p, is_now in ((c_hi, ph, phase == "down"),
                                (c_lo, pl, phase == "up")):
            if p is None:
                continue
            cell.font = Font(bold=is_now,
                             color="C00000" if p > 0 else "008000")
            if is_now:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")

    # ---- 派生指标区 ----
    derived = compute_derived(dates, merged)
    r0 = 4 + len(INDICATORS) + 1
    c = ws.cell(row=r0, column=1, value="派生指标")
    c.font = Font(bold=True, size=12)
    for k, d in enumerate(derived.values()):
        r = r0 + 1 + k
        ws.cell(row=r, column=1, value=d["label"])
        ws.cell(row=r, column=2, value=d["value_txt"])
        ws.cell(row=r, column=C_NOTE, value=d["note"])

    # ---- 宏观状态区 ----
    regime = build_regime(dates, merged)
    r1 = r0 + len(derived) + 2
    c = ws.cell(row=r1, column=1, value="宏观状态")
    c.font = Font(bold=True, size=12)
    lvl_color = {"green": "008000", "yellow": "BF8F00", "red": "C00000"}
    for k, item in enumerate(regime["items"]):
        r = r1 + 1 + k
        ws.cell(row=r, column=1, value=item["label"])
        c2 = ws.cell(row=r, column=2, value=item["text"])
        c2.font = Font(color=lvl_color[item["level"]])
        ws.cell(row=r, column=C_NOTE, value=item["text"])
    r2 = r1 + len(regime["items"]) + 1
    ws.cell(row=r2, column=1, value="综合判断")
    c3 = ws.cell(row=r2, column=2, value=regime["summary"])
    c3.font = Font(bold=True)
    c3.alignment = Alignment(wrap_text=True)

    # ---- 宏观因子区（四因子 + Regime 象限） ----
    factors = build_factors(dates, merged)
    r3 = r2 + 2
    c = ws.cell(row=r3, column=1, value="宏观因子")
    c.font = Font(bold=True, size=12)
    for k, f in enumerate(factors["factors"]):
        r = r3 + 1 + k
        ws.cell(row=r, column=1, value=f["label"] + "因子")
        c2 = ws.cell(row=r, column=2,
                     value=f"{f['arrow']} {f['text']}"
                           + (f"（{f['score']:+.2f}σ）" if f["score"] is not None else ""))
        c2.font = Font(color=("C00000" if f["arrow"] == "↑"
                              else "008000" if f["arrow"] == "↓" else "BF8F00"))
    r4 = r3 + len(factors["factors"]) + 1
    ws.cell(row=r4, column=1, value="宏观象限")
    c4 = ws.cell(row=r4, column=2,
                 value=f"{factors['quadrant']} —— {factors['quadrant_note']}")
    c4.font = Font(bold=True)

    # ---- Z-Score 区 ----
    zscores = build_zscores(dates, merged)
    r5 = r4 + 2
    c = ws.cell(row=r5, column=1, value="动量Z-Score(3月)")
    c.font = Font(bold=True, size=12)
    for k, z in enumerate(zscores):
        r = r5 + 1 + k
        ws.cell(row=r, column=1, value=z["label"])
        zt = (f"{z['z']:+.2f}σ" if z["z"] is not None else "-")
        pt = (f"百分位 {z['pct']}%" if z["pct"] is not None else "")
        c2 = ws.cell(row=r, column=2, value=f"{zt}（{pt}）")
        if z["z"] is not None:
            c2.font = Font(color="C00000" if z["z"] > 2
                           else "008000" if z["z"] < -2 else "333333")

    # 列宽
    ws.column_dimensions["A"].width = 24
    for colL in ("B", "C", "D"):
        ws.column_dimensions[colL].width = 12
    ws.column_dimensions["E"].width = 18      # 前期高点
    ws.column_dimensions["F"].width = 10      # 距高点
    ws.column_dimensions["G"].width = 18      # 前期低点
    ws.column_dimensions["H"].width = 10      # 距低点
    ws.column_dimensions["I"].width = 14      # 状态/区间
    ws.column_dimensions["J"].width = 60      # 备注
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    # 清除可能残留的 A1:E1 合并（旧版）
    for rng in list(ws.merged_cells.ranges):
        if str(rng) == "A1:E1":
            ws.unmerge_cells("A1:E1")


def rebuild_chart_sheet(wb, dates: list, merged: dict) -> None:
    """重建『趋势图』工作表：4x3 原生折线图。"""
    if CHART_SHEET in wb.sheetnames:
        del wb[CHART_SHEET]
    ws = wb.create_sheet(CHART_SHEET)
    n = len(dates)
    last_row = 3 + n
    header = [i[0] for i in INDICATORS]

    ws["A1"] = "宏观指标趋势图（每日自动更新，随『历史数据』刷新）"
    ws["A1"].font = Font(bold=True, size=14)
    ws.sheet_view.showGridLines = False

    # 图表锚点网格：4 列，行数随指标数量自动扩展（避免 INDICATORS 增减后越界）
    n_cols = 4
    col0s = ("A", "K", "U", "AG")
    step = 17
    n_rows = (len(INDICATORS) + n_cols - 1) // n_cols
    anchors = [f"{col0s[c]}{3 + r * step}"
               for r in range(n_rows) for c in range(n_cols)]
    for idx, (col, sym, label, _, _, color, _grp) in enumerate(INDICATORS):
        if col not in header:
            continue
        col_idx = header.index(col) + 2
        ch = LineChart()
        ch.title = label
        ch.style = 2
        ch.height, ch.width = 8.2, 15.5
        ch.legend = None
        data_ref = Reference(wb[HIST_SHEET], min_col=col_idx, min_row=3,
                             max_row=last_row)
        cats_ref = Reference(wb[HIST_SHEET], min_col=1, min_row=4,
                             max_row=last_row)
        ch.add_data(data_ref, titles_from_data=True)
        ch.set_categories(cats_ref)
        s = ch.series[0]
        s.graphicalProperties.line.solidFill = color.lstrip("#")
        s.graphicalProperties.line.width = 16000
        s.smooth = False
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.x_axis.tickLblSkip = max(1, n // 8)
        ch.x_axis.tickMarkSkip = max(1, n // 8)
        ch.x_axis.txPr = None
        ws.add_chart(ch, anchors[idx])
    log(f"  趋势图工作表已重建（{n} 个交易日，{len(INDICATORS)} 张图表）")


# ----------------------------------------------------------------------------
# HTML 交互仪表盘 v2（四层决策结构）
# ----------------------------------------------------------------------------
ECHARTS_CDN = "https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"

HTML_TMPL = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Macro Decision Dashboard</title>
<script src="__CDN__"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", "Microsoft YaHei", "PingFang SC", "Hiragino Sans", sans-serif;
         background: #f5f6f8; color: #222; padding: 18px; max-width: 1400px;
         margin: 0 auto; }
  h1 { font-size: 18px; margin-bottom: 2px; }
  .sub { color: #888; font-size: 12px; margin-bottom: 14px; display: flex;
         gap: 14px; flex-wrap: wrap; align-items: center; }
  .dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%;
         margin-right: 4px; vertical-align: 1px; }
  .dot.green { background: #2e9e44; } .dot.yellow { background: #d9a400; }
  .dot.red { background: #d43a3a; }
  .evalctrl { display: inline-flex; align-items: center; gap: 6px; }
  .evalctrl input[type="date"] { font-size: 12px; padding: 2px 6px;
    border: 1px solid #d5d9e0; border-radius: 6px; color: #3756a4;
    font-weight: 600; background: #fff; }
  .evalctrl button { font-size: 11px; padding: 2px 10px; cursor: pointer;
    border: 1px solid #d5d9e0; border-radius: 6px; background: #fff;
    color: #666; }
  .evalctrl button:hover { background: #f0f3f8; }
  .langctrl { display: inline-flex; align-items: center; gap: 6px; }
  .langctrl select { font-size: 12px; padding: 2px 6px; cursor: pointer;
    border: 1px solid #d5d9e0; border-radius: 6px; background: #fff;
    color: #3756a4; font-weight: 600; }

  /* 第一层：宏观状态横幅 */
  .regime { background: #fff; border-radius: 12px; padding: 14px 18px;
            box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 14px; }
  .regime h2 { font-size: 14px; margin-bottom: 10px; color: #555; }
  .chips { display: flex; gap: 10px; flex-wrap: wrap; }
  .chip { border-radius: 8px; padding: 8px 14px; min-width: 120px;
          border: 1px solid #e5e7eb; }
  .chip .k { font-size: 11px; color: #888; }
  .chip .v { font-size: 15px; font-weight: 700; margin-top: 2px; }
  .chip.green .v { color: #2e9e44; } .chip.yellow .v { color: #c88a00; }
  .chip.red .v { color: #d43a3a; }
  .verdict { margin-top: 10px; font-size: 13px; background: #f8f9fb;
             border-left: 3px solid #3756a4; padding: 8px 12px;
             border-radius: 0 6px 6px 0; color: #333; }

  /* 第一层B：宏观象限 + 四因子 */
  .quad { margin-top: 10px; display: flex; gap: 10px; align-items: center;
          flex-wrap: wrap; }
  .quad .q { font-size: 15px; font-weight: 700; color: #fff;
             background: #3756a4; padding: 6px 16px; border-radius: 8px; }
  .quad .qn { font-size: 12px; color: #666; }
  .factor { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }
  .fchip { flex: 1; min-width: 150px; border: 1px solid #e5e7eb;
           border-radius: 8px; padding: 8px 12px; background: #fafbfc; }
  .fchip .k { font-size: 11px; color: #888; }
  .fchip .v { font-size: 16px; font-weight: 700; margin-top: 2px; }
  .fchip .s { font-size: 11px; color: #999; }
  .fchip .v.up { color: #2e9e44; } .fchip .v.dn { color: #d43a3a; }
  .fchip .v.fl { color: #c88a00; }

  /* Z-Score 行 */
  .zbar { background: #fff; border-radius: 10px; padding: 12px 16px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 14px; }
  .zbar h2 { font-size: 14px; margin-bottom: 6px; }
  .zbar .hint { font-size: 11px; color: #999; margin-bottom: 8px; }
  .zrow { display: flex; gap: 8px; flex-wrap: wrap; }
  .zitem { border: 1px solid #eceef2; border-radius: 6px; padding: 5px 10px;
           font-size: 11px; min-width: 108px; }
  .zitem b { font-size: 13px; display: block; }

  /* 分歧检测 */
  .diverge { background: #fff8f0; border: 1px solid #f0d9b5; border-radius: 10px;
             padding: 12px 16px; margin-bottom: 14px; }
  .diverge h2 { font-size: 14px; margin-bottom: 8px; color: #9a6b1f; }
  .diverge .item { font-size: 13px; line-height: 1.6; margin-bottom: 6px;
                   color: #5a4630; }
  .diverge .item b { color: #b05e19; }

  /* Regime 历史 */
  .hist { background: #fff; border-radius: 10px; padding: 12px 16px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 14px; }
  .hist h2 { font-size: 14px; margin-bottom: 8px; }
  table.htab { border-collapse: collapse; font-size: 12px; }
  table.htab th, table.htab td { border: 1px solid #eceef2; padding: 5px 12px;
                                  text-align: center; }
  table.htab th { background: #f8f9fb; }
  .g-up { color: #2e9e44; font-weight: 700; } .g-dn { color: #d43a3a; font-weight: 700; }
  .g-fl { color: #c88a00; }

  /* 资产倾向 */
  .bias { background: #fff; border-radius: 10px; padding: 14px 16px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 14px;
          overflow-x: auto; }
  .bias h2 { font-size: 14px; margin-bottom: 4px; }
  .bias .hint { font-size: 11px; color: #999; margin-bottom: 8px; }
  table.btab { border-collapse: collapse; font-size: 12px; }
  table.btab th, table.btab td { border: 1px solid #eceef2; padding: 5px 14px; }
  table.btab th { background: #f8f9fb; }

  /* 相关性窗口切换 */
  .tabs { display: flex; gap: 6px; margin-bottom: 8px; }
  .tab { font-size: 12px; padding: 4px 14px; border: 1px solid #d5d9e0;
         border-radius: 6px; cursor: pointer; background: #fff; color: #555; }
  .tab.on { background: #3756a4; color: #fff; border-color: #3756a4; }

  /* 分组标题 */
  .grp { font-size: 14px; font-weight: 700; margin: 16px 0 8px; color: #333; }

  /* 第二层：指标卡 */
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
          gap: 12px; }
  .card { background: #fff; border-radius: 10px; padding: 12px 14px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); min-width: 0; }
  .card h3 { font-size: 13px; font-weight: 600; display: flex;
             justify-content: space-between; align-items: baseline;
             white-space: nowrap; overflow: hidden; }
  .kpi { font-size: 19px; font-weight: 700; margin: 2px 0 6px; }
  .chg { font-size: 11px; margin-left: 8px; font-weight: 400; }
  .up { color: #c00000; }  .down { color: #008000; }  .flat { color: #888; }
  /* 前期摆动高点(H)/低点(L)：黄底加粗＝当前所处阶段 */
  .swings { display: flex; flex-wrap: wrap; gap: 10px; margin: 0 0 6px;
            font-size: 11px; color: #666; }
  .swings .sw { padding: 1px 6px; border-radius: 3px; white-space: nowrap;
                background: #f6f6f6; }
  .swings .sw.now { background: #FFF2CC; font-weight: 700; color: #333; }
  .swings .sw .tag { font-weight: 700; color: #999; margin-right: 3px; }
  .swings .sw.now .tag { color: #B8860B; }
  .swings .sw .dt { color: #999; margin-left: 3px; }
  .chart { width: 100%; height: 190px; min-width: 0; position: relative;
           overflow: hidden; }
  .chart div, .chart canvas { max-width: 100% !important; }

  /* 每张图的时间范围快捷按钮 */
  .ranges { display: flex; flex-wrap: wrap; gap: 3px; margin: 0 0 5px; }
  .ranges button { font-size: 10px; padding: 1px 6px; border: 1px solid #d5d9e0;
                   border-radius: 4px; background: #fff; color: #777;
                   cursor: pointer; line-height: 1.5; }
  .ranges button:hover { background: #f0f3f8; }
  .ranges button.on { background: #3756a4; color: #fff; border-color: #3756a4; }

  /* 第三层：派生指标 + 相关性 */
  .derived { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
             gap: 10px; margin-bottom: 14px; }
  .dbox { background: #fff; border-radius: 10px; padding: 10px 14px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); }
  .dbox .k { font-size: 11px; color: #888; }
  .dbox .v { font-size: 17px; font-weight: 700; margin: 2px 0; }
  .dbox .n { font-size: 11px; color: #666; line-height: 1.4; }
  .corr { background: #fff; border-radius: 10px; padding: 14px 16px;
          box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 14px;
          overflow-x: auto; }
  .corr h2 { font-size: 14px; margin-bottom: 4px; }
  .corr .hint { font-size: 11px; color: #999; margin-bottom: 8px; }
  table.ctab { border-collapse: collapse; font-size: 11px; }
  table.ctab th, table.ctab td { border: 1px solid #eceef2; padding: 4px 9px;
                                  text-align: center; }
  table.ctab th { background: #f8f9fb; font-weight: 600; }
  table.ctab td.lab { font-weight: 600; background: #f8f9fb; }

  /* 第四层：自动解读 */
  .insight { background: #fff; border-radius: 10px; padding: 16px 18px;
             box-shadow: 0 1px 4px rgba(0,0,0,.08); }
  .insight h2 { font-size: 14px; margin-bottom: 10px; }
  .insight .item { font-size: 13px; line-height: 1.7; color: #333;
                   margin-bottom: 8px; }
  .insight .item b { color: #3756a4; }
</style>
</head>
<body>
<h1>📊 <span data-ui="ui_title">Macro Decision Dashboard</span></h1>
<div class="sub">
  <span><span data-ui="ui_dstat_label">Data status:</span> <span id="dstat"></span></span>
  <span><span data-ui="ui_last_update">Last updated:</span> __TIME__</span>
  <span><span data-ui="ui_data_source">Data source:</span> Yahoo Finance</span>
  <span data-ui="ui_color_note">Red = up, Green = down</span>
  <span class="langctrl"><span data-ui="ui_lang_label">Language</span>
    <select id="langSel"></select></span>
  <span class="evalctrl"><span data-ui="ui_eval_label">Evaluation start date</span>
    <input type="date" id="evalStart" value="__EVAL12__" min="1980-01-01"
           max="__EVALMAX__">
    <button id="evalReset" data-ui="ui_reset">Reset</button>
  </span>
</div>

<div class="regime">
  <h2 data-ui="ui_sec_regime">🌎 Current Macro State</h2>
  <div class="chips" id="chips"></div>
  <div class="quad">
    <span class="q" id="quad"></span>
    <span class="qn" id="quadNote"></span>
  </div>
  <div class="factor" id="factors"></div>
  <div class="verdict" id="verdict"></div>
</div>

<div class="diverge" id="divbox" style="display:none">
  <h2 data-ui="ui_sec_diverge">⚠️ Macro Divergence Alerts</h2>
  <div id="divitems"></div>
</div>

<div id="cards"></div>

<div class="grp" data-ui="ui_sec_derived">🧮 Derived Indicators</div>
<div class="derived" id="derived"></div>

<div class="zbar">
  <h2 data-ui="ui_sec_zscore">📐 Momentum Z-Score &amp; Historical Percentile</h2>
  <div class="hint" data-ui="ui_zscore_hint">Z-Score hint</div>
  <div class="zrow" id="zrow"></div>
</div>

<div class="corr">
  <h2 data-ui="ui_sec_corr">🔗 Asset Return Correlation Matrix</h2>
  <div class="hint" data-ui="ui_corr_hint">Correlation hint</div>
  <div class="tabs">
    <button class="tab" data-n="30" data-ui="ui_t30">30D</button>
    <button class="tab on" data-n="90" data-ui="ui_t90">90D</button>
    <button class="tab" data-n="365" data-ui="ui_t1y">1Y</button>
  </div>
  <div id="corrtab"></div>
</div>

<div class="hist">
  <h2 data-ui="ui_sec_hist">🕰 Macro Regime History (Last 6 Months)</h2>
  <div id="histtab"></div>
</div>

<div class="bias">
  <h2 data-ui="ui_sec_bias">⚖️ Asset Bias (Regime Bias)</h2>
  <div class="hint" data-ui="ui_bias_hint">Bias hint</div>
  <div id="biastab"></div>
</div>

<div class="insight">
  <h2 data-ui="ui_sec_insight">🧠 Today's Macro Read</h2>
  <div id="insights"></div>
</div>

<script>
const DATA = __DATA__;
const REGIME = __REGIME__;
const DERIVED = __DERIVED__;
const STATUS = __STATUS__;
const FACTORS = __FACTORS__;
const ZSCORES = __ZSCORES__;
const DIVERG = __DIVERG__;
const BIAS = __BIAS__;
const RHIST = __HISTORY__;

// ---------- 多语言（i18n）----------
const L10N = __L10N__;            // {key: [en, zh, es, fr, de, ja, ru, pt]}
const IND_L10N = __IND_L10N__;    // {col: {lang: [label, desc]}}
const LANG_LIST = __LANGS__;
const LANG_NAMES = __LANG_NAMES__;
const LANG_IDX = {};
LANG_LIST.forEach((l, i) => { LANG_IDX[l] = i; });
const DEFAULT_LANG = "__DEFLANG__";
let LANG = DEFAULT_LANG;
try {
  const v = localStorage.getItem("macroLang");
  if (v && LANG_IDX[v] != null) LANG = v;
} catch (e) {}

// 模板渲染：占位符 {x}；参数值以 "@" 开头表示嵌套翻译键
function t(key, params) {
  const v = L10N[key];
  if (!v) return key;
  let s = v[LANG_IDX[LANG]] || v[0];
  if (params) {
    for (const k in params) {
      let val = params[k];
      if (typeof val === "string" && val.charAt(0) === "@") val = t(val.slice(1));
      s = s.split("{" + k + "}").join(val);
    }
  }
  return s;
}
function indLabel(col) {
  const e = IND_L10N[col];
  return (e && e[LANG] && e[LANG][0]) || col;
}
function indDesc(col) {
  const e = IND_L10N[col];
  return (e && e[LANG] && e[LANG][1]) || "";
}
// 静态界面文字（data-ui 属性）
function applyStaticTexts() {
  document.title = "📊 " + t("ui_title");
  document.documentElement.lang = LANG;
  document.querySelectorAll("[data-ui]").forEach(el => {
    el.textContent = t(el.dataset.ui);
  });
  const rb = document.getElementById("evalReset");
  if (rb) rb.title = t("ui_reset_tip");
}
// 语言选择器：切换后保存并整页刷新（数据已内嵌，刷新零开销）
(function initLangSel() {
  const sel = document.getElementById("langSel");
  LANG_LIST.forEach(l => {
    const o = document.createElement("option");
    o.value = l;
    o.textContent = LANG_NAMES[l] || l;
    if (l === LANG) o.selected = true;
    sel.appendChild(o);
  });
  sel.addEventListener("change", () => {
    try { localStorage.setItem("macroLang", sel.value); } catch (e) {}
    location.reload();
  });
})();
applyStaticTexts();

// ---------- 数据状态 ----------
(function () {
  const el = document.getElementById("dstat");
  if (STATUS.level === "live")
    el.innerHTML = '<span class="dot green"></span>' + t("ui_live");
  else if (STATUS.level === "partial")
    el.innerHTML = '<span class="dot yellow"></span>' + t("ui_partial", {t: STATUS.text});
  else
    el.innerHTML = '<span class="dot yellow"></span>' + t("ui_cache");
})();

// ---------- 第一层：宏观状态 ----------
(function () {
  const chips = document.getElementById("chips");
  REGIME.items.forEach(it => {
    const label = (it.label_key || it.key) ? t(it.label_key || it.key) : it.label;
    const text = it.state_key ? t(it.state_key, it.params) : it.text;
    const d = document.createElement("div");
    d.className = "chip " + it.level;
    d.innerHTML = `<div class="k">${label}</div><div class="v">${text}</div>`;
    d.title = text;
    chips.appendChild(d);
  });
  const summary = REGIME.sum_bits
    ? t("sum_prefix") + REGIME.sum_bits.map(b => t(b)).join(" + ") + t("sum_suffix")
    : REGIME.summary;
  document.getElementById("verdict").textContent = t("ui_v_prefix") + summary;
  document.getElementById("quad").textContent = FACTORS.quadrant_key
    ? t(FACTORS.quadrant_key) : FACTORS.quadrant;
  document.getElementById("quadNote").textContent = FACTORS.quadrant_note_key
    ? t(FACTORS.quadrant_note_key) : FACTORS.quadrant_note;
  const froot = document.getElementById("factors");
  FACTORS.factors.forEach(f => {
    const cls = f.arrow === "↑" ? "up" : (f.arrow === "↓" ? "dn" : "fl");
    const name = f.label_key ? t(f.label_key) : f.label;
    const state = f.state_key ? t(f.state_key) : f.text;
    const el = document.createElement("div");
    el.className = "fchip";
    el.innerHTML = `<div class="k">${t("ui_f_k", {n: name})}</div>
      <div class="v ${cls}">${f.arrow} ${state}</div>
      <div class="s">${f.score == null ? "" : t("ui_f_score", {s: f.score.toFixed(2)})}</div>`;
    froot.appendChild(el);
  });
})();

// ---------- 宏观分歧 ----------
(function () {
  if (!DIVERG || !DIVERG.length) return;
  document.getElementById("divbox").style.display = "";
  const root = document.getElementById("divitems");
  DIVERG.forEach(d => {
    const name = d.name_key ? t(d.name_key) : d.name;
    const text = d.text_key ? t(d.text_key, d.params) : d.text;
    const p = document.createElement("div");
    p.className = "item";
    p.innerHTML = `<b>【${name}】</b>${text}`;
    root.appendChild(p);
  });
})();

// ---------- Z-Score 行 ----------
(function () {
  const root = document.getElementById("zrow");
  ZSCORES.forEach(z => {
    const zc = z.z == null ? "#888" : (z.z > 2 ? "#c00000" : (z.z < -2 ? "#008000" : "#333"));
    const el = document.createElement("div");
    el.className = "zitem";
    el.innerHTML = `${indLabel(z.col)}<b style="color:${zc}">${z.z == null ? "-" : (z.z > 0 ? "+" : "") + z.z.toFixed(2) + "σ"}</b>
                    <span style="color:#999">${z.pct == null ? "-" : t("ui_z_pctile", {p: z.pct})}</span>`;
    root.appendChild(el);
  });
})();

// ---------- 第二层：分组指标卡 ----------
const charts = [];

// 评估起始日：默认「最近 12 个月」（滚动窗口，每次打开自动顺延），
// 用户可自定义并记忆（localStorage 持久化）
function defaultEvalStart() {
  const d = new Date();
  const day = d.getDate();
  d.setDate(1);                       // 先归到 1 号，避免 3/31 减一月变成 3/3
  d.setMonth(d.getMonth() - 12);
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, lastDay));
  const p = n => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
}
let EVAL_START = defaultEvalStart();
try {
  // v2：默认值改为滚动 12 个月，旧版本写死的日期不再沿用
  const saved = localStorage.getItem("macroEvalStart2");
  if (saved && /^\d{4}-\d{2}-\d{2}$/.test(saved)) EVAL_START = saved;
  localStorage.removeItem("macroEvalStart");
} catch (e) {}

// 序列中第一个 >= 评估起始日的索引（早于序列起点则 0）
function idxOnOrAfter(ind) {
  for (let i = 0; i < ind.dates.length; i++) {
    if (ind.dates[i] >= EVAL_START) return i;
  }
  return 0;
}
// 起始日（或之后首个有值日）到当前的涨跌幅 %，返回 {pct, baseDate}
function sincePct(ind) {
  let i0 = idxOnOrAfter(ind);
  let base = null, baseDate = null;
  for (let i = i0; i < ind.values.length; i++) {
    if (ind.values[i] != null) { base = ind.values[i]; baseDate = ind.dates[i]; break; }
  }
  let cur = null;
  for (let i = ind.values.length - 1; i >= 0; i--) {
    if (ind.values[i] != null) { cur = ind.values[i]; break; }
  }
  if (base == null || cur == null || base === 0) return null;
  return { pct: (cur - base) / base * 100, baseDate: baseDate };
}
function zoomStartPct(ind) {
  return Math.min(100, 100 * idxOnOrAfter(ind) / ind.dates.length);
}

// ---------- 每张图的时间范围快捷选项 ----------
// key: eval=评估起始日 / w1=近1周 / m1=近1月 / q1=近季度 / h1=近半年 /
//      y1=近1年 / y3=近3年 / y5=近5年 / y10=近10年 / all=全部
// 按钮文字按当前语言取 ui_r_<key>
const RANGES = [
  { key: "eval" },
  { key: "w1",   days: 7 },
  { key: "m1",   days: 30 },
  { key: "q1",   days: 91 },
  { key: "h1",   days: 182 },
  { key: "y1",   days: 365 },
  { key: "y3",   days: 1095 },
  { key: "y5",   days: 1825 },
  { key: "y10",  days: 3650 },
  { key: "all" }
];
// 各指标记住用户上次选择（localStorage 持久化），默认跟随评估起始日
function savedRangeKey(ind) {
  try {
    const k = localStorage.getItem("macroRange." + ind.col);
    if (k && RANGES.some(r => r.key === k)) return k;
  } catch (e) {}
  return "eval";
}
// 以序列最后一个交易日为锚，往前推 N 个自然日的起始索引
function rangeStartIdx(ind, days) {
  const n = ind.dates.length;
  if (!n || days == null) return 0;
  const targetMs = Date.parse(ind.dates[n - 1]) - days * 86400000;
  for (let i = 0; i < n; i++) {
    if (Date.parse(ind.dates[i]) >= targetMs) return i;
  }
  return 0;
}
function zoomStartForKey(ind, key) {
  if (key === "eval") return zoomStartPct(ind);
  const r = RANGES.find(x => x.key === key);
  if (!r) return zoomStartPct(ind);
  if (r.days == null) return 0;   // 全部
  return Math.min(100, 100 * rangeStartIdx(ind, r.days) / ind.dates.length);
}
function syncRangeButtons(ind, key) {
  const row = document.getElementById("r_" + ind.col);
  if (!row) return;
  row.querySelectorAll("button").forEach(b => {
    b.classList.toggle("on", b.dataset.key === key);
  });
}
// 切换某张图的时间范围（仅调整 dataZoom，无需重建图表）
function applyRange(ind, key) {
  const c = charts.find(x => x.__col === ind.col);
  if (c) {
    const zs = zoomStartForKey(ind, key);
    c.setOption({ dataZoom: [
      { type: "inside", start: zs, end: 100 },
      { type: "slider", start: zs, end: 100 }
    ]});
    c.__rangeMode = key;
  }
  try { localStorage.setItem("macroRange." + ind.col, key); } catch (e) {}
  syncRangeButtons(ind, key);
}
function fmtVal(v) {
  if (v == null || !isFinite(v)) return "-";
  if (Math.abs(v) >= 10000) return Math.round(v).toLocaleString("en-US");
  return v >= 1000 ? v.toFixed(1) : v >= 10 ? v.toFixed(2)
       : v >= 1 ? v.toFixed(3) : v.toFixed(4);
}
// 当前所选范围对应的起始索引（与图表 dataZoom 的窗口完全一致）
function windowStartIdx(ind, key) {
  if (key === "eval") return idxOnOrAfter(ind);
  const r = RANGES.find(x => x.key === key);
  if (!r || r.days == null) return 0;
  return rangeStartIdx(ind, r.days);
}
// 在所选范围内直接取实际最高点/最低点，并与范围内最新值比较
function swingsInWindow(ind, key) {
  const n = ind.dates.length;
  if (!n) return null;
  const i0 = windowStartIdx(ind, key);
  let hi = null, lo = null, cur = null;
  for (let i = i0; i < n; i++) {
    const v = ind.values[i];
    if (v == null || !isFinite(v)) continue;
    cur = v;                                   // 范围内最后一个有效值
    if (!hi || v > hi.v) hi = { d: ind.dates[i], v: v };
    if (!lo || v < lo.v) lo = { d: ind.dates[i], v: v };
  }
  if (!hi || !lo || cur == null) return null;
  const chg = (from, to) =>
    (from == null || from <= 0 || !isFinite(from)) ? null : (to - from) / from * 100;
  return { hi: hi, lo: lo, cur: cur,
           hi_chg: chg(hi.v, cur),      // 距最高点：≤0，越接近 0 越强势
           lo_chg: chg(lo.v, cur),      // 距最低点：≥0，越大说明反弹越多
           phase: (lo.d > hi.d) ? "up" : "down" };   // 哪个更晚出现，黄底标哪个
}
// H/L 标签：随所选时间范围实时重算；黄底＝当前所处阶段
function swingHtml(ind, key) {
  let s = swingsInWindow(ind, key);
  // 兜底：范围内有效点不足时退回服务端预计算的全序列摆动点
  if (!s && (ind.hi != null || ind.lo != null)) {
    s = { hi: { d: ind.hi_date, v: ind.hi }, lo: { d: ind.lo_date, v: ind.lo },
          hi_chg: ind.hi_chg, lo_chg: ind.lo_chg, phase: ind.phase };
  }
  if (!s) return "";
  const pctTxt = p => p == null ? "-" : (p > 0 ? "+" : "") + p.toFixed(2) + "%";
  const cls = p => p > 0 ? "up" : (p < 0 ? "down" : "flat");
  const rname = t("ui_r_" + key);
  let out = "";
  out += `<span class="sw${s.phase === "down" ? " now" : ""}">` +
         `<span class="tag">H</span>${fmtVal(s.hi.v)}` +
         `<span class="dt">${(s.hi.d || "").slice(2)}</span> ` +
         `<span class="${cls(s.hi_chg)}">${pctTxt(s.hi_chg)}</span></span>`;
  out += `<span class="sw${s.phase === "up" ? " now" : ""}">` +
         `<span class="tag">L</span>${fmtVal(s.lo.v)}` +
         `<span class="dt">${(s.lo.d || "").slice(2)}</span> ` +
         `<span class="${cls(s.lo_chg)}">${pctTxt(s.lo_chg)}</span></span>`;
  return `<div class="swings" title="${t("ui_swing_tip", {r: rname})}">${out}</div>`;
}
function updateSwing(ind, key) {
  const box = document.getElementById("sw_" + ind.col);
  if (box) box.innerHTML = swingHtml(ind, key);
}
function buildCard(ind) {
  const card = document.createElement("div");
  card.className = "card";
  const cls1 = ind.chg1 > 0 ? "up" : (ind.chg1 < 0 ? "down" : "flat");
  const clsy = ind.yoy > 0 ? "up" : (ind.yoy < 0 ? "down" : "flat");
  const rangeBtns = RANGES.map(r =>
    `<button data-key="${r.key}"` +
    (r.key === "eval" ? ' title="' + t("ui_eval_tip", {d: EVAL_START}) + '"' : "") +
    `>${t("ui_r_" + r.key)}</button>`).join("");
  card.innerHTML = `<h3><span title="${indDesc(ind.col)}">${indLabel(ind.col)}</span>
    <span class="chg">${t("ui_yoy")} <span class="${clsy}">${ind.yoy_txt}</span>
    ｜ ${ind.monthly ? t("ui_mom") : t("ui_chg1m")} <span class="${cls1}">${ind.chg1_txt}</span>
    ｜ <span id="sc_${ind.col}"></span></span></h3>
    <div class="kpi">${fmtVal(ind.latest)}</div>
    <div class="swings" id="sw_${ind.col}"></div>
    <div class="ranges" id="r_${ind.col}">${rangeBtns}</div>
    <div class="chart" id="c_${ind.col}"></div>`;
  card.querySelectorAll(".ranges button").forEach(b => {
    b.addEventListener("click", () => applyRange(ind, b.dataset.key));
  });
  return card;
}
// 刷新各卡片「自起始日 ±x.xx%」标注（数据起点晚于所选日期时显示实际基准月）
function updateSinceLabels() {
  DATA.forEach(ind => {
    const el = document.getElementById("sc_" + ind.col);
    if (!el) return;
    const r = sincePct(ind);
    if (r == null) { el.innerHTML = ""; return; }
    const p = r.pct;
    const cls = p > 0 ? "up" : (p < 0 ? "down" : "flat");
    const bm = r.baseDate.slice(0, 7);
    const tag = bm > EVAL_START.slice(0, 7)
      ? t("ui_since_star", {ym: bm})
      : t("ui_since", {ym: EVAL_START.slice(0, 7)});
    el.innerHTML = `${tag} <span class="${cls}">${p >= 0 ? "+" : ""}${p.toFixed(2)}%</span>`;
    if (bm > EVAL_START.slice(0, 7)) {
      el.title = t("ui_since_tip", {ym: bm});
    } else { el.title = ""; }
  });
}
function optionFor(ind, rangeKey) {
  const zoomStart = zoomStartForKey(ind, rangeKey || "eval");
  return {
    animation: false,
    tooltip: { trigger: "axis" },
    grid: { left: 50, right: 12, top: 8, bottom: 42 },
    xAxis: { type: "category", data: ind.dates,
             axisLabel: { fontSize: 10 } },
    yAxis: { type: "value", scale: true,
             axisLabel: { fontSize: 10 } },
    dataZoom: [
      { type: "inside", start: zoomStart, end: 100,
        zoomOnMouseWheel: true, moveOnMouseMove: true },
      { type: "slider", start: zoomStart, end: 100, height: 14, bottom: 4,
        showDetail: false, brushSelect: false, zoomOnMouseWheel: true,
        showDataShadow: false, borderColor: "#c8ccd4" }
    ],
    series: [{ type: "line", data: ind.values, showSymbol: false,
               lineStyle: { width: 1.4, color: ind.color },
               itemStyle: { color: ind.color } }]
  };
}
function renderAll() {
  charts.length = 0;
  DATA.forEach(ind => {
    const el = document.getElementById("c_" + ind.col);
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const c = echarts.init(el, null, {
      renderer: "canvas",
      width: rect.width > 10 ? rect.width : undefined,
      height: 190
    });
    const mode = savedRangeKey(ind);          // 恢复该指标上次选择的时间范围
    c.setOption(optionFor(ind, mode));
    charts.push(c);
    c.__col = ind.col;                 // 图表 → 指标 映射
    c.__rangeMode = mode;              // 当前生效的时间范围
    new ResizeObserver(() => c.resize()).observe(el);
  });
  DATA.forEach(ind => {
    const mode = savedRangeKey(ind);
    syncRangeButtons(ind, mode);
    updateSwing(ind, mode);
  });
  updateSinceLabels();
}
// 评估起始日变更：仅重置处于「起始日」模式的图表；用户单独选了
// 时间范围（如 1年/5年）的图表保持不变；涨跌幅标注全部重算
function applyEvalStart(dateStr) {
  EVAL_START = dateStr;
  try { localStorage.setItem("macroEvalStart", dateStr); } catch (e) {}
  charts.forEach(c => {
    if (c.__rangeMode && c.__rangeMode !== "eval") return;
    const ind = DATA.find(d => d.col === c.__col);
    if (!ind) return;
    const zs = zoomStartPct(ind);
    c.setOption({ dataZoom: [
      { type: "inside", start: zs, end: 100 },
      { type: "slider", start: zs, end: 100 }
    ]});
  });
  // 刷新「起始日」按钮的悬停提示
  DATA.forEach(ind => {
    const b = document.querySelector(
      `#r_${ind.col} button[data-key="eval"]`);
    if (b) b.title = t("ui_eval_tip", {d: EVAL_START});
  });
  updateSinceLabels();
}
(function initEvalCtrl() {
  const input = document.getElementById("evalStart");
  input.value = EVAL_START;
  input.addEventListener("change", () => {
    if (/^\d{4}-\d{2}-\d{2}$/.test(input.value)) applyEvalStart(input.value);
  });
  document.getElementById("evalReset").addEventListener("click", () => {
    const d = defaultEvalStart();
    input.value = d;
    applyEvalStart(d);
  });
})();
(function buildCards() {
  const root = document.getElementById("cards");
  const groups = {};
  DATA.forEach(ind => {
    (groups[ind.group] = groups[ind.group] || []).push(ind);
  });
  // 分组键为语言无关的 gkey（rates/fx/risk/cmd），标题按当前语言渲染
  __GKEY_ORDER__.forEach(g => {
    if (!groups[g] || !groups[g].length) return;
    const h = document.createElement("div");
    h.className = "grp";
    h.textContent = t("ui_g_" + g);
    root.appendChild(h);
    const grid = document.createElement("div");
    grid.className = "grid";
    groups[g].forEach(ind => grid.appendChild(buildCard(ind)));
    root.appendChild(grid);
  });
})();
function startRender() {
  requestAnimationFrame(() => setTimeout(renderAll, 60));
}
if (document.readyState === "complete") startRender();
else window.addEventListener("load", startRender);
window.addEventListener("resize", () => charts.forEach(c => c.resize()));

// ---------- 第三层：派生指标 ----------
(function () {
  const root = document.getElementById("derived");
  DERIVED.forEach(d => {
    const label = d.label_key ? t(d.label_key) : d.label;
    const note = d.note_key ? t(d.note_key) : d.note;
    const el = document.createElement("div");
    el.className = "dbox";
    el.innerHTML = `<div class="k">${label}</div>
                    <div class="v">${d.value_txt}</div>
                    <div class="n">${note}</div>`;
    root.appendChild(el);
  });
})();

// ---------- 相关性矩阵（日收益率 Pearson，窗口可切换 30/90/365） ----------
(function () {
  const keys = ["US10Y", "US2Y", "Gold", "Copper", "DXY", "VIX", "SPX", "NDX", "Brent", "HYG"];
  const nameOf = k => indLabel(k);
  // 日收益率序列（全量）
  const rets = {};
  DATA.forEach(ind => {
    if (keys.indexOf(ind.col) < 0) return;
    const rs = [];
    let prev = null;
    ind.values.forEach(v => {
      if (v != null && prev != null) rs.push((v - prev) / prev);
      if (v != null) prev = v;
    });
    rets[ind.col] = rs;
  });
  function corr(a, b) {
    const n = Math.min(a.length, b.length);
    if (n < 15) return null;
    a = a.slice(-n); b = b.slice(-n);
    let ma = 0, mb = 0;
    for (let i = 0; i < n; i++) { ma += a[i]; mb += b[i]; }
    ma /= n; mb /= n;
    let sab = 0, va = 0, vb = 0;
    for (let i = 0; i < n; i++) {
      const da = a[i] - ma, db = b[i] - mb;
      sab += da * db; va += da * da; vb += db * db;
    }
    const den = Math.sqrt(va * vb);
    return den === 0 ? null : sab / den;
  }
  function bg(r) {
    if (r == null) return "#fff";
    const t = Math.min(1, Math.abs(r));
    const a = 0.12 + 0.4 * t;
    return r > 0 ? `rgba(192,0,0,${a})` : `rgba(0,128,0,${a})`;
  }
  function render(N) {
    let html = '<table class="ctab"><tr><th></th>' +
               keys.map(k => `<th>${nameOf(k)}</th>`).join("") + "</tr>";
    keys.forEach(rk => {
      html += `<tr><td class="lab">${nameOf(rk)}</td>`;
      keys.forEach(ck => {
        if (rk === ck) {
          html += `<td style="background:#f0f0f0">1.00</td>`;
        } else {
          const r = corr((rets[rk] || []).slice(-N), (rets[ck] || []).slice(-N));
          html += `<td style="background:${bg(r)}">${r == null ? "-" : r.toFixed(2)}</td>`;
        }
      });
      html += "</tr>";
    });
    html += "</table>";
    document.getElementById("corrtab").innerHTML = html;
  }
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(b => b.classList.remove("on"));
      btn.classList.add("on");
      render(parseInt(btn.dataset.n, 10));
    });
  });
  render(90);
})();

// ---------- Regime 历史 ----------
(function () {
  if (!RHIST || !RHIST.length) return;
  const cls = a => a === "↑" ? "g-up" : (a === "↓" ? "g-dn" : "g-fl");
  let html = '<table class="htab"><tr><th>' + t("ui_th_month") + '</th>' +
             RHIST.map(h => `<th>${h.month}</th>`).join("") +
             '</tr><tr><td class="lab">' + t("ui_th_growth") + '</td>' +
             RHIST.map(h => `<td class="${cls(h.g)}">${h.g}</td>`).join("") +
             '</tr><tr><td class="lab">' + t("ui_th_infl") + '</td>' +
             RHIST.map(h => `<td class="${cls(h.i)}">${h.i}</td>`).join("") +
             '</tr></table>';
  document.getElementById("histtab").innerHTML = html;
})();

// ---------- 资产倾向 ----------
(function () {
  if (!BIAS || !BIAS.length) return;
  let html = '<table class="btab"><tr><th>' + t("ui_th_asset") + '</th><th>' +
             t("ui_th_bias") + '</th><th>' + t("ui_th_note") + '</th></tr>';
  BIAS.forEach(b => {
    const asset = b.asset_key ? t(b.asset_key) : b.asset;
    const bias = b.bias_key ? t(b.bias_key) : b.bias;
    const note = b.note_key ? t(b.note_key) : (b.note || "");
    html += `<tr><td style="font-weight:600">${asset}</td><td style="text-align:center">${bias}</td><td style="color:#666">${note || "-"}</td></tr>`;
  });
  html += "</table>";
  document.getElementById("biastab").innerHTML = html;
})();

// ---------- 第四层：自动解读 ----------
(function () {
  const root = document.getElementById("insights");
  const renderInsight = i => {
    const title = i.t ? t(i.t) : i[0];
    const text = i.segs
      ? i.segs.map(s => t(s[0], s[1] || undefined)).join("")
      : i[1];
    return {title: title, text: text};
  };
  REGIME.insights.forEach(i => {
    const r = renderInsight(i);
    const p = document.createElement("div");
    p.className = "item";
    p.innerHTML = `<b>【${r.title}】</b>${r.text}`;
    root.appendChild(p);
  });
  const sum = document.createElement("div");
  sum.className = "item";
  const summary = REGIME.sum_bits
    ? t("sum_prefix") + REGIME.sum_bits.map(b => t(b)).join(" + ") + t("sum_suffix")
    : REGIME.summary;
  sum.innerHTML = `<b>【${t("ui_sum_title")}】</b>${summary}`;
  root.appendChild(sum);
})();
</script>
</body>
</html>
"""


def build_html(dates: list, merged: dict, status: dict = None) -> None:
    latest_d = dates[-1]
    m1, m3 = month_ref(latest_d, 30), month_ref(latest_d, 90)
    m1y = month_ref(latest_d, 365)
    # 近 DAILY_KEEP 个交易日保留日线，更早的按每 5 个交易日抽样（控制 HTML 体积）
    daily_keep = 800
    n = len(dates)
    if n > daily_keep:
        keep_idx = set(range(n - daily_keep, n))       # 最近 800 个交易日：全保留
        keep_idx.update(range(0, n - daily_keep, 5))   # 更早历史：每 5 个交易日取 1 点
        keep_idx.add(n - 1)
        idxs = sorted(keep_idx)
    else:
        idxs = list(range(n))
    slim_dates = [dates[i] for i in idxs]
    payload = []
    for col, sym, label, _, _, color, group in INDICATORS:
        cur = value_on_or_before(dates, merged, col, latest_d)
        if cur is None:
            continue
        # 前期摆动高低点（与看板同一套算法）
        _pairs = series_pairs(dates, merged, col)
        if col in MONTHLY_COLS:
            # 月度：按数据点回看（上月 / 3个月前 / 12个月前）
            _vals = [p[1] for p in _pairs]
            v1 = _vals[-2] if len(_vals) >= 2 else None
            v3 = _vals[-4] if len(_vals) >= 4 else None
            vy = _vals[-13] if len(_vals) >= 13 else None
        else:
            v1 = value_on_or_before(dates, merged, col, m1)
            v3 = value_on_or_before(dates, merged, col, m3)
            vy = value_on_or_before(dates, merged, col, m1y)
        p1, p3 = chg_pct(cur, v1), chg_pct(cur, v3)
        py = chg_pct(cur, vy)
        hi, lo = recent_swings(_pairs, SWING_PCT.get(col, SWING_PCT_DEFAULT))
        phase = swing_phase(hi, lo)
        payload.append({
            "col": col, "label": label, "color": color, "latest": cur,
            "group": GKEY.get(group, group),
            "chg1": p1, "chg3": p3, "yoy": py,
            "chg1_txt": f"{p1:+.2f}%" if p1 is not None else "-",
            "chg3_txt": f"{p3:+.2f}%" if p3 is not None else "-",
            "yoy_txt": f"{py:+.2f}%" if py is not None else "-",
            "monthly": col in MONTHLY_COLS,
            # 前期摆动高低点
            "hi": hi[1] if hi else None,
            "hi_date": hi[0] if hi else None,
            "lo": lo[1] if lo else None,
            "lo_date": lo[0] if lo else None,
            "hi_chg": chg_pct(cur, hi[1]) if hi else None,
            "lo_chg": chg_pct(cur, lo[1]) if lo else None,
            "phase": phase,
            # ZigZag 阈值传给前端：高低点随所选时间范围在浏览器端实时重算
            "swing_pct": SWING_PCT.get(col, SWING_PCT_DEFAULT),
            "dates": slim_dates,
            "values": [merged.get(d, {}).get(col) for d in slim_dates],
        })
    regime = build_regime(dates, merged)
    derived = compute_derived(dates, merged)
    factors = build_factors(dates, merged)
    zscores = build_zscores(dates, merged)
    divergences = build_divergences(dates, merged)
    bias = build_asset_bias(factors)
    rhist = build_regime_history(dates, merged)
    if status is None:
        status = {"level": "cache", "text": ""}
    # 评估起始日默认「最近 12 个月」：随每次生成自动顺延
    _today = dt.date.today()
    _y, _m = _today.year, _today.month - 12
    while _m <= 0:
        _m += 12
        _y -= 1
    _last = (dt.date(_y + (_m // 12), (_m % 12) + 1, 1) - dt.timedelta(days=1)).day
    _eval12 = dt.date(_y, _m, min(_today.day, _last)).isoformat()
    html = (HTML_TMPL.replace("__CDN__", ECHARTS_CDN)
            .replace("__TIME__", f"{dt.datetime.now():%Y-%m-%d %H:%M}")
            .replace("__EVAL12__", _eval12)
            .replace("__EVALMAX__", _today.isoformat())
            .replace("__DATA__", json.dumps(payload, ensure_ascii=False))
            .replace("__REGIME__", json.dumps(regime, ensure_ascii=False))
            .replace("__DERIVED__", json.dumps(list(derived.values()),
                                                ensure_ascii=False))
            .replace("__STATUS__", json.dumps(status, ensure_ascii=False))
            .replace("__FACTORS__", json.dumps(factors, ensure_ascii=False))
            .replace("__ZSCORES__", json.dumps(zscores, ensure_ascii=False))
            .replace("__DIVERG__", json.dumps(divergences, ensure_ascii=False))
            .replace("__BIAS__", json.dumps(bias, ensure_ascii=False))
            .replace("__HISTORY__", json.dumps(rhist, ensure_ascii=False))
            .replace("__L10N__", json.dumps(TR, ensure_ascii=False))
            .replace("__IND_L10N__", json.dumps(IND_L10N, ensure_ascii=False))
            .replace("__LANGS__", json.dumps(list(LANGS), ensure_ascii=False))
            .replace("__LANG_NAMES__", json.dumps(LANG_NAMES, ensure_ascii=False))
            .replace("__DEFLANG__", DEFAULT_LANG)
            .replace("__GKEY_ORDER__",
                     json.dumps(GKEY_ORDER, ensure_ascii=False)))
    HTML.write_text(html, encoding="utf-8")
    log(f"  HTML 决策仪表盘已生成：{HTML.name}"
        f"（Regime: {factors['quadrant']}，分歧 {len(divergences)} 项）")


# ----------------------------------------------------------------------------
# Cloudflare Pages 自动发布
# ----------------------------------------------------------------------------
def publish_to_cloudflare() -> None:
    """将最新仪表盘发布到 Cloudflare Pages（任何失败均不影响本地更新）。

    依赖 economy/cf_config.json：{"token": "...", "account_id": "...",
    "project": "macro-dashboard"}；文件不存在则直接跳过。
    网络抖动时自动重试（最多 2 次，间隔 8 秒），保证日常定时发布可靠。
    """
    # CI（GitHub Actions）环境由工作流单独用 wrangler 部署，这里跳过，
    # 避免找不到本机 WorkBuddy 托管的 wrangler 路径而误报。
    if os.environ.get("GITHUB_ACTIONS"):
        return
    if not CF_CONFIG.exists():
        return
    try:
        cfg = json.loads(CF_CONFIG.read_text(encoding="utf-8"))
        token = cfg.get("token")
        account = cfg.get("account_id")
        project = cfg.get("project", "macro-dashboard")
        if not token or not account:
            log("  [发布] cf_config.json 缺 token/account_id，跳过云端发布")
            return
        if not WRANGLER_NODE.exists() or not WRANGLER_JS.exists():
            log("  [发布] 找不到 wrangler，跳过云端发布")
            return

        def _attempt(attempt_no):
            with tempfile.TemporaryDirectory() as td:
                deploy_dir = Path(td) / "site"
                deploy_dir.mkdir()
                # Pages 要求入口文件名为 index.html
                shutil.copy2(HTML, deploy_dir / "index.html")
                # 计划任务上下文里 HOME/APPDATA 可能缺失，wrangler 会静默失败
                env = {**os.environ,
                       "CLOUDFLARE_API_TOKEN": token,
                       "CLOUDFLARE_ACCOUNT_ID": account,
                       "HOME": str(Path.home()),
                       "USERPROFILE": str(Path.home()),
                       "APPDATA": os.environ.get("APPDATA")
                                  or str(Path.home() / "AppData" / "Roaming"),
                       "LOCALAPPDATA": os.environ.get("LOCALAPPDATA")
                                       or str(Path.home() / "AppData" / "Local"),
                       "CI": "1",
                       "WRANGLER_SEND_METRICS": "false"}
                r = subprocess.run(
                    [str(WRANGLER_NODE), str(WRANGLER_JS), "pages", "deploy",
                     str(deploy_dir), "--project-name", project,
                     "--branch", "main"],
                    capture_output=True, text=True, env=env, timeout=180,
                    encoding="utf-8", errors="replace")
            out = (r.stdout or "") + (r.stderr or "")
            ok = (r.returncode == 0
                  and ("Success" in out or "Deployment complete" in out))
            if not ok:
                out = f"[returncode={r.returncode}] {out}"
            return ok, out

        for attempt_no in (1, 2):
            ok, out = _attempt(attempt_no)
            if ok:
                log(f"  [发布] Cloudflare Pages 发布成功（项目 {project}）")
                return
            if attempt_no == 1:
                log(f"  [发布] 第 1 次发布失败，8 秒后重试：" + out.strip()[-120:])
                time.sleep(8)
        log("  [发布] 发布失败：" + out.strip()[-300:])
    except Exception as e:
        log(f"  [发布] 发布异常（不影响本地更新）：{e}")


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
def save_with_backup(wb) -> None:
    BACKUP_DIR.mkdir(exist_ok=True)
    if XLSX.exists():
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(XLSX, BACKUP_DIR / f"宏观指标跟踪_{stamp}.xlsx")
        backups = sorted(BACKUP_DIR.glob("*.xlsx"))
        for old in backups[:-BACKUP_KEEP]:
            old.unlink(missing_ok=True)
    for attempt in range(3):
        try:
            wb.save(XLSX)
            return
        except PermissionError:
            log(f"  文件被占用（OneDrive/Excel 可能正打开），"
                f"第 {attempt + 1} 次重试...")
            time.sleep(3)
    raise PermissionError("Excel 文件被锁定，请关闭 Excel 后重试")


def main() -> int:
    log("=" * 60)
    log("开始更新宏观指标 (v3 宏观决策支持系统)")
    try:
        if not XLSX.exists():
            log(f"错误：找不到文件 {XLSX}")
            return 1
        fetched, errors = fetch_all()
        backfill_from_fred(fetched, errors)   # FRED 补长历史（US2Y/Brent/WTI）
        quality_layer(fetched)            # P0 数据质量层：异常点检测与剔除
        status = {"level": "live", "text": ""}
        if errors:
            status = {"level": "partial",
                      "text": "、".join(errors.keys()) + " 使用缓存数据"}
        wb = openpyxl.load_workbook(XLSX)
        ws_hist = wb[HIST_SHEET]
        added, updated, dates, merged = merge_history(ws_hist, fetched)
        log(f"  历史数据合并完成：新增 {added} 天，更新 {updated} 个值，"
            f"共 {len(dates)} 天")
        refresh_snapshot(wb[SNAP_SHEET], dates, merged)
        log("  最新看板已刷新（含状态评级、派生指标、宏观状态）")
        rebuild_chart_sheet(wb, dates, merged)
        save_with_backup(wb)
        build_html(dates, merged, status=status)
        publish_to_cloudflare()     # 发布最新仪表盘到 Cloudflare Pages
        log("✔ 更新成功完成")
        return 0
    except Exception:
        log("✘ 更新失败：\n" + traceback.format_exc())
        return 2


if __name__ == "__main__":
    sys.exit(main())
