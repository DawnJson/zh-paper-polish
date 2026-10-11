# -*- coding: utf-8 -*-
"""中文论文文风机械校验：按人写语料基线判密度，而不是把标点一律清零。

默认基线取自 2015 年后中文核心期刊全文 165 篇，存于 corpus/baseline.json，分位数表见
references/human-baseline.md；--baseline 换成按目标刊语料重建的基线
（corpus/build_baseline.py），下文门禁数值即随之改变。
门禁取人写 p99：超过它说明该稿比 99% 的已发表论文更极端，判违规才站得住。
B 级取 p10 与 p95 双边：高于 p95 是机械堆砌，低于 p10 是过度清洗后的塑料文风，
两头都不像人写，但都只提示不阻断。

输入：Markdown 与纯文本按物理行读；.docx 按文档顺序读正文段落与表格，报告里的行号即
docx 正文层段落序号（含空段）。表格、图题表题与表注不进密度统计，只做 V/T 级检查。

判定口径全部字面可复算，不做语义推断（括号里的数值是默认基线 baseline.json 的取值）：
  F1 一字线   U+2014 及同族（– − － ― ー）；逐处定性为
              破折号 / 标准编号 / 数值区间 / 数学负号 / 概念串联 / 其他一字线族，
              连排 ≥2 个一字线记破折号；单个一字线两侧任一为汉字记概念串联。
              默认豁免 标准编号 + 数值区间 + 数学负号；其余三类各按密度判
              （--dash-k 0.56、--dash-concept-k 3.45、--dash-other-k 0.92）
  F2 斜杠     含全角 ／ 与 ∕；默认豁免 标准代号 / 物理单位 / 分数比例 / 英文串联，
              并列替代 ≤ --slash-k 2.14；其余未豁免 kind 计 0 即判
  F3 双引号   直角 “…” 与弯引号 “…” ≤ --quote-k 3.04 每千字
  F4 连续并列 同段枚举族 ≥--fam-para-min 成员（默认 3）记 1 组，豁免族按组扣除，
              同段 ≤ --fam-para-max 7 组，跨段 ≤ --fam-span 7 组
  F5 骨架重复 相邻句首 3 字相同（仅统计）
  F6 英文摘要 Abstract / Keywords 行只判混入的中文标点与中文字符，严禁
  F7 冒号     逐处定性为 清单冒号 / 正文冒号 / 比例时间 / 公式引出；后两类豁免，
              正文冒号 ≤ --colon-k 4.19，清单冒号 ≤ --colon-list-k 1.65
  F8 书名号   成对书名号 ≤ --book-k 0.14 每千字；法规文件名、标准正式名称与孤立
              书名号不计入
  F9 数字分节 --num-group 默认 plain：四位以上数值一律写 1200，写成 1 200 即判违规；
              grouped 反向，either 只判混用。代号、序号与年份分节一律严禁
  F10 量值间隔 --unit-space 默认 space：数值与单位一律留空，写 80 m/min；
              --qty-lead 默认 tight：汉字与数值一律贴合，写 下游6～8 m处、较2000年增长约122%
              （不带单位的年份、百分数、计数只查留空一侧）；百分号前留空一律严禁。
              两项都可切 tight/space/either，either 只判同一稿件内混用
  F11 浪纹线   数值范围的连接号，GB/T 15834 4.13.2 规定为全角 ～；半角 ~ 一律严禁
  A1 隐形字符 零宽、双向控制符、软连字符、Unicode 私有区，一律严禁
              （人写语料中仅见于 PDF 取字残留）
  A7 泛权威   研究表明 / 众所周知 一类断言左 25 字与右 40 字都无引用标记，
              ≤ --auth-k 0.33；[待补引]、[5,19,22] 与 (2019) 都算引用标记
  A11 段首实体 段落（≥40 字）首句以专业实体或目的状语起句的比例 ≥ --opener-min
              （人写 p10 80.0%，中位 91.7%），低于阈值告警
  A12 并列悬空 顿号并列提出若干项，段内后文一项也没再回应（其后至少两句才判）
  A13 压缩简称 时延、归并、粒度、全量、精标注、根因一类词表命中，附还原提示
  A14 术语首现 词表内术语全文首次出现时，同句无"即/是指"或中文括注；
              --known-terms 剔除读者已熟悉的术语，--terms-file 追加本文术语
  A15 抽象施事 抽象名词作主语接施事动作（先验转向、需求倒逼、本质导向）
  A2/A4/A5/A6/A8/A9/A10/A12～A15  段首零回指、句尾空转挂尾、虚动词名词化、弱谓词密度、
              顿号链长、连词密度、相邻句字二元组重合度、并列悬空与可读性三项，一律只告警不阻断
  L1 分号     每段 ≤ --semi-para 18 处，全文 ≤ --semi-k 4.45 每千字
  B* 分布偏离 14 项密度加长句占比、句长变异、段均字数、段首立论率，落在 p10~p95 之外即列出；
              ≥4 项时提示回填（人写仅 12.1% 达到）
  I* 参考量   冒号总数、长句、跨段重复片段，仅统计不判违规
  V  用词     scripts/lexicon.tsv 逐词计数：同段同类 ≥2 处告警；单词 ≥2 处且超该词人写 p99
              告警；类别 ≥2 处且超类别人写 p99 进门禁。表格、图题、表注逐处列出，不计密度
  T1 缩写首现 全大写缩写首次出现未给全称、定义晚于首现、重复定义或英文全称不一致（告警）
  T2 术语变体 scripts/terms.tsv 与 --terms 里同组写法在同一稿件并用（告警）
  C1 的字链   单个小句 ≥3 个结构助词「的」（告警，附人写 p95）
  C2 递进堆叠 不仅……而且……更…… 一类一句三层递进（告警）
  C3 等长排比 同一句内逗号分隔的 ≥3 个等长短句（3～8 字）连排（告警）
  C4 节号回指 正文用「2.2.4节」「第3节」指代内容，全文 ≤ --secref-k 0.98 每千字（门禁）；
              段首句带节号、一句 ≥2 处回指逐条列出，路线图句（本文结构安排如下…、第2节介绍…）不列
  C5 压缩句   ≥50 字且不含结构助词「的」的句子占 ≥50 字句的比例 ≤ --nde-long 16.67%（门禁）；
              逐句列出，含 ≥3 个数的参数报告句标 [数值]；「的」密度低于人写 p05 另提示
  C6 图表括注 括号内容以图表号开头、含「见/如＋图表号」或以图表号收尾，一律严禁：
              （图3）（见表1）（如图2(a)所示）（R=0.019，见图4(c)）；图表引述写进句子
              （表1列出… / …如图3所示）。括号里顺带提到图表的说明语（与图1(c)中标记的面）不计。
              人写 267 篇 77 篇出现，属书写口径，不按基线放行
  P  待补引   [待补引] 占位逐处列出；A7 视其为已落引用，投稿前须补齐

豁免按“逐处定性的 kind”给，不是按符号给：`--allow-dash=标准编号` 只放过
GB/T 4897—2015 这类标准编号里的一字线，破折号照判。

用法：
  python style_lint.py 正文.md
  python style_lint.py --report=_stylelint.txt 稿件.docx
  python style_lint.py --allow-fam='数字分项 N）' --semi-para=4 --semi-k=3.0 正文.md
  python style_lint.py --baseline=目标刊.json --terms=稿件术语.tsv 正文.md

严禁项非 0、或限量项超门禁即以退出码 1 结束，可直接作为装配前置门禁；
源文件、基线或词表缺失退 2。
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_BASELINE = os.path.join(HERE, 'corpus', 'baseline.json')
DEFAULT_LEXICON = os.path.join(HERE, 'lexicon.tsv')
DEFAULT_TERMS = os.path.join(HERE, 'terms.tsv')

# ---------------- 需排除的非正文行 ----------------
SKIP_PAT = [
    re.compile(r'^!\['),                              # 插图
    re.compile(r'^\{\{[A-Z]+\d*\}\}$'),               # 表格等占位符
    re.compile(r'^\*{0,2}(图|表)\s?\d+\*{0,2}(\s|$|[：:])'),  # 中文图题/表题（含加粗）
    re.compile(r'^\*{0,2}(Fig|Tab)(\.|le|ure)', re.I),  # 英文图题/表题（含加粗）
    re.compile(r'^\*{0,2}(关键词|中图分类号|文献标识码|国家标准学科分类代码|'
               r'收稿日期|基金项目|作者简介|摘\s*要|通信作者|引用格式|DOI)'),
    re.compile(r'^#{1,6}\s'),                         # Markdown 标题
    re.compile(r'^\d+(\.\d+)*\s'),                    # 编号章节标题（带空格）
    re.compile(r'^\d+(\.\d+)*\s*[\u4e00-\u9fffA-Za-z][^。；！？]{1,28}$'),
    re.compile(r'^[A-Za-z][A-Za-z0-9 ,.:;()\-]+$'),   # 纯英文行（英文标题）
]

# 表格行与分隔线：不进门禁与排除行提醒，只作为辅助文本进 V/T 级检查
TABLE_PAT = re.compile(r'^([+|]|\s*[-=:+|]{4,}\s*$)')
GRID_GAP = re.compile(r'\S {3,}\S')


def is_table_line(s):
    """管道表、网格表分隔线，以及 pandoc 简单表里靠多空格分栏的数据行。"""
    return bool(TABLE_PAT.match(s)) or len(GRID_GAP.findall(s)) >= 2

# ---------------- 枚举族 ----------------
FAMILIES = [
    ('其一/其二/其三', re.compile(r'其[一二三四五六七八九]')),
    ('首先/其次/最后', re.compile(r'首先|其次|再次|最后|然后')),
    ('一方面/另一方面', re.compile(r'另?一方面')),
    ('一是/二是/三是', re.compile(r'[一二三四五六七八九]是[，,、]')),
    ('第N阶段/层/类/条/方面', re.compile(
        r'第[一二三四五六七八九十](阶段|层|类|条|方面|部分|点|步|环节|主线|条主线)')),
    ('数字分项N）', re.compile(r'[1-9]）')),
    ('字母分项（a）', re.compile(r'（[a-dA-D]）')),
    ('圈号①②③', re.compile(r'[①②③④⑤⑥⑦⑧⑨]')),
]
FAMILY_NAMES = [name for name, _ in FAMILIES]

# ---------------- 人写语料基线（corpus/baseline.json，由 corpus/build_baseline.py 生成） ----------------
# B 级双边比对的 14 项密度（每千字）：低于 p10 记过净，高于 p95 记过密
BAND_METRICS = ['colon_body_k', 'colon_list_k', 'quote_k', 'dash_break_k', 'dash_concept_k',
                'dash_other_k', 'slash_pair_k', 'book_k', 'authority_k', 'semi_k', 'dun_k',
                'conj_k', 'weak_k', 'nominal_k']
BAND_LABELS = {
    'colon_body_k': '正文冒号', 'colon_list_k': '清单冒号', 'quote_k': '双引号',
    'dash_break_k': '破折号', 'dash_concept_k': '概念串联', 'dash_other_k': '其他一字线族',
    'slash_pair_k': '并列斜杠', 'book_k': '书名号', 'authority_k': '无引用泛权威',
    'semi_k': '分号', 'dun_k': '顿号', 'conj_k': '逻辑连词', 'weak_k': '弱谓词',
    'nominal_k': '名词化'}
# 段落形态下限（低于人写 p10 即为过净）
FLOOR_METRICS = ['long80', 'sent_cv', 'para_chars']
# 门禁默认取基线 p99：Config 属性名 → 基线指标键（即 report() 汇总字典的键）
GATE_P99 = [('colon_k', 'colon_body_k'), ('colon_list_k', 'colon_list_k'),
            ('quote_k', 'quote_k'), ('dash_k', 'dash_break_k'),
            ('dash_concept_k', 'dash_concept_k'), ('dash_other_k', 'dash_other_k'),
            ('slash_k', 'slash_pair_k'), ('book_k', 'book_k'), ('auth_k', 'authority_k'),
            ('semi_k', 'semi_k'), ('semi_para', 'semi_max'), ('fam_para_max', 'fam'),
            ('fam_span', 'fam_span')]
GATE_INT = {'semi_para', 'fam_para_max', 'fam_span'}
# 同取 p99，但旧基线可能没有这两项（C4 节号回指、C5 无「的」长句占比），缺则不判
GATE_OPT = [('secref_k', 'secref_k'), ('nde_long', 'nde_long')]


def load_baseline(path):
    with io.open(path, encoding='utf-8') as fh:
        base = json.load(fh)
    missing = [k for k in BAND_METRICS + FLOOR_METRICS + [key for _, key in GATE_P99]
               + ['opener'] if k not in base.get('metrics', {})]
    if missing:
        raise ValueError('%s 缺少指标：%s' % (path, '、'.join(missing)))
    return base


# ---------------- 段首类型（A11） ----------------
OPENER_PATS = [
    ('背景垫句', re.compile(r'^(随着|近年来|当前，|当今|在.{0,12}背景下|在.{0,10}时代)')),
    ('提示词起句', re.compile(r'^(值得注意|值得一提|众所周知|不难看出|显而易见|需要指出|需要说明)')),
    ('枚举词起句', re.compile(r'^(首先|其次|再次|最后|此外|另外|另一方面|一方面|其一|其二|'
                          r'第[一二三四五六七八九十][、，]?)')),
    ('连接词起句', re.compile(r'^(因此|然而|但是|但|同时|因而|所以|由于|尽管|虽然|总之|综上)')),
    ('目的状语起句', re.compile(r'^(为了?[^，。]{2,20}[，]|针对[^，。]{2,20}[，]|面向[^，。]{2,20}[，])')),
    ('本文/本节起句', re.compile(r'^(本文|本节|本研究|本小节|我们|本课题组)')),
]
# 计入 A11 分子的段首类型（人写 ≥40 字段落占比 p10 80.0%、中位 91.7%，
# 见 references/human-baseline.md）
OPENER_OK = {'实体主语起句', '目的状语起句'}


def opener_kind(text):
    for name, pat in OPENER_PATS:
        if pat.match(text):
            return name
    return '实体主语起句'

CITE = re.compile(r'\{\{[^}]*\}\}')
PENDING = re.compile(r'\[待补引\]')
# docx 里的引用是正文字符 [12]、[6-7]、[5,19,22]，统计口径与 md 的 {{key}} 一样要剥掉
DOCX_CITE = re.compile(r'\[\d+(?:\s*[-\u2013\u2014,\uff0c]\s*\d+)*\]')
IMG = re.compile(r'!\[[^\]]*\]\([^)]*\)(\{[^}]*\})?')
SENT_END = re.compile(r'[。！？]')
EN_PAT = re.compile(r'^\*{0,2}(Abstract|Key\s*words)\b', re.I)
CAPTION_ZH = SKIP_PAT[2]
BANNED_CHARS = re.compile(r'[\u2014\u2013\u2212\uff0d/\uff0f\u2215"\u201c]')


REFS_HEAD = re.compile(r'^#{0,6}\s*\*{0,2}(参\s*考\s*文\s*献|References)\*{0,2}\s*[:：]?\s*$', re.I)
REF_ITEM = re.compile(r'^\\?\[\d+\\?\]')
NOTE_LINE = re.compile(r'^\*{0,2}(注|说明|Note)\s*[:：]')


def _classify(items, docx=False):
    """items 为 (行号, 文本, 单元格定位或 None)，返回 (正文段落, 英文摘要行, 排除行提醒, 辅助文本)。

    正文段落元素为 (行号, 统计文本, 原文)：统计文本已剥引用标记、[待补引] 与插图标记；
    原文只剥插图标记，留给 A7 查引标、P 级列待补引。
    辅助文本 = 表格、图题表题与表注，元素为 (定位, 统计文本)，只进 V/T 级检查。
    排除行提醒 = 被剔除但仍含一字线/斜杠/双引号的行，供人工判断。
    """
    out, en, notice, aux = [], [], [], []
    in_refs = False
    for lineno, raw, cell in items:
        s = raw.strip()
        if not s:
            continue
        body = IMG.sub('', PENDING.sub('', CITE.sub('', DOCX_CITE.sub('', s) if docx else s))).strip()
        if cell:
            aux.append((cell, body))
            continue
        if REFS_HEAD.match(s):
            in_refs = True
            continue
        if in_refs or REF_ITEM.match(s):
            continue
        if (not docx and is_table_line(s)) or NOTE_LINE.match(s):
            aux.append(('行%d' % lineno, body))
            continue
        if EN_PAT.match(s):
            en.append((lineno, body))
            continue
        if CAPTION_ZH.match(s):
            aux.append(('行%d' % lineno, body))
        if any(p.match(s) for p in SKIP_PAT) or len(body) < 20:
            if BANNED_CHARS.search(body):
                notice.append((lineno, body))
            continue
        out.append((lineno, body, IMG.sub('', s).strip()))
    return out, en, notice, aux


def load_paragraphs(path):
    """按扩展名读 md/纯文本或 docx，返回值见 _classify。"""
    if path.lower().endswith('.docx'):
        if HERE not in sys.path:
            sys.path.insert(0, HERE)
        from docx_text import read_docx
        items = [(b['n'], b['text'].replace('\n', ' '),
                  None if b['kind'] == 'para' else b['loc'])
                 for b in read_docx(path)]
        return _classify(items, docx=True)
    with io.open(path, encoding='utf-8') as fh:
        return _classify([(i, raw, None) for i, raw in enumerate(fh, 1)])


def sentences(text):
    parts, buf = [], []
    for ch in text:
        buf.append(ch)
        if SENT_END.match(ch):
            parts.append(''.join(buf).strip())
            buf = []
    if ''.join(buf).strip():
        parts.append(''.join(buf).strip())
    return [p for p in parts if p]


def quote_spans(text):
    """成对提取双引号；返回 [(引号内文字, 是否弯引号)]。"""
    found = []
    for m in re.finditer(r'"([^"]{1,60})"', text):
        found.append((m.group(1), False))
    for m in re.finditer(r'[\u201c]([^\u201d]{1,60})[\u201d]', text):
        found.append((m.group(1), True))
    return found


STD_CODE = re.compile(
    r'(?:GB|JB|JG|HG|SJ|YB|QB|DL|NY|LY|SN|CJ|TB|MT|AQ|ISO|IEC|ASTM|EN|DIN|'
    r'JIS|ANSI|IEEE|BS|GOST|Q)\s*/?\s*[TZY]?\s*[\d.\u2014\u2013-]*\s*$', re.I)
CJK_CHAR = re.compile(r'[\u4e00-\u9fff]')
NEG_LEFT = re.compile(r'(?:^|[=＝（(\[【，,：:；;\s\u4e00-\u9fff])$')


def dash_hits(text):
    """一字线族逐处定性；破折号与概念串联违规，负号/数值区间/标准编号可豁免。"""
    hits = []
    for m in re.finditer(r'[\u2014\u2013\u2212\uff0d\u2015\u30fc]+', text):
        run, s, e = m.group(0), m.start(), m.end()
        left, right = text[max(0, s - 12):s], text[e:e + 12]
        num_left = re.search(r'\d(?:\.\d+)?\s*$', left)
        num_right = re.match(r'\s*\d', right)
        if len(run) >= 2 and run[0] in '\u2014\u2015\uff0d\u2013':
            kind = '破折号'
        elif num_left and num_right and (
                STD_CODE.search(left) or re.match(r'\s?\d{4}', right)):
            kind = '标准编号'
        elif num_left and num_right:
            kind = '数值区间'
        elif num_right and not num_left and NEG_LEFT.search(left):
            kind = '数学负号'
        elif CJK_CHAR.search(left[-1:]) or CJK_CHAR.match(right[:1]):
            kind = '概念串联'
        else:
            kind = '其他一字线族'
        hits.append((kind, run, '\u2026%s%s%s\u2026' % (left, run, right)))
    return hits


COMMON_UNITS_LEFT = {
    'm', 'km', 'cm', 'mm', 'um', 'nm', 'kg', 'g', 'mg', 'ug', 't',
    'l', 'ml', 'r', 'rad', 's', 'min', 'h', 'd', 'pa', 'kpa', 'mpa',
    'n', 'kn', 'j', 'kj', 'mj', 'w', 'kw', 'mw', 'v', 'mv', 'kv',
    'a', 'ma', 'ka', 'hz', 'khz', 'mhz', 'ghz', 'px', 'bit', 'byte',
    'b', 'kb', 'mb', 'gb', 'tb', 'db', 'deg', 'k', 'mol', 'cd'
}
COMMON_UNITS_RIGHT = {
    's', 'min', 'h', 'd', 'a', 'yr', 'm', 'km', 'cm', 'mm', 'um', 'nm',
    'm2', 'm3', 'cm2', 'cm3', 'mm2', 'mm3', 'l', 'ml', 'kg', 'g', 'mg',
    't', 'mol', 'k', 'px', 'step', 'sample', 'pixel', 'frame'
}


def slash_hits(text):
    """斜杠逐处定性（含全角 ／ 与 ∕）。"""
    hits = []
    for m in re.finditer(r'[/／∕]', text):
        s = m.start()
        left, right = text[max(0, s - 12):s], text[s + 1:s + 13]
        left_token = re.search(r'([a-zA-Z0-9]+)$', left)
        right_token = re.match(r'^([a-zA-Z0-9]+)', right)
        if re.search(r'[A-Z]$', left) and re.match(r'[A-Z]', right):
            kind = '标准代号'
        elif (left_token and right_token and 
              left_token.group(1).lower() in COMMON_UNITS_LEFT and 
              right_token.group(1).lower() in COMMON_UNITS_RIGHT):
            kind = '物理单位'
        elif re.search(r'\d$', left) and re.match(r'\d', right):
            kind = '分数比例'
        elif re.search(r'[一-鿿]$', left) or re.match(r'[一-鿿]', right):
            kind = '并列替代'
        else:
            kind = '英文串联'
        hits.append((kind, '/', '…%s/%s…' % (left, right)))
    return hits

LIST_LEAD = re.compile(
    r'(包括|包含|分为|划分为|归纳为|概括为|主要有|具体为|如下|涉及|由|有)[^，。；]{0,6}$')
COLON_OK = {'比例时间', '公式引出'}


def colon_hits(text):
    """正文冒号逐处定性。

    剥除元数据前缀（摘 要：/Abstract:/关键词：…）与 http/doi 里的冒号后，
    余下冒号分四类：清单冒号（A3 空转冒号，机械列举）、正文冒号（其余），两者按密度限量，
    比例时间（3:1、08:30）与公式引出（冒号后即为公式或行末，语料显示是人写常规用法），后两类豁免。
    """
    t = re.sub(r'^(摘\s*要|摘要|Abstract|Keywords|关键词|中图分类号|文献标识码)\s*[：:]\s*', '', text)
    t = re.sub(r'https?://[^\s]+', '', t)
    t = re.sub(r'doi:[^\s]+', '', t)
    hits = []
    for m in re.finditer(r'：|:', t):
        left, right = t[:m.start()], t[m.end():m.end() + 40]
        ctx = t[max(0, m.start() - 10):m.end() + 15].replace('\n', ' ')
        if re.search(r'\d\s*$', left) and re.match(r'\s*\d', right):
            kind = '比例时间'
        elif not right.strip() or re.match(r'\s*(\$|\\\(|\\\[|\{\{)', right):
            kind = '公式引出'
        elif LIST_LEAD.search(left) or right.count('、') >= 2 or right.count('；') >= 2:
            kind = '清单冒号'
        else:
            kind = '正文冒号'
        hits.append((kind, m.group(0), ctx))
    return hits



LAW_DOC = re.compile(
    r'(法|条例|规划|意见|方案|通知|决定|办法|纲要|计划|工程|规定|准则|公约|'
    r'指南|细则|草案|报告|白皮书|导则|规程|规范|标准)$')
# 书名号左侧 30 字内出现标准编号或"国家标准"一类限定语，则该书名号内是标准正式名称，
# 按 GB/T 1.1 属规范写法，不计入密度。
STD_NEAR = re.compile(
    r'(?:GB/?T?|ISO|IEC|ANSI|ASTM|JIS|DIN|EN|LY/?T|JB/?T|QB/?T|YB/?T|SJ/?T|'
    r'GJB|HG/?T|NY/?T|CJ/?T|TB/?T|JGJ|JTG|T/[A-Z]{2,})\s*[\d][\d.:\u2014\u2013\-/]*\s*$'
    r'|(?:国家标准|国标|行业标准|团体标准|地方标准|国际标准)[^。；]{0,12}$')


def book_mark_hits(text):
    """书名号《》逐处定性。

    法规、规划与政策文件名按 GB/T 15834 必须用书名号，定性为 法规文件 并默认豁免；
    左邻标准编号或"国家标准"一类限定语的，定性为 标准名称，同样豁免；
    抽字截断产生的孤立《或》定性为 单书名号，只入报告不计密度；其余为 成对书名号。
    """
    hits = []
    paired_spans = []
    for m in re.finditer(r'《([^》]{1,60})》', text):
        inner = m.group(1)
        left = text[max(0, m.start() - 30):m.start()]
        if STD_NEAR.search(left):
            kind = '标准名称'
        elif LAW_DOC.search(inner):
            kind = '法规文件'
        else:
            kind = '成对书名号'
        hits.append((kind, inner, m.group(0)))
        paired_spans.append((m.start(), m.end()))
    for m in re.finditer(r'[《》]', text):
        if not any(start <= m.start() < end for start, end in paired_spans):
            hits.append(('单书名号', m.group(0), m.group(0)))
    return hits


# ---------------- F9：多位数分节（GB/T 15835 第 8.1、4.2 与第 10 条） ----------------
# 千分空可以是普通空格、不换行空格或各类窄空；全部识别，字符种类单独报出。
NUM_SEP = '\u0020\u00a0\u2009\u202f\u2007\u2008'
NUM_SEP_NAMES = {
    '\u0020': '普通空格', '\u00a0': '不换行空格', '\u2009': '窄空',
    '\u202f': '窄不换行空格', '\u2007': '数字空格', '\u2008': '标点空格',
}
NUM_GROUPED = re.compile(
    r'(?<![\d.])(\d{1,3}(?:[%s]\d{3})+)(?![\d])' % NUM_SEP)
NUM_PLAIN = re.compile(r'(?<![\d.])(\d{4,})(?![\d])')
# 代号、代码与序号按 GB/T 15835 第 10 条一律不分节
NUM_ID_LEFT = re.compile(
    r'(?:GB|ISO|IEC|ANSI|ASTM|JIS|DIN|GJB|LY|JB|QB|CN|No|编号|文号|代号|型号|'
    r'序号|刊号|批号|专利|第)[\dA-Za-z\s./\-\u2014\u2013]{0,12}$')
# 右邻"年"且值落在 1000~2999 才算年份；年产、年均一类量词义的"年"排除在外
NUM_YEAR_RIGHT = re.compile(
    r'^\s*(?:年(?!产|均|销|龄|限|利|度|末|初)|\u2014\d|-\d{2}-)')


def _is_year(digits, right):
    return (len(digits) == 4 and 1000 <= int(digits) <= 2999
            and NUM_YEAR_RIGHT.match(right) is not None)


def num_group_hits(text):
    """多位数分节逐处定性（F9）。

    GB/T 15835 第 8.1 允许四位以上数值以千分空分节（1200 与 1 200 都对），
    第 4.2 要求全篇体例统一，第 10 条规定代号、代码与序号即使多位也不分节。
    据此定性为 数值分节 / 数值未分节（两者皆合规，但不得混用）与
    编号分节 / 年份分节（严禁）/ 编号未分节（合规）。
    """
    hits = []
    for m in NUM_GROUPED.finditer(text):
        left = text[max(0, m.start() - 14):m.start()]
        right = text[m.end():m.end() + 4]
        names = '+'.join(NUM_SEP_NAMES[c]
                         for c in sorted({c for c in m.group(1)
                                          if c in NUM_SEP}))
        if NUM_ID_LEFT.search(left):
            kind = '编号分节'
        elif _is_year(re.sub(r'\D', '', m.group(1)), right):
            kind = '年份分节'
        else:
            kind = '数值分节'
        hits.append((kind, names, left[-8:] + m.group(0) + right))
    for m in NUM_PLAIN.finditer(text):
        left = text[max(0, m.start() - 14):m.start()]
        right = text[m.end():m.end() + 4]
        if NUM_ID_LEFT.search(left) or _is_year(m.group(1), right):
            kind = '编号未分节'
        else:
            kind = '数值未分节'
        hits.append((kind, '', left[-8:] + m.group(0) + right))
    return hits


# ---------------- F10：数值与单位的间隔体例（GB 3101 5.3.1） ----------------
# 只收常见法定计量单位与工程常用单位；百分号不是单位，按中文期刊惯例紧跟数字（另行检查）。
UNITS = set("""mm cm dm m km nm um μm s ms us μs min h d g mg kg t
Hz kHz MHz GHz Pa kPa MPa GPa N kN W kW mW V mV kV A mA J kJ
lx lm cd rad L mL mol bit px dpi ppi fps
r/min m/min m/s mm/s cm/s km/h mm/min m/s2 N/mm2 g/cm3 kg/m3
m² m³ cm² cm³ mm² mm³ km² Mm³ m/s² N/mm² g/cm³ kg/m³
℃ ℉""".split())
_UNIT_TOK = (r'(?:[A-Za-z\u03bc\u00b5\u2103\u2109\u03a9]+[\u00b2\u00b3]?)'
             r'(?:/[A-Za-z]+[\d\u00b2\u00b3]?)?')
UNIT_RE = re.compile(
    r'(?<=\d)([ \u00a0\u2009\u202f]?)'
    r'(' + _UNIT_TOK + r')'
    r'(?![A-Za-z0-9\u00b2\u00b3])')


# 汉字紧邻量值（数值+计量单位）时的前置间隔体例，例如「下游 6～8 m」与「下游6～8 m」
QTY_LEAD = re.compile(
    r'([\u4e00-\u9fff])([ \u00a0\u2009\u202f]?)'
    r'(\d[\d.]*(?:\s*[\uff5e~\u2014\u2013-]\s*\d[\d.]*)?)'
    r'[ \u00a0\u2009\u202f]?'
    r'(' + _UNIT_TOK + r')'
    r'(?![A-Za-z0-9\u00b2\u00b3])')
# 不带计量单位的数值（年份、百分数、计数、序号）与汉字之间的空格：较 2000 年、约 122%、2000 年。
# 只收留空一侧并记作「量值前留空」：tight 口径下违规，space 口径不追究不带单位的数值。
NUM_CJK_GAP = re.compile(
    r'[\u4e00-\u9fff][ \u00a0\u2009\u202f]+(?=\d)'
    r'|(?<=[\d%％])[ \u00a0\u2009\u202f]+(?=[\u4e00-\u9fff])')
NUM_UNIT_AT = re.compile(
    r'\d[\d.]*(?:\s*[\uff5e~\u2014\u2013-]\s*\d[\d.]*)?[ \u00a0\u2009\u202f]?'
    r'(' + _UNIT_TOK + r')(?![A-Za-z0-9\u00b2\u00b3])')
# 百分号紧跟数字：122%，不写 122 %
PCT_GAP = re.compile(r'(?<=\d)[ \u00a0\u2009\u202f]+(?=[%％])')
# GB/T 15834 4.13.2：连接号的规范形式是短横线 - 、一字线 ― 与浪纹线 ～（全角）；
# 半角 ~ 不是国标连接号形式。
TILDE_RANGE = re.compile(r'(?<=\d)\s*([\uff5e~])\s*(?=\d)')


def unit_space_hits(text):
    """数值与单位、汉字与量值之间的间隔体例，以及范围波浪号的全半角，逐处定性（F10）。

    GB 3101 规定数值与单位符号之间留一个空隙。留空与贴合在中文期刊里都能见到，
    本脚本不判定哪种更对，只判同一稿件内两种体例并存——那是外部工具回灌的典型痕迹。
    汉字与量值之间的空格同理：国标不作规定，但全篇必须统一。
    范围波浪号则有唯一正解：GB/T 15834 的浪纹线是全角 ～，半角 ~ 一律判违规。
    """
    hits = []
    for m in UNIT_RE.finditer(text):
        unit = m.group(2)
        if unit not in UNITS:
            continue
        i = m.start()
        while i > 0 and (text[i - 1].isdigit() or text[i - 1] == '.'):
            i -= 1
        if i > 0 and ('A' <= text[i - 1] <= 'Z' or 'a' <= text[i - 1] <= 'z'):
            continue  # 型号名里的数字（YOLOv5s、ResNet50）不是量值
        kind = '单位留空' if m.group(1) else '单位贴合'
        ctx = text[max(0, m.start() - 8):m.end() + 2]
        hits.append((kind, unit, ctx))
        tail = re.match(r'[ \u00a0\u2009\u202f]+(?=[\u4e00-\u9fff])', text[m.end():])
        if tail:  # 单位与其后汉字之间留空：6～8 m 处 → 6～8 m处
            hits.append(('量值前留空', unit,
                         text[max(0, m.start() - 6):m.end() + len(tail.group(0)) + 2]))
    for m in QTY_LEAD.finditer(text):
        if m.group(4) not in UNITS:
            continue
        kind = '量值前留空' if m.group(2) else '量值前贴合'
        hits.append((kind, m.group(4),
                     text[max(0, m.start() - 6):m.end() + 2]))
    for m in NUM_CJK_GAP.finditer(text):
        if text[m.start()] >= '\u4e00':
            u = NUM_UNIT_AT.match(text, m.end())
            if u and u.group(1) in UNITS:
                continue  # 带计量单位的量值已由 QTY_LEAD 计入
        else:
            i = m.start()
            while i > 0 and (text[i - 1].isdigit() or text[i - 1] in '.%％'):
                i -= 1
            if i > 0 and ('A' <= text[i - 1] <= 'Z' or 'a' <= text[i - 1] <= 'z'):
                continue  # 型号名里的数字（YOLOv5 模型）不是数值
        hits.append(('量值前留空', '',
                     text[max(0, m.start() - 6):m.end() + 6]))
    for m in PCT_GAP.finditer(text):
        hits.append(('百分号前留空', '%',
                     text[max(0, m.start() - 6):m.end() + 2]))
    for m in TILDE_RANGE.finditer(text):
        kind = '全角浪纹线' if m.group(1) == '\uff5e' else '半角波浪号'
        hits.append((kind, m.group(1),
                     text[max(0, m.start() - 8):m.end() + 8]))
    return hits

# ---------------- A 级：AI 腔的字面可复算特征 ----------------
INVISIBLE = re.compile(
    r'[\u200b-\u200f\u202a-\u202e\u2060\ufeff\u00ad\ue000-\uf8ff]')
INVISIBLE_NAMES = {
    '\u200b': '零宽空格', '\u200c': '零宽不连字', '\u200d': '零宽连字',
    '\u200e': '左至右标记', '\u200f': '右至左标记', '\u2060': '词连接符',
    '\ufeff': '字节序标记', '\u00ad': '软连字符',
}
ANAPHORA_HEAD = re.compile(
    r'^(这|该|此|其|它|它们|他们|上述|这些|那些|这一|这种|这类|如此|由此|与此)')
AUTHORITY = re.compile(
    r'(研究表明|研究显示|研究发现|大量研究|已有研究|众多学者|业界普遍认为|'
    r'学界普遍认为|普遍认为|众所周知|一般认为|实践证明|事实证明|文献表明|'
    r'相关研究指出|学者们指出)')
CITE_MARK = re.compile(
    r'(\{\{[^}]*\}\}|\[\d+(?:\s*[\-\u2013,，]\s*\d+)*\]|\(\d{4}[a-z]?\)|\[待补引\])')
TAIL_HOLLOW = re.compile(
    r'[，,][^。！？；]{0,14}'
    r'(?:提供了[^。！？]{0,12}(?:支撑|支持|基础|参考|思路|借鉴|可能)|'
    r'奠定了[^。！？]{0,10}基础|开辟了[^。！？]{0,10}(?:道路|方向|路径)|'
    r'具有[^。！？]{0,8}(?:重要|积极|广阔|巨大)[^。！？]{0,8}(?:意义|价值|前景|潜力)|'
    r'(?:推动|促进|助力)[^。！？]{0,12}(?:发展|进步|升级|落地))'
    r'[。！？]?$')
NOMINAL = re.compile(
    r'(?:进行|实现|开展|加以|予以|做出|作出|给予|完成|执行)[了过]?'
    r'(?:一定|充分|深入|全面|有效|系统)?[\u4e00-\u9fff]{2}(?=[，。、；：？！）)]|$)')
WEAK_PRED = re.compile(r'(?:是|为|具有|存在|属于|成为|作为|得到|受到|有着|拥有)')
CONJ = re.compile(
    r'(?:以及|并且|同时|然而|因此|从而|此外|另外|不仅|而且|与此同时|基于此|'
    r'综上所述|总之|其中|进而|并|和|与|且)')


def invisible_hits(text):
    """零宽字符、双向控制符、软连字符、Unicode 私有区（A1，严禁）。"""
    hits = []
    for m in INVISIBLE.finditer(text):
        ch = m.group(0)
        name = INVISIBLE_NAMES.get(ch)
        if name is None:
            name = '私有区字符' if '\ue000' <= ch <= '\uf8ff' else '双向控制符'
        ctx = text[max(0, m.start() - 8):m.start()] + '␣' + text[m.end():m.end() + 8]
        hits.append(('%s U+%04X' % (name, ord(ch)), ch, ctx))
    return hits


def authority_hits(text):
    """泛权威断言两侧都无引用标记（A7，限量）。

    中文综述常把引标放在断言前一句句尾（如“…优于石墨[91]。研究表明，MoS2…”），
    因此左 25 字与右 40 字任一侧有引标即视为已落引用，只有两侧皆无才计入。
    """
    hits = []
    for m in AUTHORITY.finditer(text):
        left = text[max(0, m.start() - 25):m.start()]
        right = text[m.start():m.end() + 40]
        if CITE_MARK.search(left) or CITE_MARK.search(right):
            continue
        hits.append(('无引用泛权威', m.group(0),
                     text[max(0, m.start() - 6):m.end() + 24].replace('\n', ' ')))
    return hits


def tail_hits(sents):
    """句尾空转挂尾：末句以“为…提供了参考/具有重要意义”一类套话收束（A4，告警）。"""
    return [s[-28:] for s in sents if TAIL_HOLLOW.search(s)]


def nominal_hits(text):
    """虚动词名词化：进行/实现/开展 + 双音节名词（A5，密度告警）。"""
    return [m.group(0) for m in NOMINAL.finditer(text)]


def weak_pred_count(text):
    return len(WEAK_PRED.findall(text))


def conj_count(text):
    return len(CONJ.findall(text))


def dun_chains(text):
    """逐个小句的顿号链长度（成员数）；顿号链不跨逗号、冒号、分号统计。"""
    out = []
    for seg in re.split(r'[。；！？，,：:（）()]', text):
        n = seg.count('、')
        if n:
            out.append(n + 1)
    return out


# A12：并列悬空——顿号并列链提出若干项，后文一项也没再回应
DANGLE_SUFFIX = re.compile(r'(波动|变化|影响|问题|因素|现象|特性|效应|能力)$')
DANGLE_TAIL = re.compile(r'(以及|与|和|及)')


def _dangle_key(item):
    """取并列项的可检索核心词：剥掉泛化后缀，留末 2 字（项本身不足 3 字则全取）。

    并列项常带修饰语，末 2 字才是可在后文检索的中心词。
    """
    s = re.sub(r'[^\u4e00-\u9fffA-Za-z0-9]', '', item.strip())
    s = DANGLE_SUFFIX.sub('', s) or s
    if len(s) < 2:
        return ''
    return s if len(s) <= 3 else s[-2:]


def dangling_lists(sents, tail_min=2):
    """段内并列悬空逐处（A12，告警）。

    一个顿号并列链列出若干项，若其后段内仍有 tail_min 句以上文本，
    而所有并列项的核心词在后文一次都没有再出现，则该并列提出后即被丢掉。
    收束句之后没有正文，不参与判定。
    """
    hits = []
    for i, s in enumerate(sents):
        tail = ''.join(sents[i + 1:])
        if len(sents) - i - 1 < tail_min:
            break
        for seg in re.split(r'[。；！？，,：:（）()]', s):
            if '、' not in seg:
                continue
            parts = seg.split('、')
            # 末项的并列边界止于"与/和/及"，其后是谓语而非并列成分
            parts[-1] = DANGLE_TAIL.split(parts[-1])[0]
            parts = [p for p in parts if p.strip()]
            keys = [k for k in (_dangle_key(p) for p in parts) if k]
            if len(keys) < 2:
                continue
            if not any(k in tail for k in keys):
                hits.append((seg.strip(), keys))
    return hits


def _bigrams(s):
    cjk = re.sub(r'[^\u4e00-\u9fffA-Za-z0-9]', '', s)
    return {cjk[i:i + 2] for i in range(len(cjk) - 1)}


def overlap_pairs(sents, min_len=12):
    """相邻句实词重合度（字二元组 Jaccard，A10）。"""
    out = []
    for a, b in zip(sents, sents[1:]):
        if len(a) < min_len or len(b) < min_len:
            continue
        ga, gb = _bigrams(a), _bigrams(b)
        if not ga or not gb:
            continue
        j = len(ga & gb) / float(len(ga | gb))
        out.append((round(j, 3), a[:20], b[:20]))
    return out


def length_cv(sents):
    """句长变异系数：机械对称拆句的提示项（R3，不进门禁）。"""
    lens = [len(s) for s in sents if len(s) >= 6]
    if len(lens) < 6:
        return None
    mean = sum(lens) / float(len(lens))
    if mean <= 0:
        return None
    var = sum((x - mean) ** 2 for x in lens) / float(len(lens))
    return round((var ** 0.5) / mean, 3)


# ---------------- A13～A15：可读性（"说人话"）字面线索 ----------------
# 只抓字面线索，判不了目标读者的层次，一律只告警；名词串与读者层次错位靠人读，
# 病例见 references/plain-speech.md。
# A13 压缩简称：把过程或对象压成两三个字，读者需反推。值为还原提示。
JARGON = {
    '根因': '问题的根源/原因',
    '精标注': '精细标注',
    '全量': '全部……',
    '时延': '写明是哪一段时间（推理耗时/处理耗时/端到端耗时）',
    '归并': '写明做了什么（合并/剔除重复框）',
    '粒度': '写明具体指什么（图像级/像素级、全局/局部）',
    '横比': '横向比较',
    '提精度': '提高精度',
    '期刊实例': '删去，出处由引用标记承担',
    '光机电': '光学、机械与电气',
    '落地': '部署/应用',
    '打通': '连通/衔接',
    '抓手': '写明具体手段',
}
# 人写语料里"时延""粒度"多带限定语出现（通信时延、端到端时延、细粒度），指向明确，
# 只查不带限定语的裸用；"安全量化"一类跨词误切同样排除。
JARGON_GUARD = {
    '时延': r'(?<!计算)(?<!通信)(?<!往返)(?<!到端)(?<!固有)(?<!传输)(?<!处理)(?<!推理)'
            r'(?<!排队)(?<!网络)(?<!控制)(?<!检测)(?<!单帧)(?<!平均)(?<![低高总])时延',
    '粒度': r'(?<![细粗多单])(?<!不同)(?<!同一)粒度',
    '全量': r'(?<!安)全量(?!化)',
}


def _jargon_pattern(words):
    parts = [JARGON_GUARD.get(w, re.escape(w))
             for w in sorted(words, key=len, reverse=True)]
    return re.compile('|'.join(parts)) if parts else None


JARGON_RE = _jargon_pattern(JARGON)
# A14 术语首现未解释：综述里常见、读者层次常不具备的机器学习与检测术语。
# 目标读者已熟悉的用 --known-terms 剔除，本文领域术语用 --terms-file 追加。
TERMS = (
    '经验风险', '真实风险', '结构风险', '源域', '目标域', '结构假设', '归纳偏置',
    '重采样', '重加权', '梯度贡献', '正样本', '负样本', '标签分配', '伪标签',
    '负迁移', '概念漂移', '感受野', '下采样', '过拟合', '先验', '长尾', '尾部类别',
    '特征空间', '知识蒸馏', '解耦头', '锚框', '非极大值抑制', '交并比',
)
# 同句内术语之后的解释标记：即/也就是/是指……，或括号内给出不少于 4 个汉字的中文释义
TERM_CUE = re.compile(r'^[^。！？；]{0,6}?(?:即|也就是|是指|指的是|意为|亦即)')


def jargon_hits(text, jargon_re=JARGON_RE):
    """压缩简称逐处（A13，告警）。"""
    if jargon_re is None:
        return []
    return [m.group(0) for m in jargon_re.finditer(text)]


def _term_explained(text, m):
    after = text[m.end():m.end() + 16]
    if TERM_CUE.match(after):
        return True
    if after.startswith(('（', '(')):
        inside = re.split(r'[）)]', after[1:], 1)[0]
        if len(CJK_CHAR.findall(inside)) >= 4:
            return True
    return '所谓' in text[max(0, m.start() - 8):m.start()]


def term_first_hits(text, term_re, seen):
    """术语首现未解释（A14，告警）：每个术语全文只查第一次出现，seen 跨段累积。"""
    out = []
    if term_re is None:
        return out
    for m in term_re.finditer(text):
        term = m.group(0)
        if term in seen:
            continue
        seen.add(term)
        if not _term_explained(text, m):
            s = max(0, m.start() - 10)
            out.append((term, text[s:m.end() + 14]))
    return out


# A15 抽象名词施事：抽象名词作主语，谓语却是施事动作，读者须自行补出"谁"在做
ABSTRACT_ACTOR = re.compile(
    r'(先验|范式|趋势|需求|瓶颈|痛点|本质|正偏态|稀缺|匮乏|特性|属性|主线)'
    r'(?:本身|已|则|正|也|又|便|逐渐|逐步|正在)?'
    r'(转向|推动|驱动|催生|倒逼|促成|走向|迈向|导向|引领|赋能|呼唤|主导)')


def abstract_actor_hits(text):
    """抽象名词施事逐处（A15，告警）。"""
    return [text[max(0, m.start() - 8):m.end() + 8]
            for m in ABSTRACT_ACTOR.finditer(text)]


# "权重加权"会被误切出"重加权"
TERM_GUARD = {'重加权': r'(?<!权)重加权'}


def build_term_re(known=(), extra=()):
    terms = (set(TERMS) | set(extra)) - set(known)
    if not terms:
        return None
    return re.compile('|'.join(TERM_GUARD.get(t, re.escape(t))
                               for t in sorted(terms, key=len, reverse=True)))


def family_hits(text):
    """同段内出现 ≥2 个成员的枚举族；返回 {族名: [成员, ...]}。"""
    hits = {}
    for name, pat in FAMILIES:
        members = sorted({m.group(0) for m in pat.finditer(text)})
        if len(members) >= 2:
            hits[name] = members
    return hits


# ---------------- V 级：词表驱动的用词检查 ----------------
_LOOKAROUND = re.compile(r'\(\?<?[=!][^)]*\)')


def entry_label(entry):
    """词条显示名：正则去掉 re: 前缀与前后断言，如 re:(?<!辅)助力 → 助力。"""
    if not entry.startswith('re:'):
        return entry
    return _LOOKAROUND.sub('', entry[3:])


def _compile_entry(entry):
    return re.compile(entry[3:] if entry.startswith('re:') else re.escape(entry))


def _table_rows(path):
    """读制表符分隔的词表：跳过空行与 # 注释行，返回 [(行号, [列...])]。"""
    out = []
    with io.open(path, encoding='utf-8') as fh:
        for n, line in enumerate(fh, 1):
            line = line.rstrip('\r\n')
            if not line.strip() or line.lstrip().startswith('#'):
                continue
            cols = line.split('\t')
            if len(cols) < 2:
                raise ValueError('%s 第 %d 行缺少制表符分隔的第二列' % (path, n))
            out.append((n, cols))
    return out


def load_lexicon(path):
    """类别<TAB>词条<TAB>改法 → [(类别, 词条, 正则, 改法)]。"""
    return [(c[0], c[1], _compile_entry(c[1]), c[2] if len(c) > 2 else '')
            for _, c in _table_rows(path)]


def load_terms(paths):
    """规范写法<TAB>变体1,变体2<TAB>说明 → [(规范写法显示名, [(写法, 正则)], 说明)]。

    变体以逗号分隔，因此正则写法里不能出现逗号。
    """
    groups = []
    for path in paths:
        for _, c in _table_rows(path):
            forms = [c[0]] + [x.strip() for x in re.split(r'[,，]', c[1]) if x.strip()]
            groups.append((entry_label(c[0]), [(f, _compile_entry(f)) for f in forms],
                           c[2] if len(c) > 2 else ''))
    return groups


def _longest_matches(text, patterns):
    """patterns 为 [(键, 正则)]；同一起点取最长，与已取区间重叠的让给先出现者。"""
    found = []
    for key, pat in patterns:
        for m in pat.finditer(text):
            if m.end() > m.start():
                found.append((m.start(), m.end(), key))
    found.sort(key=lambda x: (x[0], x[0] - x[1]))
    out, last = [], -1
    for s, e, key in found:
        if s >= last:
            out.append((s, e, key))
            last = e
    return out


def lexicon_hits(text, lexicon):
    """V 级逐处：[(类别, 词条, 命中文字, 上下文)]。"""
    pats = [((cat, entry), pat) for cat, entry, pat, _ in lexicon]
    return [(key[0], key[1], text[s:e], text[max(0, s - 10):e + 10])
            for s, e, key in _longest_matches(text, pats)]


# ---------------- T 级：缩写首现与术语变体 ----------------
# 前后都不贴字母、数字与连字符的全大写串：R-CNN、PBD-YOLO、YOLOv5s、mAP、LiDAR 一类型号名不计
ABBR = re.compile(r'(?<![A-Za-z0-9\-])[A-Z]{2,}\d*(?![A-Za-z\-])')
ROMAN = re.compile(r'^[IVXLCDM]+$')
# 标准代号与通用缩写默认免定义；--abbr-known 追加
ABBR_KNOWN = {
    'GB', 'GJB', 'JB', 'JG', 'JGJ', 'JTG', 'HG', 'SJ', 'YB', 'QB', 'DL', 'NY', 'LY', 'SN',
    'CJ', 'TB', 'MT', 'AQ', 'ISO', 'IEC', 'ASTM', 'EN', 'DIN', 'JIS', 'ANSI', 'IEEE', 'BS',
    'GOST', 'PDF', 'DOI', 'ISBN', 'ISSN', 'URL', 'ID', 'OK', 'NG', 'RGB', 'LED', 'CPU', 'GPU'}
ABBR_NOT_DEF = re.compile(r'^\s*(如|例如|比如|包括|含|即|见|参见|e\.g\.|i\.e\.)')


def _abbr_def(text, s, e):
    """该处缩写是否在给定义：落在全称括号里（…（adaptive spatial feature fusion，ASFF）、
    自适应空间特征融合（ASFF）），或紧跟括号给出全称（CNN（卷积神经网络））。
    是则返回括号内文字，否则返回 None。"""
    left = text[max(0, s - 80):s]
    lp = max(left.rfind('（'), left.rfind('('))
    if lp >= 0 and not re.search(r'[）)]', left[lp:]):
        m = re.match(r'[^（）()]{0,20}?[）)]', text[e:])
        if m:
            inner = left[lp + 1:] + text[s:e + m.end() - 1]
            if '、' not in inner and not ABBR_NOT_DEF.match(inner):
                return inner.strip()
    m = re.match(r'\s*[（(]([^（）()]{2,80})[）)]', text[e:])
    if m and not ABBR_NOT_DEF.match(m.group(1)) and (
            CJK_CHAR.search(m.group(1)) or len(m.group(1).split()) >= 2):
        return m.group(1).strip()
    return None


def _abbr_fullname(inner, abbr):
    """从定义括号里取英文全称（≥2 个英文词才算），小写归一后用于比对前后是否一致。"""
    words = re.findall(r'[A-Za-z]+', inner.replace(abbr, ' '))
    return ' '.join(w.lower() for w in words) if len(words) >= 2 else None


def abbr_audit(rows, known):
    """T1：正文里全大写缩写的首现与定义，返回 [(缩写, 问题, 说明)]。"""
    seen, order, seq = {}, [], 0
    for r in rows:
        for m in ABBR.finditer(r['text']):
            a = m.group(0)
            if a in known or ROMAN.match(a):
                continue
            seq += 1
            d = seen.get(a)
            if d is None:
                d = seen[a] = dict(first=(seq, r['idx'], r['lineno']), defs=[], n=0)
                order.append(a)
            d['n'] += 1
            inner = _abbr_def(r['text'], m.start(), m.end())
            if inner is not None:
                d['defs'].append((seq, r['idx'], r['lineno'], inner))
    out = []
    for a in order:
        d = seen[a]
        _, fi, fl = d['first']
        if not d['defs']:
            out.append((a, '未定义', '段%d 行%d 首现，全文未给全称（共 %d 处）'
                        % (fi, fl, d['n'])))
            continue
        dseq, di, dl, _ = d['defs'][0]
        if dseq != d['first'][0]:
            out.append((a, '定义晚于首现', '段%d 行%d 首现，段%d 行%d 才给全称'
                        % (fi, fl, di, dl)))
        if len(d['defs']) >= 2:
            out.append((a, '重复定义', '、'.join('段%d 行%d' % (x[1], x[2])
                                             for x in d['defs'])))
        names = {n for n in (_abbr_fullname(x[3], a) for x in d['defs']) if n}
        if len(names) >= 2:
            out.append((a, '全称不一致', ' / '.join(sorted(names))))
    return out


def term_variants(rows, aux, groups):
    """T2：同组 ≥2 种写法在同一稿件出现。返回 [(规范写法, {写法: 处数}, {写法: 首见位置}, 说明)]。"""
    texts = [('段%d 行%d' % (r['idx'], r['lineno']), r['text']) for r in rows] + list(aux)
    out = []
    for canon, forms, note in groups:
        pats = [(entry_label(f), pat) for f, pat in forms]
        counts, first = {}, {}
        for loc, text in texts:
            for _, _, label in _longest_matches(text, pats):
                counts[label] = counts.get(label, 0) + 1
                first.setdefault(label, loc)
        if len(counts) >= 2:
            out.append((canon, counts, first, note))
    return out


# ---------------- C 级：中文特有句式 ----------------
DE_WORDS = re.compile(r'目的|的确|的士|的话|有的|别的|似的|众矢之的')
CLAUSE_SPLIT = re.compile(r'[，,。；;：:！？、（）()“”"]')


def de_chains(text, n=3):
    """C1：单个小句（逗号、顿号、括号等之间）含 ≥n 个结构助词「的」。"""
    return [c.strip() for c in CLAUSE_SPLIT.split(text)
            if DE_WORDS.sub('', c).count('的') >= n]


STACK = re.compile(
    r'(?:不仅|不但|不只|不光)[^。！？]{1,60}?(?:而且|并且|也|还)[^。！？]{1,60}?(?:更|甚至|乃至)'
    r'|既[^。！？，]{1,30}?又[^。！？]{1,30}?(?:还|更)')


def stack_hits(sents):
    """C2：一句之内三层递进（不仅……而且……更……、既……又……还……）。"""
    return [s[:60] for s in sents if STACK.search(s)]


PARALLEL_SEG = re.compile(r'^[\u4e00-\u9fff]{3,8}$')


def parallel_runs(sents, min_run=3):
    """C3：同一句内逗号或分号分隔的 ≥min_run 个等长短句（3～8 个汉字）连排。"""
    out = []
    for s in sents:
        run = []
        for seg in re.split(r'[，,；;]', s.rstrip('。！？')):
            seg = seg.strip()
            if not PARALLEL_SEG.match(seg) or (run and len(seg) != len(run[-1])):
                if len(run) >= min_run:
                    out.append('，'.join(run))
                run = [seg] if PARALLEL_SEG.match(seg) else []
                continue
            run.append(seg)
        if len(run) >= min_run:
            out.append('，'.join(run))
    return out


# C4：节号回指。只收带编号的指代（2.2.4节、第3节、第二章），「本节」「本章」是自指，不计。
# 路线图句照计密度，逐条清单里不列：「本文结构安排如下：第1节……」，以及分句写的
# 「第2节介绍……」「本文第1节回顾……」（人写语料 194 处句首节号里 103 处是这种导航句）。
SECREF = re.compile(
    r'(?:第\s*)?\d+(?:\s*\.\s*\d+){1,2}\s*节'
    r'|第\s*\d+(?:\s*[、，,和与及至～~-]\s*\d+)*\s*[节章]'
    r'|第[一二三四五六七八九十]+(?:[节章]|部分)')
ROADMAP = re.compile(
    r'(?:安排|组织|结构|内容|章节)[^。；]{0,6}如下|(?:其余|剩余)部分'
    r'|^(?:本文)?(?:%s)[^，。；]{0,6}(?:介绍|描述|讨论|给出|总结|探讨|阐述|分析|综述|回顾|展望|'
    r'说明|评述)' % SECREF.pattern)


def secref_hits(sents):
    """C4：逐句节号回指，返回 [(句序, 回指处数, 是否路线图句, 句子)]，只含有回指的句子。"""
    out = []
    for i, s in enumerate(sents):
        n = len(SECREF.findall(s))
        if n:
            out.append((i, n, bool(ROADMAP.search(s)), s))
    return out


# C5：压缩句。≥NDE_LEN 字的句子一个结构助词「的」都没有，名词直接相接、动词被名词吃掉；
# 含 ≥3 个数值的参数报告句人写也常这样写，报告里单独标出。
NDE_LEN = 50
NUMBER = re.compile(r'\d+(?:\.\d+)?')


def de_count(text):
    """结构助词「的」的个数（目的、的确一类词内的「的」不计）。"""
    return DE_WORDS.sub('', text).count('的')


def nde_long(sents):
    """C5：返回 (≥NDE_LEN 字的句数, [(句序, 句子)] 其中不含结构助词「的」的句子)。"""
    longs = [(i, s) for i, s in enumerate(sents) if len(s) >= NDE_LEN]
    return len(longs), [(i, s) for i, s in longs if not de_count(s)]


# C6：图表括注。括号内容（剥掉 [n] 引用后）以图表号开头、含「见/如＋图表号」或以图表号收尾即判：
# （图3）（图9④）（见表1）（如图8蓝色部分所示）（R=0.019，见图4(c)）（见图13(a)[4]）。
# 括号里只是顺带提到图表的说明语（与图1(c)中标记的面）不计；允许一层内嵌括号装分图号。
PAREN_GROUP = re.compile(r'[（(]((?:[^（）()\n]|[（(][^（）()\n]{0,8}[）)]){1,60})[）)]')
_FIG_NO = r'\d+(?:[.．\-]\d+)*(?:\s*[(（]?[a-z][)）]?)?'
FIG_HEAD = re.compile(r'^\s*(?:见|参见|详见|如)?\s*[（(]?[图表]\s*\d')
FIG_SEE = re.compile(r'(?:见|如)\s*[图表]\s*\d')
FIG_TAIL = re.compile(
    r'(?:^|(?<![\u4e00-\u9fff])|(?<=[见如由据参详与和及至]))[（(]?[图表]\s*' + _FIG_NO
    + r'(?:\s*[、，,和与及至～~\-]\s*[图表]?\s*' + _FIG_NO + r')*[）)]?\s*(?:所示)?\s*$')


def _is_figparen(inner):
    inner = DOCX_CITE.sub('', inner).strip()
    return bool(FIG_HEAD.search(inner) or FIG_SEE.search(inner) or FIG_TAIL.search(inner))


def figparen_hits(text):
    """C6：逐处返回 (括号前 16 字, 括号全文)。"""
    return [(text[max(0, m.start() - 16):m.start()], m.group(0))
            for m in PAREN_GROUP.finditer(text) if _is_figparen(m.group(1))]


def pending_hits(raw):
    """P：[待补引] 逐处，返回占位前 24 字的上下文。"""
    return [raw[max(0, m.start() - 24):m.start()] for m in PENDING.finditer(raw)]


def audit(paras, long_sent=80, lexicon=(), term_re=None, jargon_re=JARGON_RE):
    if term_re is None:
        term_re = build_term_re()
    seen_terms = set()
    rows = []
    for idx, (lineno, text, raw) in enumerate(paras, 1):
        sents = sentences(text)
        prefix_rep = [a[:3] for a, b in zip(sents, sents[1:])
                      if len(a) >= 3 and len(b) >= 3 and a[:3] == b[:3]]
        n_long, nde = nde_long(sents)
        rows.append(dict(
            idx=idx, lineno=lineno, chars=len(text),
            dashes=dash_hits(text), slashes=slash_hits(text),
            quotes=quote_spans(text), book_marks=book_mark_hits(text),
            nums=num_group_hits(text), units=unit_space_hits(text),
            semi=text.count('；'),
            colon=text.count('：') + text.count(':'), colon_bad=colon_hits(text),
            dun=max([len(seg.split('、')) - 1
                     for seg in re.split(r'[。；！？，,：:（）()]', text)] or [0]),
            dun_chains=dun_chains(text),
            sents=len(sents), sent_lens=[len(s) for s in sents],
            dun_count=text.count('、'),
            opener=opener_kind(text.strip()),
            long_sent=sum(1 for s in sents if len(s) > long_sent),
            invisible=invisible_hits(text), authority=authority_hits(raw),
            pending=pending_hits(raw), lex=lexicon_hits(text, lexicon),
            de_chains=de_chains(text), stacks=stack_hits(sents),
            parallels=parallel_runs(sents), secrefs=secref_hits(sents),
            figparens=figparen_hits(text),
            long50=n_long, nde=nde, de=de_count(text),
            anaphora=bool(ANAPHORA_HEAD.match(text)) and text[:14],
            tails=tail_hits(sents), nominals=nominal_hits(text),
            weak=weak_pred_count(text), conj=conj_count(text),
            overlaps=overlap_pairs(sents), cv=length_cv(sents),
            dangles=dangling_lists(sents),
            jargon=jargon_hits(text, jargon_re),
            terms=term_first_hits(text, term_re, seen_terms),
            actors=abstract_actor_hits(text),
            fams=family_hits(text), prefix=prefix_rep, text=text))
    return rows


def audit_en(en_lines):
    """英文摘要/关键词：只判混入的中文标点（R2）；英文连字符、斜杠、直引号、分号不判。"""
    rows = []
    for lineno, text in en_lines:
        body = re.sub(r'^\*{0,2}(Abstract|Key\s*words)\*{0,2}\s*[:：]?\s*\*{0,2}\s*', '',
                      text, flags=re.I)
        sents = [s for s in re.split(r'(?<=[.])\s+', body) if s.strip()]
        bad = []
        for m in re.finditer(r'[\uff0c\u3002\uff1b\uff1a\u3001\uff08\uff09'
                             r'\u201c\u201d\u2018\u2019\uff0f\uff0d\u3010\u3011'
                             r'\u300a\u300b\u4e00-\u9fff]', body):
            ch = m.group(0)
            name = '中文字符' if '\u4e00' <= ch <= '\u9fff' else '中文标点'
            bad.append((name, ch, body[max(0, m.start() - 20):m.start() + 20]))
        rows.append(dict(
            lineno=lineno,
            label='Keywords' if re.match(r'^\*{0,2}Key\s*words', text, re.I) else 'Abstract',
            words=len(body.split()), zh=bad,
            sents=[(i, len(s.split()), s) for i, s in enumerate(sents, 1)],
            text=body))
    return rows


def cross_paragraph_families(rows):
    """同一枚举族在相邻段落各自出现 → 连续并列段落。"""
    seq = [{name for name, pat in FAMILIES if pat.search(r['text'])}
           for r in rows]
    spans = []
    for name in FAMILY_NAMES:
        run = []
        for i, present in enumerate(seq):
            if name in present:
                run.append(i)
                continue
            if len(run) >= 2:
                spans.append((name, [rows[j]['idx'] for j in run]))
            run = []
        if len(run) >= 2:
            spans.append((name, [rows[j]['idx'] for j in run]))
    return spans


_PUNCT = re.compile(r'[^\u4e00-\u9fffA-Za-z0-9]')


def repeated_spans(rows, min_len=12, limit=20):
    """跨段重复片段：≥min_len 个有效字符在两个以上段落出现。

    先剥标点与空白（重复表述常伴随标点微调），按 min_len 长度取窗口做索引，
    再把同一段内连续命中且伙伴段集合相同的窗口合并成极大片段。返回
    [(片段, [段号, ...])]，按片段长度降序，最多 limit 条。
    """
    keys, maps = [], []
    for r in rows:
        s, pos = [], []
        for i, ch in enumerate(r['text']):
            if not _PUNCT.match(ch):
                s.append(ch)
                pos.append(i)
        keys.append(''.join(s))
        maps.append(pos)

    owners = {}
    for j, s in enumerate(keys):
        for i in range(len(s) - min_len + 1):
            owners.setdefault(s[i:i + min_len], set()).add(j)

    out = []
    for j, s in enumerate(keys):
        i = 0
        while i <= len(s) - min_len:
            party = owners.get(s[i:i + min_len], set())
            if len(party) < 2:
                i += 1
                continue
            end = i + 1
            while (end <= len(s) - min_len
                   and owners.get(s[end:end + min_len]) == party):
                end += 1
            frag = s[i:end - 1 + min_len]
            out.append((frag, sorted(rows[p]['idx'] for p in party)))
            i = end
    seen, uniq = set(), []
    for frag, party in sorted(out, key=lambda x: -len(x[0])):
        sig = (frag, tuple(party))
        if sig in seen:
            continue
        seen.add(sig)
        uniq.append((frag, party))
    return uniq[:limit]


def report(path, rows, en_rows, spans, notice, fh, cfg, aux=()):
    total = sum(r['chars'] for r in rows)
    k = total / 1000.0

    def P(key, qn):
        return cfg.base['metrics'][key][qn]

    def P_opt(key, qn):
        """旧基线可能缺新指标（C4/C5），缺则返回 None。"""
        m = cfg.base['metrics'].get(key)
        return m[qn] if m else None

    def per_k(n):
        return n / k if k else 0.0

    def agg(key):
        return sum(r[key] for r in rows)

    n_quote = sum(len(r['quotes']) for r in rows)
    n_curly = sum(1 for r in rows for _, c in r['quotes'] if c)

    def kinds_of(key):
        d = {}
        for r in rows:
            for kind, _, _ in r[key]:
                d[kind] = d.get(kind, 0) + 1
        return d

    def hit_paras(key, allow):
        return sum(1 for r in rows
                   if any(kd not in allow for kd, _, _ in r[key]))

    dash_kinds, slash_kinds = kinds_of('dashes'), kinds_of('slashes')
    n_dash = sum(dash_kinds.values())
    n_dash_bad = sum(v for kd, v in dash_kinds.items()
                     if kd not in cfg.allow_dash)
    n_slash = sum(slash_kinds.values())
    n_slash_bad = sum(v for kd, v in slash_kinds.items()
                      if kd not in cfg.allow_slash)

    w = fh.write
    w('=' * 72 + '\n')
    w('源文件：%s\n' % path)
    w('正文段落 %d 段，计入统计 %d 字符（已剔除标题/图题/占位/引用标记）\n'
      % (len(rows), total))
    w('门禁（人写语料 p99 上限，每千字）：正文冒号 ≤%.2f 清单冒号 ≤%.2f '
      '双引号 ≤%.2f 破折号 ≤%.2f 概念串联 ≤%.2f 并列斜杠 ≤%.2f 书名号 ≤%.2f '
      '泛权威 ≤%.2f；分号 ≤%d/段 且 ≤%.2f/千字；跨段并列 ≤%d 组\n'
      % (cfg.colon_k, cfg.colon_list_k, cfg.quote_k, cfg.dash_k,
         cfg.dash_concept_k, cfg.slash_k, cfg.book_k, cfg.auth_k,
         cfg.semi_para, cfg.semi_k, cfg.fam_span))
    w('豁免：一字线%s 斜杠%s 枚举族%s\n'
      % (sorted(cfg.allow_dash) or '无', sorted(cfg.allow_slash) or '无',
         sorted(cfg.allow_fam) or '无'))
    w('=' * 72 + '\n\n')

    w('【汇总】\n')
    w('%-22s %8s %10s %8s\n' % ('指标', '总数', '每千字', '命中段数'))

    def line(label, n, hit):
        w('%-22s %8d %10.2f %8d\n' % (label, n, n / k if k else 0, hit))

    def dens(label, n, hit, cap, band=None):
        d = n / k if k else 0.0
        if round(d, 2) > cap:
            tag = '超限'
        elif band is not None and round(d, 2) > band:
            tag = '偏高(>人写p95 %.2f)' % band
        else:
            tag = 'ok'
        w('%-22s %8d %10.2f %8d   门禁%.2f %s\n' % (label, n, d, hit, cap, tag))

    n_dash_break = dash_kinds.get('破折号', 0) if '破折号' not in cfg.allow_dash else 0
    n_dash_concept = (dash_kinds.get('概念串联', 0)
                      if '概念串联' not in cfg.allow_dash else 0)
    n_dash_other = n_dash_bad - n_dash_break - n_dash_concept
    line('F1 一字线族 —(计数)', n_dash_bad, hit_paras('dashes', cfg.allow_dash))
    dens('   破折号(限量)', n_dash_break,
         sum(1 for r in rows if any(x == '破折号' for x, _, _ in r['dashes'])),
         cfg.dash_k, P('dash_break_k', 'p95'))
    dens('   概念串联(限量)', n_dash_concept,
         sum(1 for r in rows if any(x == '概念串联' for x, _, _ in r['dashes'])),
         cfg.dash_concept_k, P('dash_concept_k', 'p95'))
    dens('   其他一字线族(限量)', n_dash_other,
         hit_paras('dashes', cfg.allow_dash | {'破折号', '概念串联'}),
         cfg.dash_other_k, P('dash_other_k', 'p95'))
    for kind in sorted(dash_kinds):
        tag = '豁免' if kind in cfg.allow_dash else '计入'
        line('   其中 %s(%s)' % (kind, tag), dash_kinds[kind],
             sum(1 for r in rows if any(x == kind for x, _, _ in r['dashes'])))
    n_slash_pair = (slash_kinds.get('并列替代', 0)
                    if '并列替代' not in cfg.allow_slash else 0)
    n_slash_other = n_slash_bad - n_slash_pair
    line('F2 斜杠 /(计数)', n_slash_bad, hit_paras('slashes', cfg.allow_slash))
    dens('   并列替代(限量)', n_slash_pair,
         sum(1 for r in rows if any(x == '并列替代' for x, _, _ in r['slashes'])),
         cfg.slash_k, P('slash_pair_k', 'p95'))
    line('   其他斜杠(未豁免即判)', n_slash_other,
         hit_paras('slashes', cfg.allow_slash | {'并列替代'}))
    for kind in sorted(slash_kinds):
        tag = '豁免' if kind in cfg.allow_slash else '计入'
        line('   其中 %s(%s)' % (kind, tag), slash_kinds[kind],
             sum(1 for r in rows if any(x == kind for x, _, _ in r['slashes'])))
    dens('F3 双引号 (限量)', n_quote, sum(1 for r in rows if r['quotes']),
         cfg.quote_k, P('quote_k', 'p95'))
    book_kinds = kinds_of('book_marks')
    n_book_mark = book_kinds.get('成对书名号', 0)
    dens('F8 书名号 (限量)', n_book_mark,
         sum(1 for r in rows
             if any(kd == '成对书名号' for kd, _, _ in r['book_marks'])),
         cfg.book_k, P('book_k', 'p95'))
    for kind in sorted(book_kinds):
        if kind == '成对书名号':
            continue
        line('   其中 %s(不计入)' % kind, book_kinds[kind],
             sum(1 for r in rows
                 if any(kd == kind for kd, _, _ in r['book_marks'])))
    num_kinds = kinds_of('nums')
    n_num_id = num_kinds.get('编号分节', 0) + num_kinds.get('年份分节', 0)
    n_num_grouped = num_kinds.get('数值分节', 0)
    n_num_plain = num_kinds.get('数值未分节', 0)
    num_seps = sorted({nm for r in rows for kd, nm, _ in r['nums'] if nm})
    num_mixed = 1 if (n_num_grouped and n_num_plain) else 0
    line('F9 编号/年份分节(严禁)', n_num_id,
         sum(1 for r in rows
             if any(kd in ('编号分节', '年份分节') for kd, _, _ in r['nums'])))
    _g_tag = {'plain': '违规', 'grouped': '合规', 'either': '合规'}[cfg.num_group]
    _p_tag = {'plain': '合规', 'grouped': '违规', 'either': '合规'}[cfg.num_group]
    line('   其中 数值分节(%s)' % _g_tag, n_num_grouped,
         sum(1 for r in rows
             if any(kd == '数值分节' for kd, _, _ in r['nums'])))
    line('   其中 数值未分节(%s)' % _p_tag, n_num_plain,
         sum(1 for r in rows
             if any(kd == '数值未分节' for kd, _, _ in r['nums'])))
    line('   其中 编号未分节(合规)', num_kinds.get('编号未分节', 0),
         sum(1 for r in rows
             if any(kd == '编号未分节' for kd, _, _ in r['nums'])))
    w('%-22s %8s %10s %8s   %s\n'
      % ('   分节口径', cfg.num_group, '-', '-',
         {'plain': '一律写 1200', 'grouped': '一律写 1 200',
          'either': '两种皆可，不得混用（当前%s）'
                    % ('混用' if num_mixed else '一致')}[cfg.num_group]))
    unit_kinds = kinds_of('units')
    n_unit_sp = unit_kinds.get('单位留空', 0)
    n_unit_tight = unit_kinds.get('单位贴合', 0)
    unit_mixed = 1 if (n_unit_sp and n_unit_tight) else 0
    _us = {'space': ('合规', '违规'), 'tight': ('违规', '合规'),
           'either': ('合规', '合规')}[cfg.unit_space]
    line('F10 数值+单位 留空(%s)' % _us[0], n_unit_sp,
         sum(1 for r in rows if any(kd == '单位留空' for kd, _, _ in r['units'])))
    line('    数值+单位 贴合(%s)' % _us[1], n_unit_tight,
         sum(1 for r in rows if any(kd == '单位贴合' for kd, _, _ in r['units'])))
    w('%-22s %8s %10s %8s   %s\n'
      % ('   单位间隔口径', cfg.unit_space, '-', '-',
         {'space': '一律写 80 m/min', 'tight': '一律写 80m/min',
          'either': '两种皆可，不得混用（当前%s）'
                    % ('混用' if unit_mixed else '一致')}[cfg.unit_space]))
    n_lead_sp = unit_kinds.get('量值前留空', 0)
    n_lead_tight = unit_kinds.get('量值前贴合', 0)
    lead_mixed = 1 if (n_lead_sp and n_lead_tight) else 0
    n_tilde_half = unit_kinds.get('半角波浪号', 0)
    n_pct_sp = unit_kinds.get('百分号前留空', 0)
    _ql = {'tight': ('违规', '合规'), 'space': ('合规', '违规'),
           'either': ('合规', '合规')}[cfg.qty_lead]
    line('    汉字+数值 留空(%s)' % _ql[0], n_lead_sp,
         sum(1 for r in rows
             if any(kd == '量值前留空' for kd, _, _ in r['units'])))
    line('    汉字+数值 贴合(%s)' % _ql[1], n_lead_tight,
         sum(1 for r in rows
             if any(kd == '量值前贴合' for kd, _, _ in r['units'])))
    w('%-22s %8s %10s %8s   %s\n'
      % ('   前置间隔口径', cfg.qty_lead, '-', '-',
         {'tight': '一律写 下游6～8 m、较2000年', 'space': '一律写 下游 6～8 m',
          'either': '两种皆可，不得混用（当前%s）'
                    % ('混用' if lead_mixed else '一致')}[cfg.qty_lead]))
    line('    百分号前留空(严禁)', n_pct_sp,
         sum(1 for r in rows
             if any(kd == '百分号前留空' for kd, _, _ in r['units'])))
    line('F11 半角波浪号 ~(严禁)', n_tilde_half,
         sum(1 for r in rows
             if any(kd == '半角波浪号' for kd, _, _ in r['units'])))
    line('    全角浪纹线 ～(合规)', unit_kinds.get('全角浪纹线', 0),
         sum(1 for r in rows
             if any(kd == '全角浪纹线' for kd, _, _ in r['units'])))
    colon_kinds = kinds_of('colon_bad')
    n_colon_body = colon_kinds.get('正文冒号', 0)
    n_colon_list = colon_kinds.get('清单冒号', 0)
    n_colon_bad = sum(v for kd, v in colon_kinds.items() if kd not in COLON_OK)
    dens('F7 正文冒号 (限量)', n_colon_body,
         sum(1 for r in rows if any(kd == '正文冒号' for kd, _, _ in r['colon_bad'])),
         cfg.colon_k, P('colon_body_k', 'p95'))
    dens('F7 清单冒号 (限量)', n_colon_list,
         sum(1 for r in rows if any(kd == '清单冒号' for kd, _, _ in r['colon_bad'])),
         cfg.colon_list_k, P('colon_list_k', 'p95'))
    for kind in sorted(colon_kinds):
        tag = '豁免' if kind in COLON_OK else '计入'
        line('   其中 %s(%s)' % (kind, tag), colon_kinds[kind],
             sum(1 for r in rows if any(x == kind for x, _, _ in r['colon_bad'])))
    line('   其中弯引号', n_curly,
         sum(1 for r in rows if any(c for _, c in r['quotes'])))
    n_fam_bad = sum(1 for r in rows for name, mem in r['fams'].items()
                    if name not in cfg.allow_fam and len(mem) >= cfg.fam_para_min)
    line('F4 并列族 同段≥%d成员(计入)' % cfg.fam_para_min, n_fam_bad,
         sum(1 for r in rows
             if any(name not in cfg.allow_fam and len(mem) >= cfg.fam_para_min
                    for name, mem in r['fams'].items())))
    line('   其中 同段 2 成员(提示)',
         sum(1 for r in rows for name, mem in r['fams'].items()
             if name not in cfg.allow_fam and len(mem) == 2),
         sum(1 for r in rows
             if any(name not in cfg.allow_fam and len(mem) == 2
                    for name, mem in r['fams'].items())))
    line('   豁免族命中(不计入)',
         sum(1 for r in rows for name, mem in r['fams'].items()
             if name in cfg.allow_fam and len(mem) >= 2), 0)

    n_invisible = sum(len(r['invisible']) for r in rows)
    line('A1 隐形字符 (严禁)', n_invisible, sum(1 for r in rows if r['invisible']))
    n_auth = sum(len(r['authority']) for r in rows)
    dens('A7 无引用泛权威 (限量)', n_auth,
         sum(1 for r in rows if r['authority']), cfg.auth_k,
         P('authority_k', 'p95'))
    n_anaph = sum(1 for r in rows if r['anaphora'])
    line('A2 段首零回指 (告警)', n_anaph, n_anaph)
    n_tail = sum(len(r['tails']) for r in rows)
    line('A4 句尾空转挂尾 (告警)', n_tail, sum(1 for r in rows if r['tails']))
    n_nom = sum(len(r['nominals']) for r in rows)
    line('A5 虚动词名词化 (告警)', n_nom, sum(1 for r in rows if r['nominals']))
    line('A6 弱谓词 是/具有/存在', agg('weak'), sum(1 for r in rows if r['weak']))
    line('A9 连词 (告警)', agg('conj'), sum(1 for r in rows if r['conj']))
    n_dun_bad = sum(1 for r in rows for n in r['dun_chains'] if n >= cfg.dun_max)
    line('A8 顿号链 ≥%d 项 (告警)' % cfg.dun_max, n_dun_bad,
         sum(1 for r in rows if any(n >= cfg.dun_max for n in r['dun_chains'])))
    n_over = sum(1 for r in rows for j, _, _ in r['overlaps'] if j >= cfg.overlap)
    line('A10 相邻句重合 ≥%.2f (告警)' % cfg.overlap, n_over,
         sum(1 for r in rows if any(j >= cfg.overlap for j, _, _ in r['overlaps'])))
    n_dangle = sum(len(r['dangles']) for r in rows)
    line('A12 并列悬空 (告警)', n_dangle,
         sum(1 for r in rows if r['dangles']))
    n_jargon = sum(len(r['jargon']) for r in rows)
    line('A13 压缩简称 (告警)', n_jargon, sum(1 for r in rows if r['jargon']))
    n_terms = sum(len(r['terms']) for r in rows)
    line('A14 术语首现未解释 (告警)', n_terms, sum(1 for r in rows if r['terms']))
    n_actor = sum(len(r['actors']) for r in rows)
    line('A15 抽象名词施事 (告警)', n_actor, sum(1 for r in rows if r['actors']))

    # V/T/C/P：词表、缩写与术语、中文句式、待补引
    lex_base = cfg.base.get('lexicon', {})

    def lex_cap(kind, key):
        return lex_base.get(kind, {}).get(key, {}).get('p99', 0.0)

    word_n, cat_n = {}, {}
    for r in rows:
        for cat, entry, _, _ in r['lex']:
            word_n[(cat, entry)] = word_n.get((cat, entry), 0) + 1
            cat_n[cat] = cat_n.get(cat, 0) + 1
    lex_cat_bad = [(cat, n, per_k(n), lex_cap('categories', cat))
                   for cat, n in sorted(cat_n.items(), key=lambda x: -x[1])
                   if n >= 2 and round(per_k(n), 2) > lex_cap('categories', cat)]
    lex_word_hi = [(cat, entry, n) for (cat, entry), n in word_n.items()
                   if n >= 2 and round(per_k(n), 2) > lex_cap('words', entry)]
    lex_stacks = []
    for r in rows:
        by_cat = {}
        for cat, entry, hit, _ in r['lex']:
            by_cat.setdefault(cat, []).append(hit)
        lex_stacks.extend((r, cat, hits) for cat, hits in sorted(by_cat.items())
                          if len(hits) >= 2)
    aux_lex = [(loc, lexicon_hits(text, cfg.lexicon)) for loc, text in aux]
    aux_lex = [(loc, hits) for loc, hits in aux_lex if hits]
    n_lex = sum(cat_n.values())
    line('V 用词词表 (类别超p99进门禁)', n_lex, sum(1 for r in rows if r['lex']))
    line('   其中 表格/图题/表注(不计密度)', sum(len(h) for _, h in aux_lex), 0)
    abbrs = abbr_audit(rows, cfg.abbr_known)
    variants = term_variants(rows, aux, cfg.terms)
    line('T1 缩写首现 (告警)', len(abbrs), len({a for a, _, _ in abbrs}))
    line('T2 术语变体 (告警)', len(variants), 0)
    n_de = sum(len(r['de_chains']) for r in rows)
    n_stack = sum(len(r['stacks']) for r in rows)
    n_par = sum(len(r['parallels']) for r in rows)
    n_pending = sum(len(r['pending']) for r in rows)
    n_figparen = sum(len(r['figparens']) for r in rows)
    line('C1 的字链 ≥3 (告警)', n_de, sum(1 for r in rows if r['de_chains']))
    line('C2 递进堆叠 (告警)', n_stack, sum(1 for r in rows if r['stacks']))
    line('C3 等长排比 (告警)', n_par, sum(1 for r in rows if r['parallels']))
    n_secref = sum(n for r in rows for _, n, _, _ in r['secrefs'])
    secref_k = per_k(n_secref)
    secref_paras = sum(1 for r in rows if r['secrefs'])
    n_long50 = agg('long50')
    n_nde = sum(len(r['nde']) for r in rows)
    nde_pct = 100.0 * n_nde / n_long50 if n_long50 else 0.0
    de_k = per_k(agg('de'))
    nde_p95 = P_opt('nde_long', 'p95')
    if cfg.secref_k is None:
        line('C4 节号回指(基线未含)', n_secref, secref_paras)
    else:
        dens('C4 节号回指(限量)', n_secref, secref_paras, cfg.secref_k,
             P_opt('secref_k', 'p95'))
    if cfg.nde_long is None:
        tag = '基线未含'
    elif round(nde_pct, 1) > cfg.nde_long:
        tag = '门禁%.1f%% 超限' % cfg.nde_long
    elif nde_p95 is not None and round(nde_pct, 1) > nde_p95:
        tag = '门禁%.1f%% 偏高(>人写p95 %.1f%%)' % (cfg.nde_long, nde_p95)
    else:
        tag = '门禁%.1f%% ok' % cfg.nde_long
    w('%-22s %8d %9.1f%% %8d   %s\n'
      % ('C5 无「的」长句(占比)', n_nde, nde_pct, sum(1 for r in rows if r['nde']), tag))
    line('C6 图表括注 (严禁)', n_figparen, sum(1 for r in rows if r['figparens']))
    line('P 待补引 (投稿前清零)', n_pending, sum(1 for r in rows if r['pending']))
    line('I4 句首重复 (仅统计)', sum(len(r['prefix']) for r in rows),
         sum(1 for r in rows if r['prefix']))
    line('L1 分号 ；(限量)', agg('semi'), sum(1 for r in rows if r['semi']))
    line('I1 冒号 ：', agg('colon'), sum(1 for r in rows if r['colon']))
    repeats = repeated_spans(rows)
    line('I2 跨段重复片段 ≥12 字', len(repeats),
         len({i for _, party in repeats for i in party}))
    flat = [r for r in rows if r['cv'] is not None and r['cv'] < cfg.cv_min]
    w('I3 单段分号最大值 %d；顿号链最长 %d 项；>%d 字长句 %d 句\n'
      % (max([r['semi'] for r in rows] or [0]),
         max([r['dun'] for r in rows] or [0]) + 1, cfg.long_sent, agg('long_sent')))
    w('I5 句长变异系数 <%.2f 的段：%s（机械对称拆句提示，不进门禁）\n'
      % (cfg.cv_min,
         '、'.join('段%d(%.2f)' % (r['idx'], r['cv']) for r in flat) or '无'))
    w('I6 密度：虚动词名词化 %.2f/千字，弱谓词 %.2f/千字，连词 %.2f/千字\n'
      % (n_nom / k if k else 0, agg('weak') / k if k else 0,
         agg('conj') / k if k else 0))

    long_rows = [r for r in rows if r['sent_lens']]
    all_lens = [x for r in long_rows for x in r['sent_lens'] if x >= 6]
    n_long80 = sum(1 for x in all_lens if x > 80)
    pct_long80 = 100.0 * n_long80 / len(all_lens) if all_lens else 0.0
    mean_len = sum(all_lens) / float(len(all_lens)) if all_lens else 0.0
    cv_all = 0.0
    if len(all_lens) >= 10 and mean_len > 0:
        var = sum((x - mean_len) ** 2 for x in all_lens) / float(len(all_lens))
        cv_all = (var ** 0.5) / mean_len
    para_chars = total / float(len(rows)) if rows else 0.0
    dun_k = agg('dun_count') / k if k else 0.0
    colon_body_k = n_colon_body / k if k else 0.0
    long_paras = [r for r in rows if r['chars'] >= 40]
    n_opener_ok = sum(1 for r in long_paras if r['opener'] in OPENER_OK)
    opener_ratio = 100.0 * n_opener_ok / len(long_paras) if long_paras else 100.0
    opener_kinds = {}
    for r in long_paras:
        opener_kinds[r['opener']] = opener_kinds.get(r['opener'], 0) + 1

    w('A11 段首实体立论率 %.1f%%（≥40 字段落 %d 段，下限 %.1f%%）：%s\n'
      % (opener_ratio, len(long_paras), cfg.opener_min,
         '、'.join('%s%d' % (kd, n) for kd, n in sorted(
             opener_kinds.items(), key=lambda x: -x[1]))))
    w('\n')

    vals = dict(
        colon_body_k=colon_body_k, colon_list_k=per_k(n_colon_list), quote_k=per_k(n_quote),
        dash_break_k=per_k(n_dash_break), dash_concept_k=per_k(n_dash_concept),
        dash_other_k=per_k(n_dash_other), slash_pair_k=per_k(n_slash_pair),
        book_k=per_k(n_book_mark), authority_k=per_k(n_auth), semi_k=per_k(agg('semi')),
        dun_k=dun_k, conj_k=per_k(agg('conj')), weak_k=per_k(agg('weak')),
        nominal_k=per_k(n_nom))
    floors = []
    for key in BAND_METRICS:
        lo, hi, v = P(key, 'p10'), P(key, 'p95'), vals[key]
        if lo and round(v, 2) < lo:
            floors.append(('过净', '%s %.2f < 人写 p10 %.2f/千字' % (BAND_LABELS[key], v, lo)))
        elif round(v, 2) > hi:
            floors.append(('过密', '%s %.2f > 人写 p95 %.2f/千字' % (BAND_LABELS[key], v, hi)))
    if pct_long80 < P('long80', 'p10'):
        floors.append(('过净', '>80 字长句占比 %.1f%% < 人写 p10 %.1f%%'
                       % (pct_long80, P('long80', 'p10'))))
    if cv_all and cv_all < P('sent_cv', 'p10'):
        floors.append(('过净', '全文句长变异系数 %.2f < 人写 p10 %.2f'
                       % (cv_all, P('sent_cv', 'p10'))))
    if para_chars < P('para_chars', 'p10'):
        floors.append(('过净', '段均字数 %.0f < 人写 p10 %.0f'
                       % (para_chars, P('para_chars', 'p10'))))
    if opener_ratio < cfg.opener_min:
        floors.append(('过净', '段首实体立论率 %.1f%% < 下限 %.1f%%'
                       % (opener_ratio, cfg.opener_min)))
    w('【B 级 分布偏离】（落在人写语料 p10~p95 之外；不进门禁，供人工判断）\n')
    if floors:
        for side, f in floors:
            w('  [%s] %s\n' % (side, f))
    else:
        w('  无（各项落在人写分布内）\n')
    dev = cfg.base.get('deviation', {})
    n_papers = sum(dev.values())

    def share(lo, hi):
        return (100.0 * sum(v for kk, v in dev.items() if lo <= int(kk) <= hi) / n_papers
                if n_papers else 0.0)

    n_dev = len(floors)
    if n_dev <= 1:
        w('  小结：偏离 %d 项，落在人写主群（%d 篇中 %.1f%% 为 0~1 项）\n'
          % (n_dev, n_papers, share(0, 1)))
    elif n_dev <= 3:
        w('  小结：偏离 %d 项，人写语料 %.1f%% 落在 2~3 项，可接受\n' % (n_dev, share(2, 3)))
    else:
        w('  小结：偏离 %d 项，人写语料仅 %.1f%% 达到 4 项以上，建议按上列逐项回填\n'
          % (n_dev, share(4, 10 ** 6)))
    w('\n')

    w('【I2 跨段重复片段】（≥12 有效字符在两段以上出现；仅统计，不判违规）\n')
    if repeats:
        for frag, party in repeats:
            w('  %2d 字 段 %s：%s\n'
              % (len(frag), '、'.join(str(i) for i in party), frag))
    else:
        w('  无\n')
    w('\n')

    w('【F8 书名号 逐处】（标准与专著改平叙加引用标记；按密度判，门禁 %.2f/千字）\n'
      % cfg.book_k)
    hit_bm = False
    for r in rows:
        for kind, content, raw in r['book_marks']:
            hit_bm = True
            tag = '计入' if kind == '成对书名号' else '豁免'
            w('  段%-3d 行%-4d [%s|%s] %s\n'
              % (r['idx'], r['lineno'], kind, tag, raw))
    if not hit_bm:
        w('  无\n')
    w('\n')
    w('【F9 多位数分节 逐处】（口径 %s：%s；代号、序号与年份一律不分节）\n'
      % (cfg.num_group,
         {'plain': '四位以上数值一律不分节，写 1200',
          'grouped': '四位以上数值一律分节，写 1 200',
          'either': '两种写法皆可，但全篇不得混用'}[cfg.num_group]))
    hit_num = False
    bad_kinds = {'编号分节', '年份分节'}
    if cfg.num_group == 'plain':
        bad_kinds.add('数值分节')
    elif cfg.num_group == 'grouped':
        bad_kinds.add('数值未分节')
    for r in rows:
        for kind, sep, raw in r['nums']:
            hit_num = True
            tag = '违规' if kind in bad_kinds else '合规'
            w('  段%-3d 行%-4d [%s|%s]%s ...%s...\n'
              % (r['idx'], r['lineno'], kind, tag,
                 ('（%s）' % sep) if sep else '', raw))
    if not hit_num:
        w('  无\n')
    if cfg.num_group == 'either' and num_mixed:
        w('  体例混用：%d 处分节与 %d 处未分节并存，须统一（GB/T 15835 第 4.2）\n'
          % (n_num_grouped, n_num_plain))
    if len(num_seps) > 1:
        w('  分隔符混用：%s；同一稿件只用一种\n' % '、'.join(num_seps))
    w('\n')
    unit_bad = set()
    if cfg.unit_space == 'space':
        unit_bad.add('单位贴合')
    elif cfg.unit_space == 'tight':
        unit_bad.add('单位留空')
    elif unit_mixed:
        unit_bad |= {'单位留空', '单位贴合'}
    if cfg.qty_lead == 'tight':
        unit_bad.add('量值前留空')
    elif cfg.qty_lead == 'space':
        unit_bad.add('量值前贴合')
    elif lead_mixed:
        unit_bad |= {'量值前留空', '量值前贴合'}
    unit_bad |= {'半角波浪号', '百分号前留空'}
    if any(kd in unit_bad for r in rows for kd, _, _ in r['units']):
        w('【F10/F11 量值书写 逐处】（单位间隔口径 %s，前置间隔口径 %s，'
          '范围波浪号须为全角 ～）\n' % (cfg.unit_space, cfg.qty_lead))
        for r in rows:
            for kind, unit, ctx in r['units']:
                tag = '违规' if kind in unit_bad else '合规'
                w('  段%-3d 行%-4d [%s|%s] ...%s...\n'
                  % (r['idx'], r['lineno'], kind, tag, ctx))
        w('\n')

    w('【F7 冒号 逐处】（清单冒号=A3 空转冒号；比例时间与公式引出豁免，其余按密度限量）\n')
    hit_colon = False
    for r in rows:
        for kind, symbol, ctx in r['colon_bad']:
            hit_colon = True
            tag = '豁免' if kind in COLON_OK else '违规'
            w('  段%-3d 行%-4d [%s|%s] ...%s...\n'
              % (r['idx'], r['lineno'], kind, tag, ctx))
    if not hit_colon:
        w('  无\n')
    w('\n')

    w('【F1 一字线族 逐处】（U+2014 及同族；按 kind 分别限量，豁免项标注在行内）\n')
    hit = False
    for r in rows:
        for kind, run, ctx in r['dashes']:
            hit = True
            tag = '豁免' if kind in cfg.allow_dash else '违规'
            w('  段%-3d 行%-4d [%s|%s] %s\n'
              % (r['idx'], r['lineno'], kind, tag, ctx))
    if not hit:
        w('  无\n')
    w('\n')

    w('【F2 斜杠 逐处】\n')
    hit = False
    for r in rows:
        for kind, _, ctx in r['slashes']:
            hit = True
            tag = '豁免' if kind in cfg.allow_slash else '违规'
            w('  段%-3d 行%-4d [%s|%s] %s\n'
              % (r['idx'], r['lineno'], kind, tag, ctx))
    if not hit:
        w('  无\n')
    w('\n')

    w('【F3 双引号 逐处】（判定用：是否为原文引用或术语首现）\n')
    hit = False
    for r in rows:
        for content, curly in r['quotes']:
            hit = True
            kind = '弯引号' if curly else '直角'
            tag = ('造词/强调' if re.search(r'[\u2014、]', content)
                   or len(content) > 8 else '短词')
            w('  段%-3d 行%-4d [%s|%s] %s\n'
              % (r['idx'], r['lineno'], kind, tag, content))
    if not hit:
        w('  无\n')
    w('\n')

    w('【F4 连续并列 — 同段内】（同段 ≥%d 成员判违规，2 成员仅提示）\n'
      % cfg.fam_para_min)
    hit = False
    for r in rows:
        for name, members in sorted(r['fams'].items()):
            hit = True
            if name in cfg.allow_fam:
                tag = '豁免'
            elif len(members) >= cfg.fam_para_min:
                tag = '违规'
            else:
                tag = '提示'
            w('  段%-3d 行%-4d [%s] %s → %s\n'
              % (r['idx'], r['lineno'], tag, name, '、'.join(members)))
    if not hit:
        w('  无\n')
    w('\n')

    w('【F4 连续并列 — 跨相邻段落】\n')
    if spans:
        for name, idxs in spans:
            w('  [%s] %s → 段 %s\n'
              % ('豁免' if name in cfg.allow_fam else '违规',
                 name, '、'.join(str(i) for i in idxs)))
    else:
        w('  无\n')
    w('\n')

    w('【F5 相邻句首重复】\n')
    hit = False
    for r in rows:
        if r['prefix']:
            hit = True
            w('  段%-3d 行%-4d %s\n'
              % (r['idx'], r['lineno'], '、'.join(r['prefix'])))
    if not hit:
        w('  无\n')
    w('\n')

    w('【L1 分号密度 逐段（仅列 ≥2 处的段）】\n')
    hit = False
    for r in sorted(rows, key=lambda x: -x['semi']):
        if r['semi'] < 2:
            continue
        hit = True
        w('  段%-3d 行%-4d %d 处 / %d 字（%.1f 每千字）\n'
          % (r['idx'], r['lineno'], r['semi'], r['chars'],
             r['semi'] * 1000.0 / r['chars']))
    if not hit:
        w('  无\n')
    w('\n')
    w('【A1 隐形字符 逐处】（零宽/双向控制/软连字符/私有区；一律违规）\n')
    hit = False
    for r in rows:
        for name, _, ctx in r['invisible']:
            hit = True
            w('  段%-3d 行%-4d [%s|违规] %s\n' % (r['idx'], r['lineno'], name, ctx))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A7 无引用泛权威 逐处】（断言后 40 字内无引用标记；门禁 %.2f/千字）\n'
      % cfg.auth_k)
    hit = False
    for r in rows:
        for _, phrase, ctx in r['authority']:
            hit = True
            w('  段%-3d 行%-4d [%s|违规] %s\n'
              % (r['idx'], r['lineno'], phrase, ctx))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A2 段首零回指】（段首代词无先行词，需回填名词；告警）\n')
    hit = False
    for r in rows:
        if r['anaphora']:
            hit = True
            w('  段%-3d 行%-4d %s…\n' % (r['idx'], r['lineno'], r['anaphora']))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A4 句尾空转挂尾】（“为…提供了参考/具有重要意义”式收束；告警）\n')
    hit = False
    for r in rows:
        for t in r['tails']:
            hit = True
            w('  段%-3d 行%-4d …%s\n' % (r['idx'], r['lineno'], t))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A5 虚动词名词化】（进行/实现/开展 + 双音节名词；告警）\n')
    hit = False
    for r in rows:
        if r['nominals']:
            hit = True
            w('  段%-3d 行%-4d %s\n'
              % (r['idx'], r['lineno'], '、'.join(r['nominals'])))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A8 顿号链 ≥%d 项】（长枚举链，建议拆句或改述；告警）\n' % cfg.dun_max)
    hit = False
    for r in rows:
        long_chains = [n for n in r['dun_chains'] if n >= cfg.dun_max]
        if long_chains:
            hit = True
            w('  段%-3d 行%-4d 链长 %s\n'
              % (r['idx'], r['lineno'], '、'.join(str(n) for n in long_chains)))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A10 相邻句实词重合 ≥%.2f】（同义反复，删一留一；告警）\n' % cfg.overlap)
    hit = False
    for r in rows:
        for j, a, b in r['overlaps']:
            if j < cfg.overlap:
                continue
            hit = True
            w('  段%-3d 行%-4d %.2f  %s… / %s…\n'
              % (r['idx'], r['lineno'], j, a, b))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A12 并列悬空】（顿号并列提出若干项，段内后文一项也没再回应；告警）\n')
    hit = False
    for r in rows:
        for seg, keys in r['dangles']:
            hit = True
            w('  段%-3d 行%-4d 未回应 %s：%s\n'
              % (r['idx'], r['lineno'], '/'.join(keys), seg[:48]))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A13 压缩简称】（过程或对象压成两三个字，读者需反推；告警）\n')
    hit = False
    for r in rows:
        for word in r['jargon']:
            hit = True
            w('  段%-3d 行%-4d %s → %s\n'
              % (r['idx'], r['lineno'], word, JARGON[word]))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A14 术语首现未解释】（全文首次出现，同句无"即/是指"或中文括注；'
      '读者已熟悉的用 --known-terms 剔除；告警）\n')
    hit = False
    for r in rows:
        for term, ctx in r['terms']:
            hit = True
            w('  段%-3d 行%-4d %s：…%s…\n' % (r['idx'], r['lineno'], term, ctx))
    if not hit:
        w('  无\n')
    w('\n')

    w('【A15 抽象名词施事】（抽象名词作主语做施事动作，须补出真正的动作者；告警）\n')
    hit = False
    for r in rows:
        for ctx in r['actors']:
            hit = True
            w('  段%-3d 行%-4d …%s…\n' % (r['idx'], r['lineno'], ctx))
    if not hit:
        w('  无\n')
    w('\n')

    fix = {entry: f for _, entry, _, f in cfg.lexicon}
    w('【V 用词词表】（%s；同段同类 ≥2 处告警「同段堆叠」；单词 ≥2 处且超该词人写 p99 告警'
      '「超人写」；类别 ≥2 处且超类别人写 p99 判违规；表格、图题、表注只逐处列出）\n'
      % cfg.lexicon_path)
    if not lex_base.get('categories'):
        w('  注意：基线未含词表统计，阈值按 0 处理；先跑 corpus/build_baseline.py\n')
    if cat_n:
        for cat, n in sorted(cat_n.items(), key=lambda x: -x[1]):
            cap = lex_cap('categories', cat)
            tag = '违规' if any(c == cat for c, _, _, _ in lex_cat_bad) else 'ok'
            w('  类别 %s：%d 处，%.2f/千字，人写 p99 %.2f  %s\n' % (cat, n, per_k(n), cap, tag))
        hi = {(c, e) for c, e, _ in lex_word_hi}
        for (cat, entry), n in sorted(word_n.items(), key=lambda x: (-x[1], x[0])):
            w('    %s〔%s〕%d 处，人写 p99 %.2f/千字%s → %s\n'
              % (entry_label(entry), cat, n, lex_cap('words', entry),
                 '  超人写' if (cat, entry) in hi else '', fix[entry]))
        for r, cat, hits in lex_stacks:
            w('  [同段堆叠] 段%-3d 行%-4d %s：%s\n'
              % (r['idx'], r['lineno'], cat, '、'.join(hits)))
        for r in rows:
            for cat, entry, hit, ctx in r['lex']:
                w('  段%-3d 行%-4d [%s|%s] …%s…\n' % (r['idx'], r['lineno'], cat, hit, ctx))
    else:
        w('  正文无命中\n')
    for loc, hits in aux_lex:
        for cat, entry, hit, ctx in hits:
            w('  %s [%s|%s] …%s…\n' % (loc, cat, hit, ctx))
    w('\n')

    w('【T1 缩写首现】（全大写缩写首次出现须给全称；标准代号与 --abbr-known 免定义；告警）\n')
    for a, kind, note in abbrs:
        w('  %s [%s] %s\n' % (a, kind, note))
    if not abbrs:
        w('  无\n')
    w('\n')

    w('【T2 术语变体】（同组写法在同一稿件并用；统一为规范写法；告警）\n')
    for canon, counts, first, note in variants:
        w('  %s：%s%s\n' % (canon, ' / '.join(
            '%s %d 处（首见 %s）' % (f, n, first[f])
            for f, n in sorted(counts.items(), key=lambda x: -x[1])),
            '  ' + note if note else ''))
    if not variants:
        w('  无\n')
    w('\n')

    def c_ref(key):
        m = cfg.base['metrics'].get(key)
        return '人写 p95 %.2f/千字' % m['p95'] if m else '基线未含'

    w('【C 中文句式】（告警；C1 %s，C2 %s，C3 %s）\n'
      % (c_ref('de_chain_k'), c_ref('stack_k'), c_ref('parallel_k')))
    hit = False
    for r in rows:
        for c in r['de_chains']:
            hit = True
            w('  段%-3d 行%-4d [C1 的×%d] %s\n'
              % (r['idx'], r['lineno'], DE_WORDS.sub('', c).count('的'), c[:60]))
        for s in r['stacks']:
            hit = True
            w('  段%-3d 行%-4d [C2 递进堆叠] %s\n' % (r['idx'], r['lineno'], s))
        for s in r['parallels']:
            hit = True
            w('  段%-3d 行%-4d [C3 等长排比] %s\n' % (r['idx'], r['lineno'], s))
    if not hit:
        w('  无\n')
    w('\n')

    sec_cap = ('门禁 %.2f/千字' % cfg.secref_k) if cfg.secref_k is not None else '基线未含'
    w('【C4 节号回指】%d 处，%.2f/千字（%s，人写 p50 %s）\n'
      % (n_secref, secref_k, sec_cap, P_opt('secref_k', 'p50')))
    w('  逐条列出段首句带节号、或一句 ≥2 处回指的句子；路线图句（本文结构安排如下…、第2节介绍…）不列。\n'
      '  改法：删掉编号，复述被指的内容（「前述按批次更新的方法」）；确需回查公式或定义时写成状语'
      '「如2.3节所述」，不作主语、不放段首。\n')
    single = []
    for r in rows:
        for i, n, roadmap, s in r['secrefs']:
            if roadmap:
                continue
            if i == 0 or n >= 2:
                w('  段%-3d 行%-4d [%s] %s\n'
                  % (r['idx'], r['lineno'], '段首' if i == 0 else '一句%d处' % n, s[:70]))
            else:
                single.append(r['idx'])
    if single:
        w('  其余单处回指 %d 处：段 %s\n' % (len(single), '、'.join(map(str, single))))
    if not n_secref:
        w('  无\n')
    w('\n')

    nde_cap = ('门禁 %.1f%%' % cfg.nde_long) if cfg.nde_long is not None else '基线未含'
    de_p05, de_p50 = P_opt('de_k', 'p05'), P_opt('de_k', 'p50')
    w('【C5 压缩句】≥%d 字句 %d 句，其中不含「的」%d 句，占 %.1f%%（%s，人写 p50 %s%%）；'
      '「的」%.1f/千字（人写 p05 %s，p50 %s）\n'
      % (NDE_LEN, n_long50, n_nde, nde_pct, nde_cap, P_opt('nde_long', 'p50'),
         de_k, de_p05, de_p50))
    if de_p05 is not None and de_k < de_p05:
        w('  提示：「的」密度低于人写 p05，全稿偏压缩，名词直接相接、动词被名词化。\n')
    w('  改法：一句只做一件事；名词串还原成主语＋动词，补回「的」。'
      '[数值] 为含 ≥3 个数的参数报告句，多可保留。\n')
    for r in rows:
        for i, s in r['nde']:
            n_num = len(NUMBER.findall(SECREF.sub('', s)))
            tags = [t for t, on in (('段首', i == 0), ('数值', n_num >= 3)) if on]
            w('  段%-3d 行%-4d [%s] %s\n'
              % (r['idx'], r['lineno'], ','.join(tags) or '句中', s[:70]))
    if not n_nde:
        w('  无\n')
    w('\n')

    w('【C6 图表括注】%d 处（严禁；人写 267 篇 77 篇出现，属书写口径，不按基线放行）\n' % n_figparen)
    w('  改法：图表作主语或谓语主干，写进句子——「表1列出了…」「…如图3所示」；'
      '括号只留给缩写全称与公式参数。\n')
    for r in rows:
        for ctx, hit in r['figparens']:
            w('  段%-3d 行%-4d …%s%s\n' % (r['idx'], r['lineno'], ctx, hit))
    if not n_figparen:
        w('  无\n')
    w('\n')

    w('【P 待补引】（占位视为已落引用，A7 不再追究；投稿前须全部补齐）\n')
    for r in rows:
        for ctx in r['pending']:
            w('  段%-3d 行%-4d …%s[待补引]\n' % (r['idx'], r['lineno'], ctx))
    if not n_pending:
        w('  无\n')
    w('\n')

    w('【排除行提醒】（标题/图题/英文摘要，不计入门禁）\n')
    if notice:
        for lineno, text in notice:
            marks = []
            if re.search(r'[\u2014\u2013\u2212\uff0d]', text):
                marks.append('一字线')
            if re.search(r'[/\uff0f\u2215]', text):
                marks.append('斜杠')
            if re.search(r'["\u201c]', text):
                marks.append('双引号')
            w('  行%-4d [%s] %s\n' % (lineno, '+'.join(marks), text[:60]))
    else:
        w('  无\n')
    w('\n')

    w('【F6 英文摘要/关键词】（只判混入的中文标点与中文字符；英文连字符、斜杠、'
      '直引号、分号按英文规范不判）\n')
    en_bad = 0
    for e in en_rows:
        w('  行%-4d %s：%d 词，%d 句\n'
          % (e['lineno'], e['label'], e['words'], len(e['sents'])))
        for name, ch, ctx in e['zh']:
            en_bad += 1
            w('    [%s|违规] U+%04X %s\n' % (name, ord(ch), ctx))
        for i, nw, s in e['sents']:
            if nw > cfg.en_words:
                w('    [长句提示] 句%d %d 词 > %d：%s…\n'
                  % (i, nw, cfg.en_words, s[:56]))
    if not en_rows:
        w('  无\n')
    w('\n')

    return dict(paras=len(rows), chars=total, dash=n_dash_bad,
                dash_all=n_dash, slash=n_slash_bad, slash_all=n_slash,
                dash_break=n_dash_break, dash_concept=n_dash_concept,
                dash_other=n_dash_other, slash_pair=n_slash_pair,
                slash_other=n_slash_other,
                dash_other_k=n_dash_other / k if k else 0.0,
                dash_break_k=n_dash_break / k if k else 0.0,
                dash_concept_k=n_dash_concept / k if k else 0.0,
                slash_pair_k=n_slash_pair / k if k else 0.0,
                quote=n_quote, quote_k=n_quote / k if k else 0.0,
                curly=n_curly, book_mark=n_book_mark,
                book_k=n_book_mark / k if k else 0.0,
                num_id=n_num_id, num_mixed=num_mixed,
                num_grouped=n_num_grouped, num_plain=n_num_plain,
                num_sep_mixed=1 if len(num_seps) > 1 else 0,
                unit_mixed=unit_mixed, lead_mixed=lead_mixed,
                unit_sp=n_unit_sp, unit_tight=n_unit_tight,
                lead_sp=n_lead_sp, lead_tight=n_lead_tight,
                tilde_half=n_tilde_half, pct_sp=n_pct_sp,
                fam=n_fam_bad,
                fam_span=sum(1 for n, _ in spans if n not in cfg.allow_fam),
                prefix=sum(len(r['prefix']) for r in rows),
                semi=agg('semi'),
                semi_max=max([r['semi'] for r in rows] or [0]),
                semi_k=agg('semi') / k if k else 0.0,
                colon=agg('colon'), colon_bad=n_colon_bad,
                colon_body=n_colon_body, colon_list=n_colon_list,
                colon_body_k=colon_body_k,
                colon_list_k=n_colon_list / k if k else 0.0,
                en_bad=en_bad,
                repeat=len(repeats),
                invisible=n_invisible, authority=n_auth,
                authority_k=n_auth / k if k else 0.0,
                anaphora=n_anaph,
                tail=n_tail, nominal=n_nom, nominal_k=n_nom / k if k else 0.0,
                weak_k=per_k(agg('weak')), conj_k=per_k(agg('conj')),
                dun_k=dun_k, long80=pct_long80, sent_cv=cv_all,
                para_chars=para_chars, opener=opener_ratio, floors=floors,
                dun_long=n_dun_bad, overlap=n_over, flat=len(flat),
                jargon=n_jargon, terms=n_terms, actors=n_actor,
                span_all=len(spans),
                lex=n_lex, lex_cat_bad=lex_cat_bad, lex_word_hi=len(lex_word_hi),
                lex_stack=len(lex_stacks),
                lex_cat_k={cat: per_k(n) for cat, n in cat_n.items()},
                lex_word_k={entry: per_k(n) for (_, entry), n in word_n.items()},
                abbr=len(abbrs), term_var=len(variants),
                de_chain=n_de, stack=n_stack, parallel=n_par,
                de_chain_k=per_k(n_de), stack_k=per_k(n_stack), parallel_k=per_k(n_par),
                secref=n_secref, secref_k=secref_k, long50=n_long50, nde=n_nde,
                nde_long=nde_pct, de_k=de_k, pending=n_pending, figparen=n_figparen)


def verdict(s, cfg):
    """返回违规条目清单；空清单即通过。

    严禁项（计 0 即判）：非并列类未豁免斜杠、隐形字符、英文摘要中文标点、C6 图表括注。
    限量项（超人写语料 p99 才判）：其他一字线族、破折号、概念串联、并列斜杠、双引号、
    书名号、正文冒号、清单冒号、泛权威、分号、同段并列、跨段并列、C4 节号回指、C5 无「的」长句占比
    （C4/C5 在旧基线缺该指标时不判）。
    """
    bad = []

    def over(label, value, cap, unit='/千字'):
        if round(value, 2) > cap:
            bad.append('%s %.2f > %.2f%s' % (label, value, cap, unit))

    over('其他一字线族密度', s['dash_other_k'], cfg.dash_other_k)
    if s['slash_other']:
        bad.append('非常规斜杠 %d 处' % s['slash_other'])
    if s['invisible']:
        bad.append('隐形字符 %d 处' % s['invisible'])
    if s['en_bad']:
        bad.append('英文摘要标点 %d 处' % s['en_bad'])
    if s['figparen']:
        bad.append('图表括注 %d 处（写成「如图3所示」「表1列出」，不挂括号）' % s['figparen'])
    over('破折号密度', s['dash_break_k'], cfg.dash_k)
    over('概念串联密度', s['dash_concept_k'], cfg.dash_concept_k)
    over('并列斜杠密度', s['slash_pair_k'], cfg.slash_k)
    over('双引号密度', s['quote_k'], cfg.quote_k)
    over('书名号密度', s['book_k'], cfg.book_k)
    if s['num_id']:
        bad.append('代号/年份分节 %d 处' % s['num_id'])
    if cfg.num_group == 'plain' and s['num_grouped']:
        bad.append('数值分节 %d 处（口径：一律写 1200，不写 1 200）' % s['num_grouped'])
    elif cfg.num_group == 'grouped' and s['num_plain']:
        bad.append('数值未分节 %d 处（口径：一律写 1 200）' % s['num_plain'])
    elif s['num_mixed']:
        bad.append('多位数分节体例混用')
    if s['num_sep_mixed']:
        bad.append('千分空分隔符混用')
    if cfg.unit_space == 'space' and s['unit_tight']:
        bad.append('数值与单位贴合 %d 处（口径：一律写 80 m/min）' % s['unit_tight'])
    elif cfg.unit_space == 'tight' and s['unit_sp']:
        bad.append('数值与单位留空 %d 处（口径：一律写 80m/min）' % s['unit_sp'])
    elif s['unit_mixed']:
        bad.append('数值与单位间隔体例混用')
    if cfg.qty_lead == 'tight' and s['lead_sp']:
        bad.append('汉字与数值之间留空 %d 处（口径：一律写 下游6～8 m、较2000年）' % s['lead_sp'])
    elif cfg.qty_lead == 'space' and s['lead_tight']:
        bad.append('汉字与量值之间贴合 %d 处（口径：一律写 下游 6～8 m）' % s['lead_tight'])
    elif s['lead_mixed']:
        bad.append('汉字与量值的前置间隔体例混用')
    if s['tilde_half']:
        bad.append('半角波浪号 ~ 用作范围 %d 处（国标形式为全角 ～）' % s['tilde_half'])
    if s['pct_sp']:
        bad.append('百分号前留空 %d 处（百分号紧跟数字，写 122%%）' % s['pct_sp'])
    over('正文冒号密度', s['colon_body_k'], cfg.colon_k)
    over('清单冒号密度', s['colon_list_k'], cfg.colon_list_k)
    over('无引用泛权威密度', s['authority_k'], cfg.auth_k)
    if s['fam'] > cfg.fam_para_max:
        bad.append('同段并列 %d 组 > %d 组' % (s['fam'], cfg.fam_para_max))
    if s['fam_span'] > cfg.fam_span:
        bad.append('跨段并列 %d 组 > %d 组' % (s['fam_span'], cfg.fam_span))
    for cat, n, d, cap in s['lex_cat_bad']:
        bad.append('用词类别「%s」%d 处 %.2f/千字 > 人写 p99 %.2f' % (cat, n, d, cap))
    # 句首重复、段首零回指、挂尾、名词化、顿号链、重合度、过净提示，以及 V 级单词与同段堆叠、
    # T 级、C1～C3、P 级均为告警，不阻断门禁
    if cfg.secref_k is not None:
        over('节号回指密度', s['secref_k'], cfg.secref_k)
    if cfg.nde_long is not None and round(s['nde_long'], 1) > cfg.nde_long:
        bad.append('压缩句占比 %.1f%% > %.1f%%（≥%d 字句中不含「的」）'
                   % (s['nde_long'], cfg.nde_long, NDE_LEN))
    if s['semi_max'] > cfg.semi_para:
        bad.append('单段分号峰值 %d > %d' % (s['semi_max'], cfg.semi_para))
    over('分号密度', s['semi_k'], cfg.semi_k)
    return bad


def _csv_set(value):
    return {x.strip() for x in (value or '').split(',') if x.strip()}


def _norm_fam(name):
    """族名归一：忽略空格与全半角括号差异，便于命令行免引号书写。"""
    out = name.replace(' ', '').replace('\u3000', '')
    return out.replace('(', '（').replace(')', '）')


_FAM_BY_NORM = {_norm_fam(n): n for n in FAMILY_NAMES}


def build_parser():
    p = argparse.ArgumentParser(
        prog='style_lint.py', add_help=True,
        description='中文论文文风门禁：阈值取自人写语料基线（默认 165 篇中文核心期刊全文）。')
    p.add_argument('src', nargs='+', help='源文件（Markdown、纯文本或 .docx）')
    p.add_argument('--report', default='_stylelint.txt', help='报告输出路径')
    p.add_argument('--baseline', default=DEFAULT_BASELINE,
                   help='人写语料基线 json（corpus/build_baseline.py 生成）；'
                        '下列未显式给出的门禁一律取其 p99')
    p.add_argument('--lexicon', default=DEFAULT_LEXICON, help='V 级用词词表（整体替换）')
    p.add_argument('--terms', action='append', default=[],
                   help='追加的术语变体表（格式同 scripts/terms.tsv），可重复给出')
    p.add_argument('--abbr-known', default='',
                   help='免定义的缩写，逗号分隔（追加到内置标准代号表）')
    p.add_argument('--allow-dash', default='标准编号,数学负号,数值区间',
                   help='豁免的一字线 kind，逗号分隔；可选 破折号/标准编号/数值区间/'
                        '数学负号/概念串联/其他一字线族')
    p.add_argument('--allow-slash', default='标准代号,物理单位,分数比例,英文串联',
                   help='豁免的斜杠 kind，逗号分隔；可选 标准代号/物理单位/分数比例/并列替代/英文串联')
    p.add_argument('--allow-fam', default='一是/二是/三是,首先/其次/最后',
                   help='豁免的枚举族名（整体替换默认两族），逗号分隔；族名见 --list-families')
    p.add_argument('--allow-fam-add', default='',
                   help='在默认豁免族之外追加豁免族名，逗号分隔')
    p.add_argument('--fam-para-min', type=int, default=3,
                   help='同段枚举族判违规所需成员数（2 成员仅提示）')
    gate = [('--fam-para-max', int, '同段枚举族成组上限'),
            ('--fam-span', int, '跨段枚举族成组上限'),
            ('--semi-para', int, '单段分号上限'),
            ('--semi-k', float, '全文分号每千字上限'),
            ('--colon-k', float, '正文冒号每千字上限'),
            ('--colon-list-k', float, '清单冒号每千字上限'),
            ('--quote-k', float, '双引号每千字上限'),
            ('--dash-k', float, '破折号每千字上限'),
            ('--dash-concept-k', float, '概念串联一字线每千字上限'),
            ('--dash-other-k', float, '其他一字线族每千字上限'),
            ('--slash-k', float, '并列替代斜杠每千字上限'),
            ('--book-k', float, '书名号每千字上限（法规与标准名、孤立书名号不计入）'),
            ('--auth-k', float, '无引用泛权威每千字上限'),
            ('--secref-k', float, 'C4 节号回指每千字上限'),
            ('--nde-long', float, 'C5 ≥50 字句中不含「的」的占比上限（%%）')]
    for flag, typ, desc in gate:
        p.add_argument(flag, type=typ, default=None, help=desc + '（默认取基线 p99）')
    p.add_argument('--num-group', choices=('plain', 'grouped', 'either'),
                   default='plain',
                   help='四位以上数值的分节口径：plain 一律不分节（默认，1200）；'
                        'grouped 一律分节（1 200）；either 两种都可但不得混用')
    p.add_argument('--unit-space', choices=('space', 'tight', 'either'),
                   default='space',
                   help='数值与单位的间隔口径：space 一律留空（默认，80 m/min）；'
                        'tight 一律贴合；either 两种皆可但不得混用')
    p.add_argument('--qty-lead', choices=('tight', 'space', 'either'),
                   default='tight',
                   help='汉字与量值的间隔口径：tight 一律贴合（默认，下游6～8 m）；'
                        'space 一律留空；either 两种皆可但不得混用')
    p.add_argument('--opener-min', type=float, default=None,
                   help='段首实体立论率下限（默认取基线 p10，告警不进门禁）')
    p.add_argument('--long-sent', type=int, default=80,
                   help='中文长句提示阈值（字符，不进门禁）')
    p.add_argument('--cv-min', type=float, default=0.25,
                   help='段内句长变异系数下限，低于此值提示机械对称拆句（不进门禁）')
    p.add_argument('--dun-max', type=int, default=6,
                   help='顿号链成员数告警阈值')
    p.add_argument('--overlap', type=float, default=0.70,
                   help='相邻句字二元组 Jaccard 告警阈值')
    p.add_argument('--en-words', type=int, default=35,
                   help='英文摘要单句词数提示阈值（不进门禁）')
    p.add_argument('--known-terms', default='',
                   help='目标读者已熟悉、不必解释的词，逗号分隔（同时从 A13 与 A14 词表剔除，'
                        '如通信稿 --known-terms=时延）')
    p.add_argument('--terms-file', default='',
                   help='追加 A14 术语的文本文件，每行一个（UTF-8）')
    p.add_argument('--list-families', action='store_true',
                   help='打印枚举族名后退出')
    return p


class Config(object):
    """命令行参数 + 基线 + 词表。基线、词表读不到时抛 IOError/ValueError，由 main 退 2。"""

    def __init__(self, args):
        self.base = load_baseline(args.baseline)
        metrics = self.base['metrics']
        for attr, key in GATE_P99:
            value = getattr(args, attr)
            if value is None:
                value = metrics[key]['p99']
                if attr in GATE_INT:
                    value = int(round(value))
            setattr(self, attr, value)
        # C4/C5 门禁：旧基线没有这两项时不判，报告注明「基线未含」
        for attr, key in GATE_OPT:
            value = getattr(args, attr)
            if value is None and key in metrics:
                value = metrics[key]['p99']
            setattr(self, attr, value)
        self.opener_min = (args.opener_min if args.opener_min is not None
                           else metrics['opener']['p10'])
        self.lexicon_path = args.lexicon
        self.lexicon = load_lexicon(args.lexicon)
        self.terms = load_terms([DEFAULT_TERMS] + args.terms)
        self.abbr_known = ABBR_KNOWN | _csv_set(args.abbr_known)
        self.allow_dash = _csv_set(args.allow_dash)
        self.allow_slash = _csv_set(args.allow_slash)
        self.allow_fam = {_FAM_BY_NORM.get(_norm_fam(x), x)
                          for x in (_csv_set(args.allow_fam)
                                    | _csv_set(args.allow_fam_add))}
        self.fam_para_min = args.fam_para_min
        self.num_group = args.num_group
        self.unit_space = args.unit_space
        self.qty_lead = args.qty_lead
        self.long_sent = args.long_sent
        self.cv_min = args.cv_min
        self.dun_max = args.dun_max
        self.overlap = args.overlap
        self.en_words = args.en_words
        extra = []
        if args.terms_file:
            with io.open(args.terms_file, encoding='utf-8') as fh:
                extra = [x.strip() for x in fh if x.strip()]
        known = _csv_set(args.known_terms)
        self.term_re = build_term_re(known, extra)
        self.jargon_re = _jargon_pattern(set(JARGON) - known)
        self.report = args.report


def lint_file(path, fh, cfg):
    """跑一份稿件：写报告到 fh，返回汇总字典（corpus/build_baseline.py 复用）。"""
    paras, en_lines, notice, aux = load_paragraphs(path)
    rows = audit(paras, long_sent=cfg.long_sent, lexicon=cfg.lexicon,
                 term_re=cfg.term_re, jargon_re=cfg.jargon_re)
    en_rows = audit_en(en_lines)
    spans = cross_paragraph_families(rows)
    return report(path, rows, en_rows, spans, notice, fh, cfg, aux)


def main(argv):
    parser = build_parser()
    if '--list-families' in argv[1:]:
        for name in FAMILY_NAMES:
            print(name)
        return 0
    args = parser.parse_args(argv[1:])

    missing = [p for p in args.src + [args.baseline, args.lexicon] + args.terms
               if not os.path.exists(p)]
    if missing:
        print('缺失：%s' % '、'.join(missing))
        return 2
    try:
        cfg = Config(args)
    except (IOError, ValueError, re.error) as e:
        print('基线或词表无法读取：%s' % e)
        return 2

    unknown = cfg.allow_fam - set(FAMILY_NAMES)
    if unknown:
        print('未知枚举族：%s（--list-families 查看族名）' % '、'.join(sorted(unknown)))
        return 2

    summaries = []
    with io.open(cfg.report, 'w', encoding='utf-8') as fh:
        for path in args.src:
            summaries.append((path, lint_file(path, fh, cfg)))

    rc = 0
    for path, s in summaries:
        print('%s: 段%d 字%d | 一字线%d(共%d) 斜杠%d(共%d) 引号%d(弯%d) 书名号%d '
              '同段并列%d 跨段并列%d 句首重复%d | 分号%d(峰值%d, %.2f/千字) '
              '冒号%d(违规%d) 重复片段%d | 英文摘要违规%d'
              % (path, s['paras'], s['chars'], s['dash'], s['dash_all'],
                 s['slash'], s['slash_all'], s['quote'], s['curly'],
                 s['book_mark'], s['fam'], s['fam_span'], s['prefix'],
                 s['semi'], s['semi_max'], s['semi_k'], s['colon'],
                 s['colon_bad'], s['repeat'], s['en_bad']))
        print('  A 级：隐形字符%d 泛权威%d | 零回指%d 挂尾%d 名词化%d(%.2f/千字) '
              '弱谓词%.2f/千字 连词%.2f/千字 顿号链%d 相邻重合%d 对称拆句段%d'
              % (s['invisible'], s['authority'], s['anaphora'], s['tail'],
                 s['nominal'], s['nominal_k'], s['weak_k'], s['conj_k'],
                 s['dun_long'], s['overlap'], s['flat']))
        print('  可读性：压缩简称%d 术语首现未解释%d 抽象名词施事%d'
              % (s['jargon'], s['terms'], s['actors']))
        print('  B 级：段首实体%.0f%% 正文冒号%.2f 顿号%.2f 长句%.0f%% 句长CV%.2f '
              '段均%.0f字 | 分布偏离%d 项'
              % (s['opener'], s['colon_body_k'], s['dun_k'], s['long80'],
                 s['sent_cv'], s['para_chars'], len(s['floors'])))
        print('  V/T/C/P：用词%d(超类别%d 超单词%d 同段堆叠%d) 缩写%d 术语变体%d | '
              '的字链%d 递进堆叠%d 等长排比%d 节号回指%d(%.2f/千字) 无的长句%d(%.1f%%) 图表括注%d | '
              '待补引%d'
              % (s['lex'], len(s['lex_cat_bad']), s['lex_word_hi'], s['lex_stack'],
                 s['abbr'], s['term_var'], s['de_chain'], s['stack'], s['parallel'],
                 s['secref'], s['secref_k'], s['nde'], s['nde_long'], s['figparen'],
                 s['pending']))
        bad = verdict(s, cfg)
        if bad:
            rc = 1
            print('  不通过：%s' % '；'.join(bad))
        else:
            print('  通过')
    if cfg.allow_fam:
        print('已豁免枚举族：%s' % '、'.join(sorted(cfg.allow_fam)))
    print('报告：%s' % cfg.report)
    return rc


if __name__ == '__main__':
    sys.exit(main(sys.argv))
