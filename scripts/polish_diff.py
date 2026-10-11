# -*- coding: utf-8 -*-
"""润色前后稿件的机械对照核验：把手工逐项核对的护栏变成可复算的检查。

只比对、不改稿，也不判文风（文风归 style_lint.py）。两个输入须同为 .md 或同为 .docx。

段落切分与对齐：
  docx  正文层段落（含空段）按序号 n 对齐，表格单元格按 loc（'表3 r2c3'）对齐；
        读取走 docx_text.read_docx
  md    每个非空物理行算一段（一行一段、段间空行），按出现次序对齐，位置记行号 L12
  对齐一律先用 difflib.SequenceMatcher 找逐字相同的锚段，锚段之间按字二元组相似度做
  全局对齐：无增删时即逐段按序号配对；有增删（含段数相同但增删相抵）时不硬按下标配对

受保护段（须逐字不变）：首个非空段（题名）、编号章节 ^\\d+(\\.\\d+)*\\s、Markdown 标题
^#{1,6}\\s、图题表题（图/表/Fig./Figure/Table 加编号，编号后须跟空白、冒号或行尾，
防止「表2汇总了…」这类正文句首被误判）、元数据行（关键词/中图分类号/Keywords 开头）、
参考文献区（参考文献 或 References 行及其后全部）、md 占位符行 {{TABLEn}} 与插图行 ![](...)。
摘要与 Abstract 段正文可改，只保护「摘 要：」「**Abstract:**」这类标签前缀。其余为可改正文。

硬检查（任一失败退出码 1）：
  H1 段数一致（docx 正文层段落总数；md 非空行数），不一致时报首个错位位置
  H2 引用序列一致：docx 取全部段落与单元格的 cites，文字与是否上标都比；md 取 {{key}}
     （全大写占位符除外）与 [n]、[n-m]、[a,b]。报首个分歧点及前后引用。
     [待补引] 单列新增与删除，不算失败
  H3 受保护段逐字不变
  H4 md 占位符与插图标记序列一致；docx 每段图形对象数（w:drawing/w:object/w:pict）一致
  H5 正文字数增幅 ≤ --max-growth 百分比（字数 = 可改正文段剔除引用标记、{{...}}、
     [待补引] 与空白后的字符数；docx 不含表格）
  H6 给了 --pages 且为 docx 时，用 Word 取前后页数，后者多于前者即失败；
     pywin32 缺失或 Word 打开失败时写明跳过原因，不算失败

告警（只列出，不影响退出码）：数值变化（逐段比数值多重集，丢失在前）、单段字数增减超过
--para-delta 百分比、docx 表格单元格增减与文字改动、改动段统计。

用法：
  python polish_diff.py 润色前.docx 润色后.docx
  python polish_diff.py --report=_polishdiff.txt --max-growth=2 --pages 前.docx 后.docx
  python polish_diff.py 前.md 后.md --para-delta=20

通过退出码 0；硬检查失败退出码 1；文件缺失、类型不一致或读取失败退出码 2。
"""
import argparse
import difflib
import io
import os
import re
import shutil
import sys
import tempfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_text import CITE_TXT  # noqa: E402

PENDING = '[待补引]'
BRACE = re.compile(r'\{\{([^{}]*)\}\}')
PLACEHOLDER_KEY = re.compile(r'^[A-Z][A-Z0-9_]*$')
MD_IMAGE = re.compile(r'!\[[^\]]*\]\([^)]*\)')
MD_CITE = re.compile(r'\{\{([^{}]*)\}\}|' + CITE_TXT.pattern)
MD_MARK = re.compile(r'\{\{[A-Z][A-Z0-9_]*\}\}|' + MD_IMAGE.pattern)
NUM = re.compile(r'\d+(?:\.\d+)?')
WS = re.compile(r'\s+')
MATCH_MIN = 0.2   # 锚间配对的相似度门槛；实测稿件跨段 p99 0.32，被改段最低 0.48

HEAD_NUM = re.compile(r'^\d+(\.\d+)*\s')
HEAD_MD = re.compile(r'^#{1,6}\s')
CAPTION = re.compile(r'^(?:(?:图|表)\s?\d+|(?:Fig\.?|Figure|Table)\s?\d+)'
                     r'(?:[-.]\d+)?(?:\s|[:：.]|$)')
META = re.compile(r'^[*\s]*(关键词|中图分类号|Key\s?words)', re.I)
ABS_LABEL = re.compile(r'^[*\s]*(?:摘\s*要|Abstract)[*\s]*[:：]?[*\s]*', re.I)
REFS = re.compile(r'^[#*\s]*(参考文献|References)[*\s:：]*$', re.I)
PH_LINE = re.compile(r'^\s*\{\{[A-Z][A-Z0-9_]*\}\}\s*$')
IMG_LINE = re.compile(r'^\s*!\[')


# ---------------- 读取与分类 ----------------

def classify(units, md):
    """就地给每段写 prot：受保护原因，空串表示可改正文；摘要段另写 label（须不变的标签前缀）。"""
    in_refs = False
    first = True
    for u in units:
        t = u['text']
        s = t.strip()
        reason = ''
        u['label'] = None
        if in_refs or REFS.match(s):
            in_refs = True
            reason = '参考文献区'
        elif first and s:
            reason = '题名'
        elif HEAD_MD.match(s) or HEAD_NUM.match(s):
            reason = '章节标题'
        elif CAPTION.match(s):
            reason = '图表题'
        elif META.match(s):
            reason = '元数据'
        elif md and PH_LINE.match(s):
            reason = '占位符'
        elif md and IMG_LINE.match(s):
            reason = '插图'
        elif s and ABS_LABEL.match(t):
            u['label'] = ABS_LABEL.match(t).group(0)
        if s:
            first = False
        u['prot'] = reason


def load_md(path):
    with io.open(path, encoding='utf-8-sig') as fh:
        lines = fh.read().split('\n')
    paras, cites, marks = [], [], []
    for no, line in enumerate(lines, 1):
        line = line.rstrip('\r')
        if not line.strip():
            continue
        loc = 'L%d' % no
        paras.append(dict(loc=loc, text=line, draw=None))
        for m in MD_CITE.finditer(line):
            if m.group(1) is not None and PLACEHOLDER_KEY.match(m.group(1).strip()):
                continue
            cites.append((m.group(0), None, loc))
        for m in MD_MARK.finditer(line):
            marks.append((m.group(0), loc))
    classify(paras, md=True)
    return dict(paras=paras, cells=[], cites=cites, marks=marks, draw_total=None)


def load_docx(path):
    from docx import Document
    from docx_text import read_docx
    blocks = read_docx(path)
    paras = [dict(loc=b['loc'], text=b['text'], draw=0)
             for b in blocks if b['kind'] == 'para']
    cells = []
    seen = set()
    for b in blocks:
        if b['kind'] != 'cell':
            continue
        loc = b['loc']
        k = 2
        while loc in seen:
            loc = '%s#%d' % (b['loc'], k)
            k += 1
        seen.add(loc)
        cells.append(dict(loc=loc, text=b['text']))
    cites = [(c[0], c[1], b['loc']) for b in blocks for c in b['cites']]
    doc = Document(path)
    xp = './/w:drawing|.//w:object|.//w:pict'
    for u, p in zip(paras, doc.paragraphs):
        u['draw'] = len(p._p.xpath(xp))
    draw_total = len(doc.element.body.xpath(xp))
    classify(paras, md=False)
    return dict(paras=paras, cells=cells, cites=cites, marks=[],
                draw_total=draw_total)


# ---------------- 对齐 ----------------

def _bigrams(s):
    s = WS.sub('', s)
    return set(s[i:i + 2] for i in range(len(s) - 1)) or {s}


def _sim(a, b):
    if a == b:
        return 1.0
    inter = len(a & b)
    return inter / float(len(a | b)) if inter else 0.0


def _nw(ia, ib, ga, gb):
    """锚段之间的全局对齐：相似度减 MATCH_MIN 作配对得分，空位 0 分。"""
    m, n = len(ia), len(ib)
    if not m or not n:
        return [(i, None) for i in ia] + [(None, j) for j in ib]
    score = [[0.0] * (n + 1) for _ in range(m + 1)]
    for x in range(1, m + 1):
        row, prev = score[x], score[x - 1]
        for y in range(1, n + 1):
            row[y] = max(prev[y], row[y - 1],
                         prev[y - 1] + _sim(ga[ia[x - 1]], gb[ib[y - 1]]) - MATCH_MIN)
    out = []
    x, y = m, n
    while x or y:
        if x and y and score[x][y] == score[x - 1][y - 1] + \
                _sim(ga[ia[x - 1]], gb[ib[y - 1]]) - MATCH_MIN:
            out.append((ia[x - 1], ib[y - 1]))
            x, y = x - 1, y - 1
        elif x and score[x][y] == score[x - 1][y]:
            out.append((ia[x - 1], None))
            x -= 1
        else:
            out.append((None, ib[y - 1]))
            y -= 1
    return out[::-1]


def align(pa, pb):
    """段数相同且无增删时结果即逐段按序号配对；有增删时不会让后续段整体错位。"""
    ta = [u['text'] for u in pa]
    tb = [u['text'] for u in pb]
    ga = [_bigrams(t) for t in ta]
    gb = [_bigrams(t) for t in tb]
    pairs = []
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            pairs.extend(zip(range(i1, i2), range(j1, j2)))
        else:
            pairs.extend(_nw(list(range(i1, i2)), list(range(j1, j2)), ga, gb))
    return pairs


# ---------------- 工具 ----------------

def strip_marks(text):
    text = BRACE.sub('', text)
    text = CITE_TXT.sub('', text)
    text = MD_IMAGE.sub('', text)
    return text.replace(PENDING, '')


def body_chars(text):
    return len(WS.sub('', strip_marks(text)))


def clip(text, n=60):
    text = text.replace('\n', '⏎')
    return text if len(text) <= n else text[:n] + '…'


def ploc(a, b):
    la = a['loc'] if a else '（无）'
    lb = b['loc'] if b else '（无）'
    return la if la == lb else '%s→%s' % (la, lb)


def num_ctx(text, num, cap=3):
    """数值在该段中的出现上下文，最多 cap 处（丢失的是哪一处无法字面确定，全列）。"""
    ctx = []
    for m in re.finditer(r'(?<![\d.])' + re.escape(num) + r'(?![\d]|\.\d)', text):
        ctx.append('「%s」' % text[max(0, m.start() - 8):m.end() + 8].replace('\n', ' '))
        if len(ctx) == cap:
            break
    return ''.join(ctx)


def num_list(cnt, text):
    keys = sorted(cnt, key=text.find)
    return '、'.join('%s%s %s' % (k, '×%d' % cnt[k] if cnt[k] > 1 else '', num_ctx(text, k))
                    for k in keys)


def cite_str(c):
    if c[1] is None:
        return c[0]
    return c[0] + ('' if c[1] else '(非上标)')


# ---------------- 核验 ----------------

def word_pages(paths):
    """返回 ([页数, ...], None) 或 (None, 跳过原因)。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        return None, 'pywin32 不可用（%s）' % e
    tmp = tempfile.mkdtemp(prefix='polishdiff_')
    pages = []
    word = None
    pythoncom.CoInitialize()
    try:
        copies = []
        for k, p in enumerate(paths):
            cp = os.path.join(tmp, '%d_%s' % (k, os.path.basename(p)))
            shutil.copy(p, cp)
            copies.append(cp)
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        for cp in copies:
            d = word.Documents.Open(cp, ReadOnly=True, AddToRecentFiles=False,
                                    ConfirmConversions=False)
            try:
                d.Repaginate()
                pages.append(int(d.ComputeStatistics(2)))
            finally:
                d.Close(False)
    except Exception as e:
        return None, 'Word 打开或统计失败（%s）' % e
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        shutil.rmtree(tmp, ignore_errors=True)
    return pages, None


def run(args, fmt, A, B):
    out = []
    w = out.append
    fails = []
    S = {}
    pa, pb = A['paras'], B['paras']
    pairs = align(pa, pb)
    unit = '段' if fmt == 'docx' else '非空行'

    w('润色前：%s' % args.before)
    w('润色后：%s' % args.after)
    w('格式：%s' % fmt)
    w('')

    # H1 段数
    w('【一、段数与对齐】')
    S['paras'] = (len(pa), len(pb))
    w('%s数 %d → %d' % (unit, len(pa), len(pb)))
    gaps = [(i, j) for i, j in pairs if i is None or j is None]
    if len(pa) != len(pb):
        fails.append('段数 %d≠%d' % (len(pa), len(pb)))
    if gaps:
        i, j = gaps[0]
        a = pa[i] if i is not None else None
        b = pb[j] if j is not None else None
        w('  首个错位：%s  %s' % (ploc(a, b), '前稿独有段' if b is None else '后稿新增段'))
        w('    文本：%s' % clip((a or b)['text'], 80))
        w('  全部未配对段（SequenceMatcher 取逐字相同段为锚，锚间按字二元组相似度对齐）：')
        for i, j in gaps:
            u = pa[i] if i is not None else pb[j]
            w('    %s %s  %s' % ('删' if j is None else '增', u['loc'], clip(u['text'])))
        if len(pa) == len(pb):
            w('  段数虽同，但有增删相抵；逐段比对按上述对齐进行')
    else:
        w('  逐段一一配对，无增删')
    w('')

    # H2 引用序列
    w('【二、引用序列】')
    ca, cb = A['cites'], B['cites']
    ka = [c[:2] for c in ca]
    kb = [c[:2] for c in cb]
    S['cites'] = (len(ca), len(cb))
    if fmt == 'docx':
        S['sup'] = (sum(1 for c in ca if c[1]), sum(1 for c in cb if c[1]))
        w('引用 %d → %d 处（其中上标 %d → %d）' % (len(ca), len(cb), S['sup'][0], S['sup'][1]))
    else:
        w('引用 %d → %d 处（{{key}} 与 [n] 形式）' % (len(ca), len(cb)))
    if ka == kb:
        S['cite_ok'] = True
        w('  序列逐处一致（文字%s）' % ('与上标状态' if fmt == 'docx' else ''))
    else:
        S['cite_ok'] = False
        k = next((x for x in range(min(len(ka), len(kb))) if ka[x] != kb[x]),
                 min(len(ka), len(kb)))
        fails.append('引用序列在第 %d 处分歧' % (k + 1))
        la = ca[k][2] if k < len(ca) else '（前稿已结束）'
        lb = cb[k][2] if k < len(cb) else '（后稿已结束）'
        w('  首个分歧：第 %d 处，前稿 %s，后稿 %s' % (k + 1, la, lb))
        w('    前稿：%s' % ' '.join(
            ('〈%s〉' if x == k else '%s') % cite_str(c)
            for x, c in enumerate(ca[max(0, k - 3):k + 4], max(0, k - 3))))
        w('    后稿：%s' % ' '.join(
            ('〈%s〉' if x == k else '%s') % cite_str(c)
            for x, c in enumerate(cb[max(0, k - 3):k + 4], max(0, k - 3))))
        w('  全部差异（difflib）：')
        sm = difflib.SequenceMatcher(None, ka, kb, autojunk=False)
        shown = 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                continue
            shown += 1
            if shown > 40:
                w('    ……其余从略')
                break
            name = {'replace': '改', 'delete': '删', 'insert': '增'}[tag]
            left = ' '.join(cite_str(c) for c in ca[i1:i2])
            right = ' '.join(cite_str(c) for c in cb[j1:j2])
            where = ploc(dict(loc=ca[i1][2]) if i1 < i2 else None,
                         dict(loc=cb[j1][2]) if j1 < j2 else None)
            w('    %s %s  %s → %s' % (name, where, left or '∅', right or '∅'))
    # 待补引
    pend = []
    for i, j in pairs:
        a = pa[i] if i is not None else None
        b = pb[j] if j is not None else None
        na = a['text'].count(PENDING) if a else 0
        nb = b['text'].count(PENDING) if b else 0
        if na != nb:
            pend.append((ploc(a, b), na, nb, (b or a)['text']))
    cmap_a = dict((c['loc'], c['text']) for c in A['cells'])
    cmap_b = dict((c['loc'], c['text']) for c in B['cells'])
    for loc in list(cmap_a) + [l for l in cmap_b if l not in cmap_a]:
        na = cmap_a.get(loc, '').count(PENDING)
        nb = cmap_b.get(loc, '').count(PENDING)
        if na != nb:
            pend.append((loc, na, nb, cmap_b.get(loc, cmap_a.get(loc, ''))))
    pa_tot = sum(u['text'].count(PENDING) for u in pa) + sum(t.count(PENDING) for t in cmap_a.values())
    pb_tot = sum(u['text'].count(PENDING) for u in pb) + sum(t.count(PENDING) for t in cmap_b.values())
    S['pending'] = (pa_tot, pb_tot)
    w('%s %d → %d 处（不计入失败）' % (PENDING, pa_tot, pb_tot))
    for loc, na, nb, text in pend:
        pos = text.find(PENDING)
        ctx = text[max(0, pos - 20):pos + 10] if pos >= 0 else clip(text, 30)
        w('  %s %s %d 处  %s' % (loc, '新增' if nb > na else '删除', abs(nb - na),
                                 '「%s」' % ctx.replace('\n', ' ')))
    w('')

    # H3 受保护段
    w('【三、受保护段】')
    prot_n = 0
    prot_bad = []
    for i, j in pairs:
        a = pa[i] if i is not None else None
        b = pb[j] if j is not None else None
        reason = (a and a['prot']) or (b and b['prot'])
        if not reason:
            if (a and a['label']) or (b and b['label']):
                prot_n += 1
                la = a['label'] if a else None
                lb = b['label'] if b else None
                if la != lb:
                    prot_bad.append((ploc(a, b), '摘要标签', a, b))
            continue
        prot_n += 1
        if a is None or b is None or a['text'] != b['text']:
            prot_bad.append((ploc(a, b), reason, a, b))
    S['prot'] = (prot_n, len(prot_bad))
    w('受保护段 %d 段（按两稿任一侧判定；摘要段只计标签前缀），改动 %d 段'
      % (prot_n, len(prot_bad)))
    for loc, reason, a, b in prot_bad:
        w('  %s [%s]' % (loc, reason))
        w('    前：%s' % (a['text'] if a else '（无此段）'))
        w('    后：%s' % (b['text'] if b else '（无此段）'))
    if prot_bad:
        fails.append('受保护段改动 %d 段' % len(prot_bad))
    w('')

    # H4 占位符与图形
    w('【四、占位符与图形对象】')
    if fmt == 'md':
        ma = [m[0] for m in A['marks']]
        mb = [m[0] for m in B['marks']]
        S['marks'] = (len(ma), len(mb))
        w('占位符与插图标记 %d → %d 个' % (len(ma), len(mb)))
        if ma != mb:
            fails.append('占位符/插图序列不一致')
            sm = difflib.SequenceMatcher(None, ma, mb, autojunk=False)
            for tag, i1, i2, j1, j2 in sm.get_opcodes():
                if tag == 'equal':
                    continue
                where = A['marks'][i1][1] if i1 < i2 else B['marks'][j1][1]
                w('  %s  %s → %s' % (where, ' '.join(ma[i1:i2]) or '∅',
                                      ' '.join(mb[j1:j2]) or '∅'))
        else:
            w('  序列一致')
    else:
        da = sum(u['draw'] for u in pa)
        db = sum(u['draw'] for u in pb)
        S['marks'] = (da, db)
        w('正文层段落图形对象 %d → %d 个；全文（含表格）%d → %d 个'
          % (da, db, A['draw_total'], B['draw_total']))
        bad = []
        for i, j in pairs:
            a = pa[i] if i is not None else None
            b = pb[j] if j is not None else None
            xa = a['draw'] if a else 0
            xb = b['draw'] if b else 0
            if xa != xb:
                bad.append('  %s  %d → %d' % (ploc(a, b), xa, xb))
        if A['draw_total'] != B['draw_total']:
            bad.append('  全文图形对象总数 %d → %d' % (A['draw_total'], B['draw_total']))
        if bad:
            fails.append('图形对象数不一致')
            out.extend(bad)
        else:
            w('  逐段一致')
    w('')

    # H5 篇幅
    w('【五、正文篇幅】')
    ta = sum(body_chars(u['text']) for u in pa if not u['prot'])
    tb = sum(body_chars(u['text']) for u in pb if not u['prot'])
    growth = (tb - ta) * 100.0 / ta if ta else (0.0 if not tb else float('inf'))
    S['chars'] = (ta, tb, growth)
    w('可改正文 %d → %d 字，增幅 %+.2f%%（上限 %+.2f%%）' % (ta, tb, growth, args.max_growth))
    if growth > args.max_growth:
        fails.append('正文增幅 %+.2f%% 超上限 %.2f%%' % (growth, args.max_growth))
    w('')

    # H6 页数
    w('【六、页数】')
    S['pages'] = None
    if not args.pages:
        S['pages_note'] = '未检（未给 --pages）'
    elif fmt != 'docx':
        S['pages_note'] = '跳过：页数仅对 docx 适用'
    else:
        pages, why = word_pages([args.before, args.after])
        if pages is None:
            S['pages_note'] = '跳过：%s' % why
        else:
            S['pages'] = pages
            S['pages_note'] = '%d → %d' % tuple(pages)
            if pages[1] > pages[0]:
                fails.append('页数 %d→%d 增加' % tuple(pages))
    w('页数 %s' % S['pages_note'])
    w('')

    # 告警：数值
    w('【告警·数值变化】（逐段比数值多重集，已剔除引用标记与 {{...}}；丢失在前）')
    lost_rows, add_rows = [], []
    for i, j in pairs:
        a = pa[i] if i is not None else None
        b = pb[j] if j is not None else None
        sa = strip_marks(a['text']) if a else ''
        sb = strip_marks(b['text']) if b else ''
        na, nb = Counter(NUM.findall(sa)), Counter(NUM.findall(sb))
        if na == nb:
            continue
        lost, added = na - nb, nb - na
        row = (ploc(a, b), lost, added, sa, sb)
        (lost_rows if lost else add_rows).append(row)
    for loc, lost, added, sa, sb in lost_rows + add_rows:
        parts = []
        if lost:
            parts.append('丢失 ' + num_list(lost, sa))
        if added:
            parts.append('新增 ' + num_list(added, sb))
        w('  %s  %s' % (loc, ' | '.join(parts)))
    S['num_warn'] = len(lost_rows) + len(add_rows)
    if not S['num_warn']:
        w('  无')
    w('')

    # 告警：单段字数
    w('【告警·单段字数变化】（可改正文，超过 ±%g%%）' % args.para_delta)
    delta_n = 0
    changed = 0
    changed_locs = []
    for i, j in pairs:
        a = pa[i] if i is not None else None
        b = pb[j] if j is not None else None
        if (a and a['prot']) or (b and b['prot']):
            continue
        if a and b and a['text'] == b['text']:
            continue
        changed += 1
        changed_locs.append(ploc(a, b))
        xa = body_chars(a['text']) if a else 0
        xb = body_chars(b['text']) if b else 0
        if xa == 0 and xb == 0:
            continue
        pct = (xb - xa) * 100.0 / xa if xa else float('inf')
        if abs(pct) > args.para_delta:
            delta_n += 1
            w('  %s  %d → %d 字（%s）' % (ploc(a, b), xa, xb,
                                         '%+.1f%%' % pct if xa else '新增段'))
    S['delta_warn'] = delta_n
    if not delta_n:
        w('  无')
    w('')

    # 告警：表格
    cell_warn = 0
    if fmt == 'docx':
        w('【告警·表格】')
        cnt_a = Counter(l.rsplit(' ', 1)[0] for l in cmap_a)
        cnt_b = Counter(l.rsplit(' ', 1)[0] for l in cmap_b)
        for t in list(cnt_a) + [t for t in cnt_b if t not in cnt_a]:
            if cnt_a[t] != cnt_b[t]:
                cell_warn += 1
                w('  表%s 单元格数 %d → %d' % (t[1:], cnt_a[t], cnt_b[t]))
        for loc in list(cmap_a) + [l for l in cmap_b if l not in cmap_a]:
            ta_, tb_ = cmap_a.get(loc), cmap_b.get(loc)
            if ta_ == tb_:
                continue
            cell_warn += 1
            w('  %s' % loc)
            w('    前：%s' % (clip(ta_, 200) if ta_ is not None else '（无此格）'))
            w('    后：%s' % (clip(tb_, 200) if tb_ is not None else '（无此格）'))
        if not cell_warn:
            w('  无')
        w('')
    S['cell_warn'] = cell_warn

    w('【告警·改动段】')
    S['changed'] = changed
    w('可改正文改动（含增删）%d 段' % changed)
    for k in range(0, len(changed_locs), 12):
        w('  ' + ' '.join(changed_locs[k:k + 12]))
    w('')

    w('【结论】')
    w('通过' if not fails else '不通过：' + '；'.join(fails))
    return out, S, fails


def summary(fmt, S, fails, report):
    lines = []
    p = S['paras']
    c = S['cites']
    cite = '引用 %d%s%d' % (c[0], '=' if c[0] == c[1] else '≠', c[1])
    if fmt == 'docx':
        cite += '（上标 %d/%d）' % S['sup']
    cite += ' 序列%s' % ('一致' if S['cite_ok'] else '分歧')
    lines.append('段数 %d%s%d | %s | 待补引 %d→%d | %s %d→%d'
                 % (p[0], '=' if p[0] == p[1] else '≠', p[1], cite,
                    S['pending'][0], S['pending'][1],
                    '图形对象' if fmt == 'docx' else '占位符/插图',
                    S['marks'][0], S['marks'][1]))
    ch = S['chars']
    lines.append('受保护段 %d 段、改动 %d | 正文 %d→%d 字（%+.2f%%）| 页数 %s'
                 % (S['prot'][0], S['prot'][1], ch[0], ch[1], ch[2], S['pages_note']))
    lines.append('告警：改动段 %d | 数值变化 %d 段 | 单段字数超限 %d 段 | 表格 %d 处'
                 % (S['changed'], S['num_warn'], S['delta_warn'], S['cell_warn']))
    lines.append('结论：' + ('通过' if not fails else '不通过：' + '；'.join(fails)))
    lines.append('报告：%s' % report)
    return lines


def build_parser():
    p = argparse.ArgumentParser(
        prog='polish_diff.py', add_help=True,
        description='润色前后稿件机械对照：引用序列、段数、受保护段、图形、篇幅、页数。')
    p.add_argument('before', help='润色前稿件（.md 或 .docx）')
    p.add_argument('after', help='润色后稿件（与前稿同类型）')
    p.add_argument('--report', default='_polishdiff.txt', help='报告输出路径')
    p.add_argument('--max-growth', type=float, default=1.0,
                   help='正文字数增幅上限（百分比）')
    p.add_argument('--para-delta', type=float, default=15.0,
                   help='单段字数变化告警阈值（百分比）')
    p.add_argument('--pages', action='store_true',
                   help='用 Word 统计前后页数（仅 docx；需 pywin32 与 Word）')
    return p


def main(argv):
    try:
        sys.stdout.reconfigure(errors='backslashreplace')
    except AttributeError:
        pass
    args = build_parser().parse_args(argv[1:])
    missing = [p for p in (args.before, args.after) if not os.path.isfile(p)]
    if missing:
        print('缺失：%s' % '、'.join(missing))
        return 2
    exts = [os.path.splitext(p)[1].lower() for p in (args.before, args.after)]
    if exts[0] != exts[1] or exts[0] not in ('.md', '.docx'):
        print('两个输入须同为 .md 或同为 .docx：%s' % '、'.join(exts))
        return 2
    fmt = exts[0][1:]
    loader = load_docx if fmt == 'docx' else load_md
    try:
        A = loader(args.before)
        B = loader(args.after)
    except Exception as e:
        print('读取失败：%s' % e)
        return 2
    out, S, fails = run(args, fmt, A, B)
    rdir = os.path.dirname(os.path.abspath(args.report))
    if not os.path.isdir(rdir):
        os.makedirs(rdir)
    with io.open(args.report, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(out) + '\n')
    for line in summary(fmt, S, fails, args.report):
        print(line)
    return 1 if fails else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
