# -*- coding: utf-8 -*-
"""从 PDF 抽取文本中还原中文正文, 按 style_lint 同口径统计人写论文的文风分布。

    python corpus_stats.py <workdir> [min_year]

读 <workdir>/txt/*.txt（自备论文 PDF 抽取的纯文本：首行 `# 题名`，
次行 `# journal=… kind=review|research key=…`，空一行后接正文），输出：
  <workdir>/md/*.md          段落化正文，供 style_lint.py 直接复算分位表
  <workdir>/stats.jsonl      逐篇指标
  <workdir>/stats_summary.json  按 综述/研究型 汇总的分位数
"""
import io, json, os, re, statistics, sys, glob, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.join(os.path.dirname(HERE), 'style_lint.py')
spec = importlib.util.spec_from_file_location("style_lint", SKILL)
SL = importlib.util.module_from_spec(spec)
spec.loader.exec_module(SL)

CJK = re.compile(r'[\u4e00-\u9fff]')
CAPTION = re.compile(r'^\s*(图|表|Fig|Tab|Table|Algorithm|算法)\s?\d')
HEADER = re.compile(r'(学\s*报|Journal of|Vol\.|No\.|卷|期|www\.|http|©|万方|CNKI|下载时间|'
                    r'第\s*\d+\s*卷|doi|DOI|ISSN|E-?mail)')
REFS = re.compile(r'^\s*(参\s*考\s*文\s*献|References|REFERENCES)\s*[:：]?\s*$')
KEYW = re.compile(r'^\s*(关\s*键\s*词|关键字)\s*[:：]')
ABS = re.compile(r'^\s*摘\s*要\s*[:：]?')
NUMLINE = re.compile(r'^[\s\d.,%()\[\]±×/+\-—–~×∼:;<>=]+$')


HALF = [(re.compile(r'(?<=[\u4e00-\u9fff）」』】])\.(?=\s|$|[\u4e00-\u9fff（【])'), '。'),
        (re.compile(r'(?<=[\u4e00-\u9fff）」』】]),'), '，'),
        (re.compile(r'(?<=[\u4e00-\u9fff）」』】]);'), '；'),
        (re.compile(r'(?<=[\u4e00-\u9fff）」』】]):'), '：')]


def normalize_zh(s):
    """部分期刊使用半角句点作中文句号，统一成全角以便同口径切句。"""
    for pat, rep in HALF:
        s = pat.sub(rep, s)
    return s


def cjk_ratio(s):
    if not s:
        return 0.0
    return len(CJK.findall(s)) / float(len(s))


def body_lines(text):
    """截取中文正文行：关键词行之后、参考文献之前；剔除页眉、图表题、公式与英文块。"""
    lines = text.split('\n')
    start, end = 0, len(lines)
    for i, l in enumerate(lines):
        if KEYW.match(l) or ABS.match(l):
            start = i + 1
        if REFS.match(l):
            end = i
            break
    out = []
    for l in lines[start:end]:
        s = l.strip()
        if len(s) < 8 or CAPTION.match(s) or NUMLINE.match(s):
            continue
        if cjk_ratio(s) < 0.45:
            continue
        if HEADER.search(s) and len(s) < 60:
            continue
        s = normalize_zh(re.sub(r'\s+', '', s))
        out.append(s)
    return out


def paragraphs(lines):
    """按“行以句末标点收尾且明显短于常规行”的规则还原段落；过短的片段并入下一段。"""
    if not lines:
        return []
    med = statistics.median(len(l) for l in lines)
    paras, buf = [], []
    for l in lines:
        buf.append(l)
        cur = ''.join(buf)
        if l[-1] in '。！？' and len(l) < med * 0.88 and len(cur) >= 60:
            paras.append(cur)
            buf = []
    if buf:
        paras.append(''.join(buf))
    return [p for p in paras if len(p) >= 40]


def pct(vals, q):
    if not vals:
        return None
    v = sorted(vals)
    i = min(len(v) - 1, max(0, int(round(q * (len(v) - 1)))))
    return v[i]


def measure(path):
    raw = io.open(path, encoding='utf-8').read()
    head, _, rest = raw.partition('\n\n')
    meta = dict(re.findall(r'(\w+)=([^\s]+)', head))
    # 页眉里最先出现的四位数常是被引文献或卷号，取表头区间内最大的合理年份更稳
    years = [int(x) for x in re.findall(r'(?:19[89]|20[0-2])\d', raw[:3000])]
    year = max(years) if years else None
    lines = body_lines(rest)
    paras = paragraphs(lines)
    body = ''.join(paras)
    if len(body) < 3000:
        return None
    sents = SL.sentences(body)
    sents = [s for s in sents if len(s) >= 6]
    if len(sents) < 40:
        return None
    k = len(body) / 1000.0
    lens = [len(s) for s in sents]
    mean = statistics.mean(lens)
    cv = statistics.pstdev(lens) / mean if mean else 0
    ov = [j for j, _, _ in SL.overlap_pairs(sents)]
    chains = SL.dun_chains(body)
    fam_para2 = sum(1 for p in paras if any(len(v) >= 2 for v in SL.family_hits(p).values()))
    fam_para3 = sum(1 for p in paras if any(len(v) >= 3 for v in SL.family_hits(p).values()))
    fam_seq = 0
    prev = set()
    for p in paras:
        cur = {n for n, pat in SL.FAMILIES if pat.search(p)}
        if cur & prev:
            fam_seq += 1
        prev = cur
    dashes = SL.dash_hits(body)
    slashes = SL.slash_hits(body)
    colons = SL.colon_hits(body)
    return dict(
        _paras=paras,
        file=os.path.basename(path), journal=meta.get('journal'), kind=meta.get('kind'),
        year=year, chars=len(body), paras=len(paras), sents=len(sents),
        sent_mean=round(mean, 1), sent_med=statistics.median(lens),
        sent_cv=round(cv, 3), sent_p90=pct(lens, 0.9), sent_max=max(lens),
        long80=round(sum(1 for x in lens if x > 80) * 100.0 / len(lens), 2),
        long60=round(sum(1 for x in lens if x > 60) * 100.0 / len(lens), 2),
        short20=round(sum(1 for x in lens if x < 20) * 100.0 / len(lens), 2),
        para_chars=round(statistics.mean([len(p) for p in paras]), 1),
        para_sents=round(len(sents) / float(len(paras)), 2),
        semi_k=round(body.count('；') / k, 2), colon_k=round(len(colons) / k, 2),
        colon_list=round(sum(1 for c in colons if c[0] == '清单冒号') / k, 2),
        dun_k=round(body.count('、') / k, 2),
        dun_max=max(chains) if chains else 0,
        dun6=sum(1 for c in chains if c >= 6),
        quote_k=round(len(SL.quote_spans(body)) / k, 2),
        dash_k=round(len(dashes) / k, 2),
        dash_break=sum(1 for d in dashes if d[0] == '破折号'),
        slash_k=round(len(slashes) / k, 2),
        slash_concept=sum(1 for s in slashes if s[0] == '并列替代'),
        paren_k=round(body.count('（') / k, 2),
        conj_k=round(SL.conj_count(body) / k, 2),
        nominal_k=round(len(SL.nominal_hits(body)) / k, 2),
        weak_k=round(SL.weak_pred_count(body) / k, 2),
        overlap_mean=round(statistics.mean(ov), 3) if ov else None,
        overlap_p90=round(pct(ov, 0.9), 3) if ov else None,
        overlap_hi=round(sum(1 for x in ov if x >= 0.7) * 100.0 / len(ov), 2) if ov else None,
        tail=len(SL.tail_hits(sents)),
        authority=len(SL.authority_hits(body)),
        anaphora=round(sum(1 for p in paras if SL.ANAPHORA_HEAD.match(p)) * 100.0 / len(paras), 1),
        fam_para2=round(fam_para2 * 100.0 / len(paras), 1),
        fam_para3=round(fam_para3 * 100.0 / len(paras), 1),
        fam_seq=fam_seq,
        invisible=len(SL.invisible_hits(body)),
        jargon_k=round(len(SL.jargon_hits(body)) / k, 2),
        terms=len(SL.term_first_hits(body, SL.build_term_re(), set())),
        actor_k=round(len(SL.abstract_actor_hits(body)) / k, 2),
    )


NUMERIC = ['chars', 'paras', 'sents', 'sent_mean', 'sent_med', 'sent_cv', 'sent_p90',
           'long80', 'long60', 'short20', 'para_chars', 'para_sents', 'semi_k', 'colon_k',
           'colon_list', 'dun_k', 'dun_max', 'dun6', 'quote_k', 'dash_k', 'dash_break',
           'slash_k', 'slash_concept', 'paren_k', 'conj_k', 'nominal_k', 'weak_k',
           'overlap_mean', 'overlap_p90', 'overlap_hi', 'tail', 'authority', 'anaphora',
           'fam_para2', 'fam_para3', 'fam_seq', 'invisible', 'jargon_k', 'terms', 'actor_k']


def summarize(rows, label):
    out = {'label': label, 'n': len(rows)}
    for key in NUMERIC:
        vals = [r[key] for r in rows if r.get(key) is not None]
        if not vals:
            continue
        out[key] = dict(med=round(statistics.median(vals), 2),
                        p10=round(pct(vals, 0.1), 2), p90=round(pct(vals, 0.9), 2),
                        mean=round(statistics.mean(vals), 2))
    return out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    work = sys.argv[1]
    min_year = int(sys.argv[2]) if len(sys.argv) > 2 else 2018
    md_dir = os.path.join(work, 'md')
    os.makedirs(md_dir, exist_ok=True)
    rows = []
    for p in sorted(glob.glob(os.path.join(work, 'txt', '*.txt'))):
        try:
            r = measure(p)
        except Exception as e:
            print('ERR', os.path.basename(p), type(e).__name__, e)
            continue
        if r and (r['year'] is None or r['year'] >= min_year):
            stem = os.path.splitext(os.path.basename(p))[0]
            with io.open(os.path.join(md_dir, stem + '.md'), 'w',
                         encoding='utf-8') as fh:
                fh.write('\n\n'.join(r.pop('_paras')) + '\n')
            rows.append(r)
    with io.open(os.path.join(work, 'stats.jsonl'), 'w', encoding='utf-8') as fh:
        for r in rows:
            fh.write(json.dumps(r, ensure_ascii=False) + '\n')
    rev = [r for r in rows if r['kind'] == 'review']
    res = [r for r in rows if r['kind'] == 'research']
    summary = dict(all=summarize(rows, 'all'),
                   review=summarize(rev, 'review'),
                   research=summarize(res, 'research'))
    json.dump(summary, io.open(os.path.join(work, 'stats_summary.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    print('papers=%d review=%d research=%d (year>=%d)' % (len(rows), len(rev), len(res), min_year))
    for key in ['sent_mean', 'sent_cv', 'long80', 'para_chars', 'para_sents', 'semi_k',
                'colon_k', 'dun_k', 'conj_k', 'nominal_k', 'weak_k', 'overlap_p90',
                'anaphora', 'fam_para2', 'fam_para3', 'dash_k', 'quote_k', 'paren_k',
                'jargon_k', 'terms', 'actor_k']:
        a = summary['review'].get(key, {})
        b = summary['research'].get(key, {})
        print('%-12s 综述 med=%-8s p90=%-8s | 研究 med=%-8s p90=%-8s'
              % (key, a.get('med'), a.get('p90'), b.get('med'), b.get('p90')))
    return 0


if __name__ == '__main__':
    sys.exit(main())
