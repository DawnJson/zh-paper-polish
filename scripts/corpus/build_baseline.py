# -*- coding: utf-8 -*-
"""用 style_lint.py 自身的统计口径，从人写语料重建基线 json。

    python build_baseline.py <out.json> <md 目录或 .md 文件>... [--source=说明] [--lexicon=词表]

输入是 corpus_stats.py 产出的 <workdir>/md/*.md（已段落化的正文）。
分两遍：
  1. 逐篇跑 style_lint.lint_file，取与运行时同口径的汇总指标，按篇求分位数；
     词表按篇计每千字密度（未命中记 0），逐词、逐类别求分位数。
  2. 用新基线再逐篇跑一遍，记下每篇的 B 级偏离项数与门禁是否通过，写入 deviation 与 pass_rate。
     通过率不计 F9～F11 书写口径（分节、单位间隔、浪纹线）：那是体例约定而非密度，
     且 PDF 取字会吞掉数值与单位之间的空格；这几项的失败篇数另记在 convention_fails。
stdout 打印分位数表（Markdown），供 references/human-baseline.md 直接替换。
改了 style_lint.py 的判定口径、lexicon.tsv，或换了语料，都要重跑。
"""
import argparse
import glob
import io
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import style_lint as SL  # noqa: E402

QUANTILES = [('p05', 0.05), ('p10', 0.10), ('p50', 0.50), ('p90', 0.90), ('p95', 0.95),
             ('p99', 0.99)]
METRICS = (SL.BAND_METRICS + SL.FLOOR_METRICS
           + ['opener', 'fam', 'fam_span', 'semi_max', 'de_chain_k', 'stack_k', 'parallel_k',
              'secref_k', 'nde_long', 'de_k'])
LABELS = dict(SL.BAND_LABELS, long80='>80 字长句占比 %', sent_cv='句长变异系数',
              para_chars='段均字数', opener='段首实体立论率 %', fam='同段并列组数',
              fam_span='跨段并列组数', semi_max='单段分号峰值', de_chain_k='C1 的字链',
              stack_k='C2 递进堆叠', parallel_k='C3 等长排比', secref_k='C4 节号回指',
              nde_long='C5 ≥50 字句无「的」占比 %', de_k='结构助词「的」')

# verdict() 里属于书写口径的条目前缀
CONVENTION = ('代号/年份分节', '数值分节', '数值未分节', '多位数分节', '千分空', '数值与单位',
              '汉字与量值', '半角波浪号', '图表括注')


def pct(vals, q):
    """最近秩分位数，与 corpus_stats.py 同一取法。"""
    v = sorted(vals)
    return v[min(len(v) - 1, max(0, int(round(q * (len(v) - 1)))))]


def quantiles(vals):
    out = {name: round(pct(vals, q), 2) for name, q in QUANTILES}
    out['max'] = round(max(vals), 2)
    return out


def make_cfg(baseline, lexicon):
    args = SL.build_parser().parse_args(['_', '--baseline=' + baseline, '--lexicon=' + lexicon])
    return SL.Config(args)


def lint_all(paths, cfg):
    sink = io.StringIO()
    return [SL.lint_file(p, sink, cfg) for p in paths]


def collect(paths):
    out = []
    for p in paths:
        if os.path.isdir(p):
            out.extend(sorted(glob.glob(os.path.join(p, '*.md'))))
        else:
            out.append(p)
    return out


def main(argv):
    ap = argparse.ArgumentParser(prog='build_baseline.py')
    ap.add_argument('out')
    ap.add_argument('src', nargs='+')
    ap.add_argument('--lexicon', default=SL.DEFAULT_LEXICON)
    ap.add_argument('--source', default='')
    args = ap.parse_args(argv[1:])

    paths = collect(args.src)
    if not paths:
        print('没有找到 .md 语料')
        return 2

    # 第一遍：指标口径与基线无关，借现有基线构造 Config 即可
    rows = lint_all(paths, make_cfg(SL.DEFAULT_BASELINE, args.lexicon))
    metrics = {key: quantiles([r[key] for r in rows]) for key in METRICS}

    lexicon = SL.load_lexicon(args.lexicon)
    words, cats = {}, {}
    for cat, entry, _, _ in lexicon:
        vals = [r['lex_word_k'].get(entry, 0.0) for r in rows]
        words[entry] = dict(quantiles(vals), df=sum(1 for v in vals if v))
        cats.setdefault(cat, None)
    for cat in cats:
        vals = [r['lex_cat_k'].get(cat, 0.0) for r in rows]
        cats[cat] = dict(quantiles(vals), df=sum(1 for v in vals if v))

    base = dict(source=args.source or '', n=len(rows), metrics=metrics,
                lexicon=dict(words=words, categories=cats))
    with io.open(args.out, 'w', encoding='utf-8') as fh:
        fh.write(json.dumps(base, ensure_ascii=False, indent=1) + '\n')

    # 第二遍：用新基线复跑，得到人写语料自身的偏离项数分布与门禁通过率
    cfg = make_cfg(args.out, args.lexicon)
    rows2 = lint_all(paths, cfg)
    dev, fails, conv = {}, {}, {}
    n_pass = 0
    for r in rows2:
        n = len(r['floors'])
        dev[str(n)] = dev.get(str(n), 0) + 1
        bad = SL.verdict(r, cfg)
        core = [b for b in bad if not b.startswith(CONVENTION)]
        n_pass += not core
        for b in bad:
            key = b.split(' ')[0].split('「')[0]
            tally = fails if b in core else conv
            tally[key] = tally.get(key, 0) + 1
    base['deviation'] = dict(sorted(dev.items(), key=lambda x: int(x[0])))
    base['pass_rate'] = round(100.0 * n_pass / len(rows2), 1)
    base['gate_fails'] = dict(sorted(fails.items(), key=lambda x: -x[1]))
    base['convention_fails'] = dict(sorted(conv.items(), key=lambda x: -x[1]))
    with io.open(args.out, 'w', encoding='utf-8') as fh:
        fh.write(json.dumps(base, ensure_ascii=False, indent=1) + '\n')

    print('语料 %d 篇；门禁通过 %.1f%%（不计书写口径）；偏离项数分布 %s'
          % (len(rows), base['pass_rate'], base['deviation']))
    for title, d in (('密度与严禁项', fails), ('书写口径（不计通过率）', conv)):
        if d:
            print('人写语料的门禁失败项，%s：%s'
                  % (title, '，'.join('%s %d 篇' % kv for kv in
                                     sorted(d.items(), key=lambda x: -x[1]))))
    print('\n| 指标 | 键 | ' + ' | '.join(n for n, _ in QUANTILES) + ' | max |')
    print('|---|---|' + '---|' * (len(QUANTILES) + 1))
    for key in METRICS:
        m = metrics[key]
        print('| %s | %s | %s | %s |' % (LABELS[key], key,
                                         ' | '.join('%.2f' % m[n] for n, _ in QUANTILES),
                                         '%.2f' % m['max']))
    print('\n| 用词类别 | 出现篇数 | p50 | p90 | p95 | p99 | max |')
    print('|---|---|---|---|---|---|---|')
    for cat, m in sorted(cats.items(), key=lambda x: -x[1]['df']):
        print('| %s | %d | %.2f | %.2f | %.2f | %.2f | %.2f |'
              % (cat, m['df'], m['p50'], m['p90'], m['p95'], m['p99'], m['max']))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
