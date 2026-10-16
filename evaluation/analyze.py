"""지시문 3판 비교 실험 — 결과 분석 (논문 Ⅴ장).

실험 당시 분석 코드(analyze.py)의 공개본이다. 계산은 바꾸지 않았고, 경로를 인자로 받게 하고
지시문 이름을 공개명(v1·v2·v3)으로 바꾸었다. 코드 안의 '정답지'는 논문의 '기준 판독문'이다.
기준 판독문은 공개하지 않으므로, 이 코드는 기준 판독문을 가진 연구자만 실행할 수 있다.

사용 예:
  python evaluation/analyze.py --sample 표본목록.csv --gold-dir 기준판독문 \
      --results 실험데이터/출력/gemini-3.8-flash --out 실험데이터/분석결과/gemini-3.8-flash

표본목록 열: 파일id, 분석집합 (필수) / 주범주, 구분, 판독문바탕, 확인필요 (하위·민감도 분석용, 없으면 생략)
기준 판독문: <gold-dir>/<파일id>.txt

산출물 (실험데이터/분석결과/):
  요약.txt           주 분석·하위 분석·민감도 분석·목표 어휘·짧은 페이지 결과
  페이지별_지표.csv   페이지 × 버전의 지표(3회 평균)와 실행 간 표준편차
  어휘_집계.csv       상세화 지시문 추가 어휘별 복원·삽입 횟수
"""
import argparse
import csv
import difflib
import re
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
import openpyxl
from Levenshtein import ratio as lev_ratio
from scipy.stats import wilcoxon

HERE = Path(__file__).parent
PROMPT_DIR = HERE.parent / 'prompts'
PROMPT_FILES = {'v1': 'v1_initial.txt', 'v2': 'v2_revised.txt', 'v3': 'v3_detailed.txt'}
SAMPLE = None    # main()에서 지정
GOLD_DIR = None
OUT_DIR = None

VERSIONS = ['v1', 'v2', 'v3']
COMPARISONS = [('v1', 'v2', '비교 1: 사료 지식 반영 (초기 v1 → 수정 v2)'), ('v2', 'v3', '비교 2: 상세화 (수정 v2 → 상세화 v3)')]
METRICS = [('f1', 'F1'), ('precision', '정밀도'), ('recall', '재현율'), ('char_acc', '문자 정확도'), ('lev', 'Levenshtein')]
N_BOOT, SEED = 10000, 20260925

# 상세화 지시문(v3)에서 추가된 어휘 (prompts/diff_v2_v3.diff의 + 행). 수정 지시문(v2)에 이미 있던 어휘는 실행 시 자동 제외한다.
V4B_TERMS = {
    '필드 값': ['進峴里', '凡谷里', '馬洞里', '一番ノ一', '五番', '一五番', '三五番', '四五番', '社寺地',
              '林野', '山林', '雜種地', '三,三〇九坪', '四〇坪', '所有者'],
    '건물명·지명': ['慶尚北道', '金銅阿彌陀如來坐像', '紫霞門', '泛影樓', '大雄殿', '極樂殿', '無說殿',
               '奉恩寺', '祇林寺'],
    '결재 어휘': ['起案', '接受', '發送', '記帳', '鈐印', '校合', '淨寫', '關係', '番號', '主務', '決裁',
              '完結', '未完結', '件名', '施行月日番號', '學務局長', '社會課長', '文書課長', '主任', '總督',
              '政務總監', '博物館長', '朝鮮總督府', '總督府', '學務局', '社會課', '文書課', '人事課', '賞勳局',
              '寶物古蹟名勝天然紀念物保存會', '保存會', '保存會長'],
}


def normalize(text):
    """기존 evaluate_vs_gold.py와 같은 정규화: 공백과 '|' 제거."""
    return re.sub(r'\s+', '', text).replace('|', '')


def score(gold, pred):
    """evaluate_vs_gold.compare_pair와 같은 계산 + Levenshtein 유사도."""
    g, p = normalize(gold), normalize(pred)
    gc, pc = Counter(g), Counter(p)
    tp = sum((gc & pc).values())
    prec = tp / len(p) if p else 0.0
    rec = tp / len(g) if g else 0.0
    f1 = 2 * prec * rec / (prec + rec) if prec + rec else 0.0
    return {'f1': f1, 'precision': prec, 'recall': rec,
            'char_acc': difflib.SequenceMatcher(None, g, p, autojunk=False).ratio(),
            'lev': lev_ratio(g, p)}


def build_term_matcher():
    v3 = (PROMPT_DIR / PROMPT_FILES['v2']).read_text(encoding='utf-8')     # 수정 지시문
    v4b = (PROMPT_DIR / PROMPT_FILES['v3']).read_text(encoding='utf-8')    # 상세화 지시문
    term_group, dropped = {}, []
    for group, terms in V4B_TERMS.items():
        for t in terms:
            if t in v3:
                dropped.append(t)             # v3에도 있던 어휘 → v4b의 효과로 볼 수 없음
            elif t in v4b:
                term_group[normalize(t)] = group
    # 긴 어휘부터 맞춰 겹치는 어휘(保存會 / 保存會長 등)를 한 번만 센다
    pattern = re.compile('|'.join(sorted(map(re.escape, term_group), key=len, reverse=True)))
    return pattern, term_group, dropped


def count_terms(pattern, text):
    return Counter(pattern.findall(normalize(text)))


def load_sample():
    path = Path(SAMPLE)
    if path.suffix.lower() == '.csv':
        rows = list(csv.DictReader(open(path, encoding='utf-8-sig')))
    else:
        ws = openpyxl.load_workbook(path, read_only=True)['표본목록']
        it = ws.iter_rows(values_only=True)
        header = list(next(it))
        rows = [dict(zip(header, r)) for r in it if r and r[0]]
    pages = []
    for r in rows:
        fid = r['파일id']
        pages.append({
            'fid': fid, 'folder': '-'.join(str(fid).split('-')[:2]),
            'category': r.get('주범주') or '-', 'group': r.get('구분') or '-', 'base': r.get('판독문바탕') or '-',
            'main': r['분석집합'] == '본 분석',
            'gold_identical': str(r.get('확인필요') or '').startswith('정답지가 Gemini 결과와 완전히 같음'),
            'gold_path': Path(GOLD_DIR) / f"{fid}.txt",
        })
    return pages


def load_outputs(results_dir, fid):
    """버전별 [run1 텍스트, run2 텍스트, ...] (결과 파일이 있는 것만)."""
    out = {}
    for v in VERSIONS:
        runs = sorted((results_dir / v).glob(f'run*/{fid}.txt'))
        out[v] = [p.read_text(encoding='utf-8') for p in runs]
    return out


def paired_stats(a, b):
    """b − a의 평균, 부트스트랩 95% CI, Wilcoxon p, 짝 순위이연상관 r."""
    d = np.asarray(b) - np.asarray(a)
    rng = np.random.default_rng(SEED)
    boots = rng.choice(d, size=(N_BOOT, len(d)), replace=True).mean(axis=1)
    lo, hi = np.percentile(boots, [2.5, 97.5])
    nz = d[d != 0]
    if len(nz) == 0:
        p, r = 1.0, 0.0
    else:
        p = wilcoxon(nz).pvalue
        ranks = np.argsort(np.argsort(np.abs(nz))) + 1
        r = (ranks[nz > 0].sum() - ranks[nz < 0].sum()) / ranks.sum()
    return {'n': len(d), 'mean': d.mean(), 'lo': lo, 'hi': hi, 'p': p, 'r': r,
            'up': int((d > 0).sum()), 'down': int((d < 0).sum())}


def holm(pvals):
    order = sorted(range(len(pvals)), key=lambda i: pvals[i])
    adj, running = [0.0] * len(pvals), 0.0
    for rank, i in enumerate(order):
        running = max(running, min(1.0, (len(pvals) - rank) * pvals[i]))
        adj[i] = running
    return adj


def verdict(s, p_adj):
    if s['lo'] <= 0 <= s['hi'] or p_adj >= 0.05:
        return '효과를 검출하지 못함'
    return '개선' if s['mean'] > 0 else '악화'


def fmt(s, pct=True):
    k = 100 if pct else 1
    return (f"n={s['n']:3d}  차이 {s['mean'] * k:+6.2f}%p  95% CI [{s['lo'] * k:+6.2f}, {s['hi'] * k:+6.2f}]  "
            f"좋아짐 {s['up']:3d} / 나빠짐 {s['down']:3d}  p={s['p']:.4f}  r={s['r']:+.2f}")


def main():
    ap = argparse.ArgumentParser(description='프롬프트 상세도 비교 실험 분석')
    ap.add_argument('--sample', required=True, help='표본목록 CSV/XLSX')
    ap.add_argument('--gold-dir', required=True, help='기준 판독문 폴더 (<파일id>.txt)')
    ap.add_argument('--results', required=True, help='run_experiment.py 출력 폴더 (<모델> 단위)')
    ap.add_argument('--out', required=True, help='분석 결과 저장 폴더')
    args = ap.parse_args()
    results_dir = Path(args.results)
    global OUT_DIR, SAMPLE, GOLD_DIR
    SAMPLE, GOLD_DIR, OUT_DIR = args.sample, Path(args.gold_dir), Path(args.out)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    pattern, term_group, dropped = build_term_matcher()
    pages = load_sample()

    rows, term_tally, short_rows = [], defaultdict(Counter), []
    incomplete = []
    for pg in pages:
        gold = pg['gold_path'].read_text(encoding='utf-8')
        outs = load_outputs(results_dir, pg['fid'])
        if any(not outs[v] for v in VERSIONS):
            incomplete.append(pg['fid'])
            continue
        gold_terms = count_terms(pattern, gold)
        rec = dict(pg, gold_len=len(normalize(gold)))
        for v in VERSIONS:
            scores = [score(gold, t) for t in outs[v]]
            for key, _ in METRICS:
                vals = [s[key] for s in scores]
                rec[f'{v}_{key}'] = float(np.mean(vals))
                rec[f'{v}_{key}_sd'] = float(np.std(vals, ddof=1)) if len(vals) > 1 else 0.0
            rec[f'{v}_runs'] = len(outs[v])
            rec[f'{v}_len'] = float(np.mean([len(normalize(t)) for t in outs[v]]))
            # 목표 어휘: 실행별로 세어 평균을 낸다
            restored = inserted = 0.0
            for t in outs[v]:
                c = count_terms(pattern, t)
                for term, n in c.items():
                    ok = min(n, gold_terms.get(term, 0))
                    term_tally[(v, term)]['복원'] += ok / len(outs[v])
                    term_tally[(v, term)]['삽입'] += (n - ok) / len(outs[v])
                    restored += ok / len(outs[v])
                    inserted += (n - ok) / len(outs[v])
            rec[f'{v}_term_restored'], rec[f'{v}_term_inserted'] = restored, inserted
        (rows if pg['main'] else short_rows).append(rec)

    lines = []
    say = lines.append
    say('프롬프트 상세도 비교 실험 — 분석 결과')
    say('=' * 100)
    say(f'결과 폴더: {results_dir}')
    say(f'본 분석 {len(rows)}장 / 짧은 페이지 {len(short_rows)}장 / 결과가 빠져 제외된 페이지 {len(incomplete)}장')
    if incomplete:
        say(f'  제외: {incomplete[:20]}{" ..." if len(incomplete) > 20 else ""}')
    say(f'실행 횟수(페이지당 평균): ' + ', '.join(f"{v} {np.mean([r[f'{v}_runs'] for r in rows]):.2f}회" for v in VERSIONS) if rows else '')
    say('')

    if rows:
        # 1. 버전별 평균
        say('[1] 버전별 평균 (페이지마다 반복 실행 결과를 평균한 뒤, 페이지 전체를 평균)')
        say(f"{'지표':<14}" + ''.join(f'{v:>10}' for v in VERSIONS) + '   실행 간 표준편차(페이지 평균)')
        for key, label in METRICS:
            say(f'{label:<14}' + ''.join(f"{np.mean([r[f'{v}_{key}'] for r in rows]) * 100:>9.2f}%" for v in VERSIONS)
                + '   ' + ', '.join(f"{v} {np.mean([r[f'{v}_{key}_sd'] for r in rows]) * 100:.2f}%p" for v in VERSIONS))
        say('')

        # 2. 주 분석
        say('[2] 주 분석 — F1, 페이지별 짝 비교 (Holm 보정은 두 비교의 F1에 적용)')
        main_stats = [paired_stats([r[f'{a}_f1'] for r in rows], [r[f'{b}_f1'] for r in rows]) for a, b, _ in COMPARISONS]
        adj = holm([s['p'] for s in main_stats])
        for (a, b, name), s, pa in zip(COMPARISONS, main_stats, adj):
            say(f'{name}')
            say(f'  {fmt(s)}  Holm 보정 p={pa:.4f}  → 판정: {verdict(s, pa)}')
        say('')

        say('[3] 보조 지표 (보정 없음)')
        for a, b, name in COMPARISONS:
            say(name)
            for key, label in METRICS[1:]:
                say(f'  {label:<12} ' + fmt(paired_stats([r[f'{a}_{key}'] for r in rows], [r[f'{b}_{key}'] for r in rows])))
        say('')

        def sub(title, key_fn):
            say(title)
            groups = defaultdict(list)
            for r in rows:
                groups[key_fn(r)].append(r)
            for g in sorted(groups):
                rs = groups[g]
                if len(rs) < 5:
                    say(f'  {g:<14} n={len(rs)} (5장 미만, 생략)')
                    continue
                for a, b, name in COMPARISONS:
                    s = paired_stats([r[f'{a}_f1'] for r in rs], [r[f'{b}_f1'] for r in rs])
                    say(f'  {g:<14} {a}→{b:<4} ' + fmt(s))
            say('')

        sub('[4] 하위 분석 ① 주 범주별 (F1)', lambda r: r['category'])
        sub('[5] 하위 분석 ② 정답지 바탕별 (F1) — 두 집단에서 같은 방향이면 정답지 편향의 영향이 작다', lambda r: r['base'])

        say('[6] 민감도 분석 — 정답지가 Gemini 결과와 완전히 같은 페이지를 제외 (F1)')
        kept = [r for r in rows if not r['gold_identical']]
        say(f'  제외 {len(rows) - len(kept)}장')
        for a, b, name in COMPARISONS:
            say(f'  {name}: ' + fmt(paired_stats([r[f'{a}_f1'] for r in kept], [r[f'{b}_f1'] for r in kept])))
        say('')

        say('[7] 목표 어휘 (상세화 지시문 v3에서 새로 추가된 어휘, 본 분석 페이지 합계, 실행 평균)')
        say(f'  수정 지시문(v2)에 이미 있어 제외한 어휘 {len(dropped)}개: {", ".join(dropped)}')
        say(f'  집계 대상 어휘 {len(term_group)}개')
        for group in V4B_TERMS:
            say(f'  〈{group}〉')
            for v in VERSIONS:
                rs = sum(term_tally[(v, t)]['복원'] for t, g in term_group.items() if g == group)
                ins = sum(term_tally[(v, t)]['삽입'] for t, g in term_group.items() if g == group)
                say(f'    {v:<4} 복원 {rs:7.1f}회   삽입 {ins:7.1f}회')
        say('  삽입이 있었던 페이지 수(실행 평균 0.5회 이상): ' + ', '.join(
            f"{v} {sum(r[f'{v}_term_inserted'] >= 0.5 for r in rows)}장" for v in VERSIONS))
        say('')

    if short_rows:
        say('[8] 짧은 페이지 (정답 20자 미만) — 원본에 없는 내용의 추가 여부')
        for v in VERSIONS:
            excess = [max(0.0, r[f'{v}_len'] - r['gold_len']) for r in short_rows]
            say(f'  {v:<4} 정답보다 긴 출력 {sum(e > 0 for e in excess):2d}장 / 초과 글자 평균 {np.mean(excess):6.1f}자 / '
                f'목표 어휘 삽입 {sum(r[f"{v}_term_inserted"] for r in short_rows):.1f}회')
        say('')

    text = '\n'.join(lines)
    print(text)
    (OUT_DIR / '요약.txt').write_text(text, encoding='utf-8')

    # 페이지별 CSV
    all_rows = rows + short_rows
    if all_rows:
        cols = ['fid', 'folder', 'category', 'group', 'base', 'main', 'gold_identical', 'gold_len']
        cols += [f'{v}_{k}' for v in VERSIONS for k, _ in METRICS] + [f'{v}_{k}_sd' for v in VERSIONS for k, _ in METRICS]
        cols += [f'{v}_{k}' for v in VERSIONS for k in ('len', 'runs', 'term_restored', 'term_inserted')]
        with open(OUT_DIR / '페이지별_지표.csv', 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore')
            w.writeheader()
            for r in all_rows:
                w.writerow({k: (round(v, 4) if isinstance(v, float) else v) for k, v in r.items()})
        with open(OUT_DIR / '어휘_집계.csv', 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            w.writerow(['어휘', '분류'] + [f'{v}_{k}' for v in VERSIONS for k in ('복원', '삽입')])
            for t, g in sorted(term_group.items(), key=lambda x: (x[1], x[0])):
                w.writerow([t, g] + [round(term_tally[(v, t)][k], 2) for v in VERSIONS for k in ('복원', '삽입')])
    print(f'\n저장: {OUT_DIR}')


if __name__ == '__main__':
    main()
