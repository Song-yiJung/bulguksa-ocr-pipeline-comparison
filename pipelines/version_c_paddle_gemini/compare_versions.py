"""모든 OCR 파이프라인 결과를 정답(검수용)과 비교해 정확도 한 표로 출력하고
results/ 폴더에 CSV + TXT 리포트로 저장.

A/B/C/D 네 패러다임 비교:
  - A standard (OCR-first):   Vision → Gemini
  - B cloud_hybrid (LLM-first): Gemini → Vision + alignment
  - C paddle_standard (OCR-first): Paddle → Gemini
  - D paddle_hybrid (LLM-first): Gemini → Paddle + alignment

B/D의 step2 결과는 alignment 마커([G:...?] / [O:...?] / [G:...|O:...?])를 포함하므로
평가 시 마커와 안의 글자를 모두 제거하여 환각 차단된 텍스트로 비교.
"""
import csv
import re
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from evaluate_vs_gold import load_gold, compare_pair

RESULTS_ROOT = Path(r"D:\OCR_PIPLINE\results")
GOLD_DIR = RESULTS_ROOT / "검수용"
VERSION_A_OLD_ROOT = Path(r"D:\jggmk2_results\output_results")  # 이전 세션 v_A 결과 (보존용)

# (라벨, 폴더(string=RESULTS_ROOT 상대 / Path=절대), 패턴, 제거할 접미사, strip_markers)
VERSIONS = [
    # ── version_A: OCR-first (Vision → Gemini) ─────────────────────────
    ("v_A step1   (Vision OCR)",                   "version_A",            "E006-*[0-9].txt",          "",             False),
    ("v_A step1+2 (Vision + Gemini ★)",            "version_A",            "E006-*_gemini_final.txt",  "_gemini_final", False),
    ("v_A_old step1   (jggmk2_results)",           VERSION_A_OLD_ROOT,     "E006-*[0-9].txt",          "",             False),
    ("v_A_old step1+2 (jggmk2_results)",           VERSION_A_OLD_ROOT,     "E006-*_gemini_final.txt",  "_gemini_final", False),
    # ── version_B: LLM-first (Gemini → Vision + alignment) ─────────────
    ("v_B step1   (Gemini-first 단독)",            "version_B",            "E006-*[0-9].txt",          "",             False),
    ("v_B step1+2 (Gemini+Vision alignment ★)",    "version_B",            "E006-*_final.txt",         "_final",       True),
    # ── version_C: OCR-first variants ─────────────────────────────────
    ("v_C v1_japan",                               "version_C",            "E006-*.txt",               "",             False),
    ("v_C v1_cht",                                 "version_C_cht",        "E006-*.txt",               "",             False),
    ("v_C v2_tuned",                               "version_C_v2_tuned",   "E006-*.txt",               "",             False),
    ("v_C v3_VL",                                  "version_C_v3_VL",      "E006-*.txt",               "",             False),
    ("v_C v3_VL_html md",                          "version_C_v3_VL_html", "E006-*[0-9].txt",          "",             False),
    ("v_C v3_VL_html clean",                       "version_C_v3_VL_html", "E006-*_clean.txt",         "_clean",       False),
    ("v_C v3_VL_html GEMINI ★",                    "version_C_v3_VL_html", "E006-*_gemini_final.txt",  "_gemini_final", False),
    # ── version_D: LLM-first (Gemini → Paddle + alignment) ─────────────
    ("v_D step1   (Gemini-first 단독)",            "version_D",            "E006-*[0-9].txt",          "",             False),
    ("v_D step1+2 (Gemini+Paddle alignment ★)",    "version_D",            "E006-*_final.txt",         "_final",       True),
]


# Alignment 마커 제거 패턴 (B/D step2 결과 평가용)
_REPLACE_MARKER = re.compile(r'\[G:[^|\]]*\|O:[^\]]*\?\]')
_SIMPLE_MARKER = re.compile(r'\[[GO]:[^\]]*\?\]')


def strip_alignment_markers(text: str) -> str:
    """[G:글자|O:글자?] / [G:글자?] / [O:글자?] 마커와 내용을 모두 제거.
    환각/누락 후보 글자를 제거하여 '신뢰 가능한 글자만' 남김 → 환각 차단 효과 평가."""
    text = _REPLACE_MARKER.sub('', text)
    text = _SIMPLE_MARKER.sub('', text)
    return text


def resolve_dir(spec) -> Path:
    """spec이 Path면 그대로, string이면 RESULTS_ROOT 기준 상대 결합."""
    return spec if isinstance(spec, Path) else (RESULTS_ROOT / spec)


def load_preds(version_dir: Path, pattern: str, strip_suffix: str) -> dict:
    out = {}
    for p in version_dir.rglob(pattern):
        if "_연구자" in p.stem:  # 정답 파일 제외
            continue
        stem = p.stem
        if strip_suffix and stem.endswith(strip_suffix):
            stem = stem[: -len(strip_suffix)]
        out[stem] = p
    return out


from collections import Counter
from itertools import combinations


# 4파이프라인 대표 (인간 합의 측정용) — 각 패러다임의 최종 산출
PARADIGM_REPRESENTATIVES = [
    ("A (Vision+Gemini)",         "version_A",            "E006-*_gemini_final.txt", "_gemini_final", False),
    ("B (Gemini+Vision)",         "version_B",            "*_final.txt",             "_final",       True),
    ("C (Paddle+Gemini)",         "version_C_v3_VL_html", "*_gemini_final.txt",      "_gemini_final", False),
    ("D (Gemini+Paddle)",         "version_D",            "*_final.txt",             "_final",       True),
]


def _normalize_for_agreement(text: str, strip_markers: bool) -> str:
    """합의 측정용 정규화: alignment 마커 제거 + 공백 제거."""
    if strip_markers:
        text = strip_alignment_markers(text)
    text = re.sub(r"\s+", "", text)
    text = text.replace("|", "")
    return text


def compute_inter_model_agreement(stem: str, model_texts: dict) -> dict:
    """정답이 없는 페이지에 대한 4모델 합의도 측정.

    Args:
        stem: 페이지 식별자 (예: H052-010-017-001)
        model_texts: dict[모델라벨 -> 정규화된 텍스트]
    Returns:
        dict[메트릭 -> value]:
          - n_models       : 측정 대상 모델 수
          - strict_ratio   : 모든 모델에 공통 글자 비율 (multiset 교집합 / 합집합)
          - majority_ratio : 과반 모델에 출현한 글자 비율
          - avg_jaccard    : 모델 페어별 Jaccard similarity 평균
          - pair_jaccards  : 페어별 Jaccard (디버그용)
    """
    counters = {label: Counter(txt) for label, txt in model_texts.items() if txt}
    n_models = len(counters)
    if n_models < 2:
        return {"n_models": n_models, "strict_ratio": 0.0, "majority_ratio": 0.0, "avg_jaccard": 0.0, "pair_jaccards": []}

    # Strict: 모든 모델에 공통 (multiset 교집합)
    strict = None
    for c in counters.values():
        strict = c.copy() if strict is None else (strict & c)
    strict_chars = sum(strict.values())

    # Union (multiset 합집합 — max count)
    union = Counter()
    for c in counters.values():
        union |= c
    union_chars = sum(union.values())

    # Majority: 각 글자가 등장하는 모델 수가 (n//2 + 1) 이상이면 majority
    majority_threshold = n_models // 2 + 1
    char_appears_in = Counter()
    for c in counters.values():
        for ch in c:
            char_appears_in[ch] += 1
    majority_chars_set = {ch for ch, n in char_appears_in.items() if n >= majority_threshold}
    # 각 majority 글자의 빈도 = median (간단히는 min)
    majority_chars = sum(
        min(c[ch] for c in counters.values() if ch in c)
        for ch in majority_chars_set
    )

    # Jaccard per pair (multiset 교집합 / 합집합)
    pair_jaccards = []
    for (a_lbl, a_c), (b_lbl, b_c) in combinations(counters.items(), 2):
        inter = sum((a_c & b_c).values())
        un = sum((a_c | b_c).values())
        j = inter / un if un > 0 else 0.0
        pair_jaccards.append((f"{a_lbl} ↔ {b_lbl}", j))
    avg_jaccard = sum(j for _, j in pair_jaccards) / len(pair_jaccards) if pair_jaccards else 0.0

    return {
        "n_models": n_models,
        "strict_ratio": strict_chars / union_chars if union_chars > 0 else 0.0,
        "majority_ratio": majority_chars / union_chars if union_chars > 0 else 0.0,
        "avg_jaccard": avg_jaccard,
        "strict_chars": strict_chars,
        "majority_chars": majority_chars,
        "union_chars": union_chars,
        "pair_jaccards": pair_jaccards,
    }


def run_agreement_analysis(both):
    """4파이프라인 대표를 모아 정답 없는 페이지 합의도 측정 + 출력."""
    # 1. 각 대표 모델의 결과 수집 (stem -> normalized text)
    model_results = {}  # label -> dict[stem -> text]
    all_stems = set()
    for label, vdir, pattern, suffix, strip_markers in PARADIGM_REPRESENTATIVES:
        version_path = resolve_dir(vdir)
        if not version_path.exists():
            continue
        preds = load_preds(version_path, pattern, suffix)
        stem_texts = {}
        for stem, p in preds.items():
            txt = p.read_text(encoding="utf-8")
            stem_texts[stem] = _normalize_for_agreement(txt, strip_markers)
            all_stems.add(stem)
        model_results[label] = stem_texts

    # 2. 정답 없는 페이지 식별 (검수용에 없는 stem)
    gold_stems = set(load_gold(GOLD_DIR).keys())
    no_gold_stems = sorted(all_stems - gold_stems)

    if not no_gold_stems:
        both("\n[합의 측정] 정답 없는 페이지가 없습니다 — 측정 건너뜀")
        return

    both()
    both("=" * 70)
    both(f"[인간 합의 측정 — 정답 없는 {len(no_gold_stems)}장]")
    both()
    both("  메트릭 설명:")
    both("    Strict   : 모든 모델 공통 글자 비율 (multiset 교집합 / 합집합)")
    both("    Majority : 과반 모델 합의 글자 비율 (4모델이면 3+, 2모델이면 2)")
    both("    Jaccard  : 모델 페어별 multiset 교집합/합집합 평균")
    both()
    both("  해석:")
    both("    - 합의도 높음 (≥ 60%) → 자동 결과 신뢰 가능 (인쇄체)")
    both("    - 합의도 중간 (30~60%) → 부분 신뢰, 인간 검수 권장")
    both("    - 합의도 낮음 (< 30%) → 인간 판독 필수 (흘림체/모호)")
    both()

    header = f"{'페이지':<28}{'#모델':>6}{'Strict':>10}{'Majority':>11}{'avgJaccard':>13}  비고"
    both(header)
    both("-" * len(header))

    rows = []
    for stem in no_gold_stems:
        model_texts = {
            label: model_results[label].get(stem, "")
            for label in model_results
            if stem in model_results[label]
        }
        if not model_texts:
            continue
        r = compute_inter_model_agreement(stem, model_texts)
        n = r["n_models"]
        if n < 2:
            both(f"{stem:<28}{n:>6}  (모델 < 2개, 측정 불가)")
            continue

        # 해석 카테고리
        s = r["strict_ratio"] * 100
        if s >= 60:
            note = "✓ 신뢰 가능"
        elif s >= 30:
            note = "△ 부분 신뢰"
        else:
            note = "⚠ 인간 판독 권장"

        line = (
            f"{stem:<28}{n:>6}"
            f"{r['strict_ratio']*100:>9.1f}%"
            f"{r['majority_ratio']*100:>10.1f}%"
            f"{r['avg_jaccard']*100:>12.1f}%  {note}"
        )
        both(line)
        rows.append({
            "stem": stem,
            "n_models": n,
            "strict_ratio": r["strict_ratio"],
            "majority_ratio": r["majority_ratio"],
            "avg_jaccard": r["avg_jaccard"],
            "note": note,
        })

    both("-" * len(header))

    # 통계 요약
    if rows:
        n_total = len(rows)
        n_high = sum(1 for r in rows if r["strict_ratio"] >= 0.6)
        n_mid = sum(1 for r in rows if 0.3 <= r["strict_ratio"] < 0.6)
        n_low = sum(1 for r in rows if r["strict_ratio"] < 0.3)
        avg_strict = sum(r["strict_ratio"] for r in rows) / n_total * 100
        avg_majority = sum(r["majority_ratio"] for r in rows) / n_total * 100
        avg_jaccard = sum(r["avg_jaccard"] for r in rows) / n_total * 100

        both()
        both(f"[합의 측정 통계 — n={n_total}]")
        both(f"  평균 Strict   : {avg_strict:.1f}%")
        both(f"  평균 Majority : {avg_majority:.1f}%")
        both(f"  평균 Jaccard  : {avg_jaccard:.1f}%")
        both()
        both(f"  분류:")
        both(f"    ✓ 신뢰 가능  (Strict ≥ 60%): {n_high}장")
        both(f"    △ 부분 신뢰  (Strict 30~60%): {n_mid}장")
        both(f"    ⚠ 인간 판독  (Strict < 30%): {n_low}장")

    return rows


def main() -> None:
    gold = load_gold(GOLD_DIR)
    gold_stems = sorted(gold.keys())

    out_lines = []  # txt 리포트용 누적

    def both(s: str = "") -> None:
        print(s)
        out_lines.append(s)

    both(f"[Gold] {GOLD_DIR}  ({len(gold)}건)")
    both(f"[생성] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    both(f"[비교 대상] {len(VERSIONS)}개 버전 / {len(gold_stems)}장 매칭")
    both()

    # ============================================================
    # 메트릭 설명 (학습자/연구자 이해 보조)
    # ============================================================
    both("[메트릭 설명]")
    both("  CharAcc   : 정답과 결과의 시퀀스 일치율 (위치 고려, difflib LCS)")
    both("              높을수록 정답과 같은 순서로 같은 글자 출력")
    both("  Precision : 결과 글자가 정답에 있을 확률 (환각 反지표, 0~100%)")
    both("              낮을수록 환각/추가 글자 많음")
    both("  Recall    : 정답 글자가 결과에 있을 확률 (누락 反지표, 0~100%)")
    both("              낮을수록 정답 글자를 빠뜨림")
    both("  F1        : Precision/Recall 조화평균 (종합 점수, 0~100%)")
    both("              종합 평가, 균형 잡힌 트레이드오프 측정")
    both()
    both("  주의: Precision/Recall은 Multiset(Counter) 기반 — 위치 무시, 빈도 고려")
    both("        (예: '佛佛佛' vs '佛佛' → TP=2, FN=1, P=100%, R=67%)")
    both()

    # ============================================================
    # 메인 표 — 4 메트릭 평균
    # ============================================================
    header = f"{'버전':<46}{'n':>4}  {'CharAcc':>8}  {'Prec':>7}  {'Recall':>7}  {'F1':>7}"
    both(header)
    both("-" * len(header))

    rows = []  # 각 버전별 결과
    for label, vdir, pattern, suffix, strip_markers in VERSIONS:
        version_path = resolve_dir(vdir)
        if not version_path.exists():
            both(f"{label:<46}  (폴더 없음)")
            continue

        preds = load_preds(version_path, pattern, suffix)
        per_stem = {}  # stem -> compare_pair result dict
        for stem in gold_stems:
            if stem not in preds:
                continue
            gtxt = gold[stem].read_text(encoding="utf-8")
            ptxt = preds[stem].read_text(encoding="utf-8")
            if strip_markers:
                ptxt = strip_alignment_markers(ptxt)
            r = compare_pair(gtxt, ptxt)
            per_stem[stem] = r

        if not per_stem:
            both(f"{label:<46}  (매칭 0건)")
            continue

        n = len(per_stem)
        avg_acc = sum(r["char_accuracy"] for r in per_stem.values()) / n
        avg_prec = sum(r["precision"] for r in per_stem.values()) / n
        avg_rec = sum(r["recall"] for r in per_stem.values()) / n
        avg_f1 = sum(r["f1"] for r in per_stem.values()) / n
        line = (
            f"{label:<46}{n:>4}  "
            f"{avg_acc*100:>7.1f}%  "
            f"{avg_prec*100:>6.1f}%  "
            f"{avg_rec*100:>6.1f}%  "
            f"{avg_f1*100:>6.1f}%"
        )
        both(line)
        rows.append({
            "label": label,
            "n": n,
            "avg_acc": avg_acc,
            "avg_prec": avg_prec,
            "avg_rec": avg_rec,
            "avg_f1": avg_f1,
            "per_stem": per_stem,
        })

    both("-" * len(header))
    if rows:
        best_acc = max(rows, key=lambda r: r["avg_acc"])
        best_f1 = max(rows, key=lambda r: r["avg_f1"])
        both()
        both(f"[CharAcc 1위] {best_acc['label']} — {best_acc['avg_acc']*100:.1f}%")
        both(f"[F1 1위]      {best_f1['label']} — {best_f1['avg_f1']*100:.1f}%")

    # ============================================================
    # 페이지별 표 — CharAcc + F1 (요약)
    # ============================================================
    if rows:
        both()
        both("=" * 70)
        both("[페이지별 Char Accuracy %]")
        sub_header = f"{'버전':<46}  " + "  ".join(
            f"{s.replace('E006-005-', ''):>9}" for s in gold_stems
        )
        both(sub_header)
        both("-" * len(sub_header))
        for r in rows:
            line = f"{r['label']:<46}  " + "  ".join(
                f"{r['per_stem'][s]['char_accuracy']*100:>8.1f}%" if s in r['per_stem'] else f"{'-':>9}"
                for s in gold_stems
            )
            both(line)

        both()
        both("[페이지별 F1 %]")
        both(sub_header)
        both("-" * len(sub_header))
        for r in rows:
            line = f"{r['label']:<46}  " + "  ".join(
                f"{r['per_stem'][s]['f1']*100:>8.1f}%" if s in r['per_stem'] else f"{'-':>9}"
                for s in gold_stems
            )
            both(line)

    # ============================================================
    # 인간 합의 측정 — 정답 없는 페이지 (4파이프라인 대표 비교)
    # ============================================================
    agreement_rows = run_agreement_analysis(both)

    # ============================================================
    # 결과 파일 저장 (TXT + CSV)
    # ============================================================
    out_dir = RESULTS_ROOT
    txt_path = out_dir / "_evaluation_compare_versions.txt"
    csv_path = out_dir / "_evaluation_compare_versions.csv"
    agreement_csv_path = out_dir / "_evaluation_inter_model_agreement.csv"

    # 합의도 CSV 저장
    if agreement_rows:
        with open(agreement_csv_path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["stem", "n_models", "strict_ratio_pct", "majority_ratio_pct", "avg_jaccard_pct", "category"])
            for r in agreement_rows:
                w.writerow([
                    r["stem"], r["n_models"],
                    f"{r['strict_ratio']*100:.2f}",
                    f"{r['majority_ratio']*100:.2f}",
                    f"{r['avg_jaccard']*100:.2f}",
                    r["note"],
                ])
        print(f"[저장] {agreement_csv_path}")

    txt_path.write_text("\n".join(out_lines), encoding="utf-8")

    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        # CSV 헤더: 평균 4개 + 페이지별 4메트릭
        header_row = (
            ["version", "n", "avg_char_acc_pct", "avg_precision_pct",
             "avg_recall_pct", "avg_f1_pct"]
            + [f"{s}_acc" for s in gold_stems]
            + [f"{s}_prec" for s in gold_stems]
            + [f"{s}_recall" for s in gold_stems]
            + [f"{s}_f1" for s in gold_stems]
        )
        w.writerow(header_row)
        for r in rows:
            row = [
                r["label"], r["n"],
                f"{r['avg_acc']*100:.2f}",
                f"{r['avg_prec']*100:.2f}",
                f"{r['avg_rec']*100:.2f}",
                f"{r['avg_f1']*100:.2f}",
            ]
            for s in gold_stems:
                row.append(f"{r['per_stem'][s]['char_accuracy']*100:.2f}" if s in r['per_stem'] else "")
            for s in gold_stems:
                row.append(f"{r['per_stem'][s]['precision']*100:.2f}" if s in r['per_stem'] else "")
            for s in gold_stems:
                row.append(f"{r['per_stem'][s]['recall']*100:.2f}" if s in r['per_stem'] else "")
            for s in gold_stems:
                row.append(f"{r['per_stem'][s]['f1']*100:.2f}" if s in r['per_stem'] else "")
            w.writerow(row)

    print()
    print(f"[저장] {txt_path}")
    print(f"[저장] {csv_path}")


if __name__ == "__main__":
    main()
