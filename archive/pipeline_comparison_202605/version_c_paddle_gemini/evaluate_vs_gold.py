"""step3: 검수용 정답(_연구자판독)과 step2 산출(_gemini_final.txt) 비교 평가.

매칭 키 = 파일 stem (예: E006-005-003-001).
메트릭:
  - 글자 일치율(공백 제거 정규화 후 difflib SequenceMatcher)
  - 글자 수 차이
  - 라인 수
  - 누락/추가 글자 샘플
"""
import argparse
import difflib
import re
from collections import Counter
from pathlib import Path


GOLD_SUFFIX_RE = re.compile(r"_gemini_final_연구자[^.]*$")


def normalize(text: str, *, strip_spaces: bool = True) -> str:
    if strip_spaces:
        text = re.sub(r"\s+", "", text)
        text = text.replace("|", "")
    return text


def load_gold(gold_dir: Path) -> dict[str, Path]:
    files = {}
    for p in gold_dir.rglob("*.txt"):
        m = GOLD_SUFFIX_RE.search(p.stem)
        if not m:
            continue
        stem = p.stem[: m.start()]
        files[stem] = p
    return files


def load_results(results_dir: Path) -> dict[str, Path]:
    files = {}
    for p in results_dir.rglob("*_gemini_final.txt"):
        stem = p.stem.replace("_gemini_final", "")
        files[stem] = p
    return files


def compare_pair(gold_text: str, pred_text: str) -> dict:
    """정답(gold) vs 결과(pred) 비교 — 4가지 학술 메트릭 + 통계.

    메트릭 의미:
      - char_accuracy : 정답과 결과의 시퀀스 일치율 (위치 고려, difflib LCS 기반).
                         높을수록 정답과 같은 순서로 같은 글자를 출력.
      - precision     : 결과 글자가 정답에 있을 확률 (환각 反지표).
                         낮을수록 환각/추가 글자 많음.
      - recall        : 정답 글자가 결과에 있을 확률 (누락 反지표).
                         낮을수록 정답 글자를 빠뜨림.
      - f1            : Precision/Recall 조화평균 (종합 점수, 0~1).

    글자 단위 Multiset(Counter) 기반 P/R/F1 — 위치 무시, 빈도 고려:
      - "佛佛佛" vs "佛佛" → TP=min(3,2)=2, FN=1, FP=0
      - Precision=2/2=100%, Recall=2/3=67%
    """
    gold_n = normalize(gold_text)
    pred_n = normalize(pred_text)

    # 1. Character Accuracy — difflib SequenceMatcher.ratio() (위치 기반)
    matcher = difflib.SequenceMatcher(None, gold_n, pred_n, autojunk=False)
    char_accuracy = matcher.ratio()

    # 2~4. Precision / Recall / F1 — Multiset 기반 (위치 무시, 빈도 고려)
    gold_counter = Counter(gold_n)
    pred_counter = Counter(pred_n)
    tp = sum((gold_counter & pred_counter).values())  # min(count)
    fp = sum(pred_counter.values()) - tp
    fn = sum(gold_counter.values()) - tp
    precision = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    recall = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    f1 = (
        2 * precision * recall / (precision + recall)
        if (precision + recall) > 0
        else 0.0
    )

    gold_set = set(gold_n)
    pred_set = set(pred_n)
    return {
        "gold_chars": len(gold_n),
        "pred_chars": len(pred_n),
        "char_accuracy": char_accuracy,
        "precision": precision,
        "recall": recall,
        "f1": f1,
        "tp": tp,
        "fp": fp,
        "fn": fn,
        "missing_uniq": "".join(sorted(gold_set - pred_set))[:40],
        "extra_uniq": "".join(sorted(pred_set - gold_set))[:40],
        "gold_lines": gold_text.count("\n") + 1,
        "pred_lines": pred_text.count("\n") + 1,
    }


def print_unified_diff(stem: str, gold_text: str, pred_text: str) -> None:
    print(f"\n===== diff: {stem} (gold vs pred) =====")
    diff = difflib.unified_diff(
        gold_text.splitlines(keepends=False),
        pred_text.splitlines(keepends=False),
        fromfile="gold",
        tofile="pred",
        lineterm="",
        n=1,
    )
    for line in diff:
        print(line)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--gold-dir",
        default=r"D:\OCR_PIPLINE\results\검수용",
    )
    parser.add_argument(
        "--results-dir",
        default=r"D:\OCR_PIPLINE\results\version_C_v3_VL_html",
    )
    parser.add_argument(
        "--show-diff",
        action="store_true",
        help="정성 분석용: 모든 매칭 건의 unified diff 출력",
    )
    args = parser.parse_args()

    gold = load_gold(Path(args.gold_dir))
    pred = load_results(Path(args.results_dir))

    common = sorted(set(gold) & set(pred))
    only_gold = sorted(set(gold) - set(pred))
    only_pred = sorted(set(pred) - set(gold))

    print(f"[Gold ] {Path(args.gold_dir)}  ({len(gold)}건)")
    print(f"[Pred ] {Path(args.results_dir)}  ({len(pred)}건)")
    print(f"[Match] 공통 {len(common)}건  /  정답만 {len(only_gold)}건  /  결과만 {len(only_pred)}건")
    print("-" * 88)

    header = (
        f"{'파일ID':<24}"
        f"{'정답자':>6}"
        f"{'결과자':>6}"
        f"{'차이':>6}"
        f"{'일치율':>8}"
        f"  누락uniq / 추가uniq"
    )
    print(header)
    print("-" * 88)

    rows = []
    for stem in common:
        gtxt = gold[stem].read_text(encoding="utf-8")
        ptxt = pred[stem].read_text(encoding="utf-8")
        r = compare_pair(gtxt, ptxt)
        rows.append((stem, r, gtxt, ptxt))
        diff_chars = r["pred_chars"] - r["gold_chars"]
        sign = "+" if diff_chars >= 0 else ""
        print(
            f"{stem:<24}"
            f"{r['gold_chars']:>6}"
            f"{r['pred_chars']:>6}"
            f"{sign}{diff_chars:>5}"
            f"{r['char_accuracy']*100:>7.1f}%"
            f"  -{r['missing_uniq']}  +{r['extra_uniq']}"
        )

    if rows:
        avg = sum(r[1]["char_accuracy"] for r in rows) / len(rows)
        total_gold = sum(r[1]["gold_chars"] for r in rows)
        total_pred = sum(r[1]["pred_chars"] for r in rows)
        print("-" * 88)
        print(
            f"{'평균/합계':<24}"
            f"{total_gold:>6}"
            f"{total_pred:>6}"
            f"{'':>6}"
            f"{avg*100:>7.1f}%"
        )

    if only_gold:
        print(f"\n[정답만 있음(결과 없음)] {only_gold}")
    if only_pred:
        print(f"\n[결과만 있음(정답 없음)] {only_pred}")

    if args.show_diff:
        for stem, _r, gtxt, ptxt in rows:
            print_unified_diff(stem, gtxt, ptxt)


if __name__ == "__main__":
    main()
