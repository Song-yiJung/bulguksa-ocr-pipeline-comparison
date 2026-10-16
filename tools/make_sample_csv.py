"""실험용 표본목록.xlsx → 공개용 표본목록.csv 변환.

공개용 CSV 열: 파일id, 상대경로, 분석집합
  - 상대경로는 로컬 폴더 이름 대신 파일 ID에서 만든다: A057-004-097-001 → A057-004/A057-004-097-001.jpg
  - 분석집합은 원래 값('본 분석' / '짧은 면')을 그대로 쓴다.
  - --exclude 에 적은 파일 ID는 공개 목록에서 뺀다(경주 불국사와 무관하여 평가에서 제외한 면).

사용 예:
  python tools/make_sample_csv.py 표본목록.xlsx 표본목록.csv --exclude A031-004-011-001 A047-003-011-001
"""
import argparse
import csv
from pathlib import Path

import openpyxl


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx', type=Path)
    ap.add_argument('csv', type=Path)
    ap.add_argument('--exclude', nargs='*', default=[])
    args = ap.parse_args()

    ws = openpyxl.load_workbook(args.xlsx, read_only=True)['표본목록']
    it = ws.iter_rows(values_only=True)
    header = list(next(it))
    col = {h: i for i, h in enumerate(header)}

    rows, skipped = [], []
    for r in it:
        if not r or not r[0]:
            continue
        fid = str(r[col['파일ID']]).strip()
        if fid in args.exclude:
            skipped.append(fid)
            continue
        folder = '-'.join(fid.split('-')[:2])
        rows.append([fid, f'{folder}/{fid}.jpg', r[col['분석 집합']]])

    with open(args.csv, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['파일id', '상대경로', '분석집합'])
        w.writerows(rows)

    main_n = sum(1 for x in rows if x[2] == '본 분석')
    print(f'{len(rows)}면 기록 (본 분석 {main_n}, 그 밖 {len(rows) - main_n}) / 제외 {len(skipped)}면: {skipped}')
    missing = set(args.exclude) - set(skipped)
    if missing:
        print(f'[주의] 제외 목록에 있으나 표본에 없는 ID: {sorted(missing)}')


if __name__ == '__main__':
    main()
