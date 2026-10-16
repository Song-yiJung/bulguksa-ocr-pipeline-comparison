import os
import json
import time
import re
import html as html_lib
from pathlib import Path
from datetime import date, datetime
from dotenv import load_dotenv

# 1. 환경 변수 로드
PROJECT_ROOT = Path(__file__).parent.parent
load_dotenv(dotenv_path=PROJECT_ROOT / ".env")

from paddleocr import PaddleOCRVL
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 버전 메타데이터
# ============================================================
VERSION_TAG = "v3_VL_html"
VERSION_NAME = f"Version C-{VERSION_TAG} (PaddleOCR-VL + HTML 분리, 전처리 없음)"
CHANGE_REASON = (
    "v3_VL의 markdown 결과에 HTML 테이블 태그가 plain_text에 섞여 가독성 저하. "
    "OpenCV 전처리는 CPU 환경에서 처리 시간 부담이 커서 제외 결정. "
    "OCR 자체는 v3_VL과 동일하되, 두 가지 출력 분리: "
    "(.md, .txt = HTML 원본 보존), (_clean.txt = HTML 제거 연구자용)."
)
COMPARE_AGAINST = (
    "v3_VL: results/version_C_v3_VL/ (HTML 미분리, plain_text에 태그 섞임)\n"
    "  v3_VL_pp: results/version_C_v3_VL_pp/ (전처리 시도, CPU 환경 부담으로 보류)\n"
    "  v1_japan: results/version_C/ (PP-OCRv5 베이스라인)"
)

# 2. 경로 설정
INPUT_DIR  = Path(r'G:\jggmk2\OCR_PIPLINE_IMAGES')
OUTPUT_DIR = Path(rf'D:\OCR_PIPLINE\results\version_C_{VERSION_TAG}')
SOURCE_XLSX = INPUT_DIR / 'OCR데이터목록.xlsx'
LOG_XLSX    = OUTPUT_DIR / f'version_C_{VERSION_TAG}_작업내역.xlsx'
VERSION_INFO_TXT = OUTPUT_DIR / '_version_info.txt'
POLL_INTERVAL = 60

# ============================================================
# 자동 트랙 제외 목록 (PaddleOCR-VL 디코더 토큰 폭주 회피)
# 흘림체 손글씨처럼 종료 토큰 못 잡는 페이지는 인간 메타 보조 트랙으로 별도 처리.
# 연구노트 _연구노트_도메인별예상_2026-05-12.md 섹션 11번 참조.
# 추가 시: 파일 stem (확장자 제외) 입력. 예: "A059-024-003-001"
# ============================================================
SKIP_FILE_IDS = {
    "A059-024-003-001",  # 흘림체 한자 편지 — Paddle VL 디코더 토큰 폭주로 30분+ 무한 생성
    "A059-024-007-001",  # 전보 1 — Paddle VL 디코더 폭주 확정 (2회 시도 모두 무한 진행)
}


def write_version_info():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    lines = [
        "=" * 60,
        f"Version C (PaddleOCR) - {VERSION_TAG}",
        "=" * 60,
        f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"실행 코드: {Path(__file__).name}",
        "",
        "[OCR 엔진]",
        "  - PaddleOCR-VL-1.5 (0.9B 시각언어모델 VLM)",
        "  - PP-DocLayoutV3 (레이아웃 분석)",
        "",
        "[전처리]",
        "  - 없음 (원본 이미지 그대로 OCR 호출)",
        "",
        "[후처리: HTML 정리]",
        "  - <img> 태그 → [Image] 또는 [인장] 표시",
        "  - <td>, <th> 셀 사이 구분자 ' | '",
        "  - <tr> 행 단위 개행 처리",
        "  - 나머지 HTML 태그 모두 제거",
        "  - HTML 엔티티 디코딩 (&amp; 등)",
        "",
        "[출력 결과]",
        "  - .json: 전체 메타데이터 통합",
        "  - .md: 원본 마크다운 (HTML 테이블 포함, 구조 보존)",
        "  - .txt: HTML 원본 (구조 분석용, .md와 동일 내용)",
        "  - _clean.txt: HTML 제거 깔끔 텍스트 (연구자/Gemini 입력용) ⭐",
        "",
        "[변경 사유]",
        f"  {CHANGE_REASON}",
        "",
        "[비교 대상]",
        f"  {COMPARE_AGAINST}",
        "",
        "=" * 60,
    ]
    VERSION_INFO_TXT.write_text("\n".join(lines), encoding="utf-8")


# 3. PaddleOCR-VL 초기화
print(f"[System] PaddleOCR-VL 모델 로드 중... ({VERSION_TAG})")
print(f"[System] CPU 환경에서 가중치 로딩에 1~3분 소요...")
try:
    ocr_engine = PaddleOCRVL()
    print(f"[System] PaddleOCR-VL 준비 완료\n")
    write_version_info()
    print(f"[System] 버전 정보 기록: {VERSION_INFO_TXT.name}\n")
except Exception as e:
    print(f"[Error] PaddleOCR-VL 초기화 실패: {e}")
    exit(1)


def _expand_range(s: str):
    m = re.match(r'(.+-)(\d+)~.+?(\d+)$', s)
    if m:
        p, start, end = m.group(1), int(m.group(2)), int(m.group(3))
        w = len(m.group(2))
        return {f'{p}{str(i).zfill(w)}' for i in range(start, end + 1)}
    return {s}


def html_table_to_text(html_text: str) -> str:
    """HTML 테이블이 포함된 마크다운에서 깔끔한 텍스트 추출"""
    if not html_text:
        return ''
    
    text = html_text
    
    # 1) <img> 태그 → [이미지/도장] 표시
    text = re.sub(r'<img[^>]*alt="([^"]*)"[^>]*/?>', r'[\1]', text)
    text = re.sub(r'<img[^>]*/?>', '[이미지]', text)
    
    # 2) 셀 구분자 처리
    text = re.sub(r'</td>\s*<td[^>]*>', ' | ', text)
    text = re.sub(r'</th>\s*<th[^>]*>', ' | ', text)
    
    # 3) 행 구분자 → 개행
    text = re.sub(r'</tr>\s*<tr[^>]*>', '\n', text)
    
    # 4) 나머지 HTML 태그 모두 제거
    text = re.sub(r'<[^>]+>', '', text)
    
    # 5) HTML 엔티티 디코딩
    text = html_lib.unescape(text)
    
    # 6) 마크다운 문법 제거
    text = re.sub(r'#+\s*', '', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    
    # 7) 공백 정리
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n+', '\n\n', text)
    text = re.sub(r' \| \n', '\n', text)
    
    return text.strip()


def extract_text_from_image(img_path: Path):
    """PaddleOCR-VL 추출 + HTML 정리 (전처리 없음)"""
    # 원본 이미지 직접 OCR 호출
    output = ocr_engine.predict(str(img_path))
    
    if not output:
        return {'markdown': '', 'clean_text': '', 'plain_text': '', 'json_data': None}
    
    result = output[0] if isinstance(output, list) else output
    
    markdown = ''
    json_data = None
    
    try:
        if hasattr(result, 'markdown'):
            md_obj = result.markdown
            if isinstance(md_obj, dict):
                markdown = md_obj.get('markdown_texts', '') or md_obj.get('text', '')
            else:
                markdown = str(md_obj)
        elif hasattr(result, 'get'):
            md_obj = result.get('markdown', {})
            if isinstance(md_obj, dict):
                markdown = md_obj.get('markdown_texts', '') or md_obj.get('text', '')
            else:
                markdown = str(md_obj) if md_obj else ''
    except Exception as e:
        print(f"   [warning] markdown 추출 실패: {e}")
    
    try:
        if hasattr(result, 'json'):
            json_data = result.json
        elif hasattr(result, 'get'):
            json_data = dict(result)
    except Exception:
        pass
    
    clean_text = html_table_to_text(markdown)
    
    return {
        'markdown': markdown,
        'clean_text': clean_text,
        'plain_text': clean_text,
        'json_data': json_data,
    }


def count_real_chars(text: str) -> int:
    """공백/개행/탭/구분자 제외한 실제 글자 수"""
    return len(text.replace("\n", "").replace(" ", "").replace("\t", "").replace("|", ""))


def write_version_xlsx(completed_ids: set):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    today = date.today().isoformat()

    source_rows = []
    source_headers = []
    if SOURCE_XLSX.exists():
        try:
            src_wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
            if 'Sheet1' in src_wb.sheetnames:
                ws = src_wb['Sheet1']
                rows_iter = ws.iter_rows(values_only=True)
                source_headers = list(next(rows_iter, []) or [])
                for row in rows_iter:
                    if any(c is not None and str(c).strip() for c in row):
                        source_rows.append(list(row))
            src_wb.close()
        except Exception as e:
            print(f"  [원본 엑셀 읽기 실패(무시): {str(e)[:120]}]")

    wb = openpyxl.Workbook()

    header_font = Font(name='맑은 고딕', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
    title_font  = Font(name='맑은 고딕', size=14, bold=True, color='305496')
    center      = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_wrap   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    thin = Side(border_style='thin', color='B4B4B4')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 시트 1: 작업현황
    ws1 = wb.active
    ws1.title = '작업현황'
    ws1['A1'] = f'{VERSION_NAME} - 작업 현황'
    ws1['A1'].font = title_font
    ws1.row_dimensions[1].height = 22

    if source_headers:
        new_headers = list(source_headers) + [f'step1_C_{VERSION_TAG}_완료일']
        for col_idx, h in enumerate(new_headers, 1):
            cell = ws1.cell(row=3, column=col_idx, value=str(h) if h is not None else '')
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        for row_idx, row_data in enumerate(source_rows, start=4):
            for col_idx, val in enumerate(row_data, 1):
                cell = ws1.cell(row=row_idx, column=col_idx, value=val)
                cell.alignment = left_wrap
                cell.border = border

            files_value = row_data[13] if len(row_data) > 13 else None
            if files_value:
                row_ids = set()
                for part in str(files_value).split('|'):
                    row_ids.update(_expand_range(part.strip()))
                if row_ids and row_ids.issubset(completed_ids):
                    last_col = len(new_headers)
                    cell = ws1.cell(row=row_idx, column=last_col, value=today)
                    cell.alignment = center
                    cell.border = border

        for col_idx in range(1, len(new_headers) + 1):
            ws1.column_dimensions[get_column_letter(col_idx)].width = 16
    else:
        ws1['A3'] = '원본 OCR데이터목록.xlsx 파일을 찾을 수 없거나 Sheet1이 없습니다.'
        ws1['A3'].font = Font(italic=True, color='808080')

    # 시트 2: 상세로그
    ws2 = wb.create_sheet('상세로그')
    ws2['A1'] = f'{VERSION_NAME} - 상세 처리 로그'
    ws2['A1'].font = title_font
    ws2.row_dimensions[1].height = 22

    detail_headers = ['순번', '처리일자', '권', '문건', '문서제목',
                      '파일ID', '실제문자수', '마크다운길이', '처리시간(초)', '상태', '경로']
    for col_idx, h in enumerate(detail_headers, 1):
        cell = ws2.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    records = []
    for f in OUTPUT_DIR.rglob('*.json'):
        if f.name.startswith('_') or any(pt.startswith('_') for pt in f.relative_to(OUTPUT_DIR).parts):
            continue
        try:
            d = json.loads(f.read_text(encoding='utf-8'))
        except Exception:
            continue

        rel = f.relative_to(OUTPUT_DIR)
        parts = rel.parts
        mod_date = datetime.fromtimestamp(f.stat().st_mtime).date().isoformat()

        if 'error' in d:
            status = '오류'
            real_chars = 0
            md_len = 0
            ptime = 0
        else:
            clean = d.get('vl_clean_text', '') or d.get('vl_plain_text', '')
            md = d.get('vl_markdown', '')
            real_chars = count_real_chars(clean)
            md_len = len(md)
            ptime = d.get('processing_time_sec', 0)
            status = '성공' if real_chars > 0 else '텍스트없음'

        records.append({
            'date': mod_date,
            'kwon': parts[0] if len(parts) > 0 else '',
            'mun':  parts[1] if len(parts) > 1 else '',
            'title':parts[2] if len(parts) > 2 else '',
            'file_id': f.stem,
            'real_chars': real_chars,
            'md_len': md_len,
            'ptime': ptime,
            'status': status,
            'path': str(rel),
        })

    records.sort(key=lambda r: (r['date'], r['path']))
    for i, r in enumerate(records, 1):
        row_idx = 3 + i
        values = [i, r['date'], r['kwon'], r['mun'], r['title'],
                  r['file_id'], r['real_chars'], r['md_len'], r['ptime'], r['status'], r['path']]
        for col_idx, v in enumerate(values, 1):
            cell = ws2.cell(row=row_idx, column=col_idx, value=v)
            if col_idx in (1, 2, 6, 7, 8, 9, 10):
                cell.alignment = center
            else:
                cell.alignment = left_wrap
            cell.border = border

            if col_idx == 10:
                if r['status'] == '오류':
                    cell.fill = PatternFill(start_color='F8CBAD', end_color='F8CBAD', fill_type='solid')
                elif r['status'] == '텍스트없음':
                    cell.fill = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
                else:
                    cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    widths = [6, 12, 25, 22, 25, 22, 10, 12, 12, 12, 60]
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = 'A4'

    # 시트 3: 요약통계
    ws3 = wb.create_sheet('요약통계')
    ws3['A1'] = f'{VERSION_NAME} - 요약 통계'
    ws3['A1'].font = title_font
    ws3.row_dimensions[1].height = 22

    total = len(records)
    success = sum(1 for r in records if r['status'] == '성공')
    no_text = sum(1 for r in records if r['status'] == '텍스트없음')
    error_count = sum(1 for r in records if r['status'] == '오류')
    total_real_chars = sum(r['real_chars'] for r in records)
    avg_chars = (total_real_chars / success) if success > 0 else 0
    avg_time = (sum(r['ptime'] for r in records) / total) if total > 0 else 0

    summary = [
        ('파이프라인', VERSION_NAME),
        ('생성일시',   datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ('', ''),
        ('총 처리 파일 수',     total),
        ('  ├ 성공',           success),
        ('  ├ 텍스트 없음',     no_text),
        ('  └ 오류',           error_count),
        ('', ''),
        ('총 실제 문자 수 (HTML 제외)', total_real_chars),
        ('성공 건당 평균 문자수',  f'{avg_chars:.1f}'),
        ('평균 처리 시간 (초)',   f'{avg_time:.1f}'),
        ('', ''),
        ('OCR 엔진', 'PaddleOCR-VL-1.5 (0.9B VLM)'),
        ('전처리', '없음 (원본 이미지 그대로)'),
        ('후처리', 'HTML → 텍스트 정리 (셀 구분자 |, 행 단위 개행)'),
        ('', ''),
        ('출력 파일 — .md', '원본 마크다운 (HTML 포함)'),
        ('출력 파일 — .txt', 'HTML 원본 (.md와 동일 내용)'),
        ('출력 파일 — _clean.txt', 'HTML 제거 (연구자 본인 검수용) ⭐'),
        ('출력 파일 — .json', '메타데이터 통합 (markdown + clean 모두)'),
        ('', ''),
        ('변경 사유', CHANGE_REASON),
        ('비교 대상', COMPARE_AGAINST),
        ('', ''),
        ('입력 디렉토리', str(INPUT_DIR)),
        ('출력 디렉토리', str(OUTPUT_DIR)),
    ]

    for i, (label, value) in enumerate(summary, start=3):
        ws3.cell(row=i, column=1, value=label).font = Font(bold=True)
        ws3.cell(row=i, column=2, value=value).alignment = left_wrap
    ws3.column_dimensions['A'].width = 32
    ws3.column_dimensions['B'].width = 70

    try:
        wb.save(LOG_XLSX)
        print(f"  [엑셀 작업내역 갱신: {LOG_XLSX.name} - 총 {total}건]")
    except PermissionError:
        print(f"  [오류: {LOG_XLSX.name} 파일이 열려있어 저장 실패]")
    except Exception as e:
        print(f"  [엑셀 저장 실패(무시): {str(e)[:120]}]")


def process_pipeline():
    if not INPUT_DIR.exists():
        print(f"[Error] 입력 경로를 찾을 수 없습니다. ({INPUT_DIR})")
        return 0

    valid_extensions = {'.png', '.jpg', '.jpeg', '.tif', '.tiff'}

    image_files_all = sorted([
        p for p in INPUT_DIR.rglob('*')
        if p.suffix.lower() in valid_extensions
        and not any(pt.startswith('_') for pt in p.relative_to(INPUT_DIR).parts)
    ])

    # 자동 트랙 제외 목록 적용 (디코더 토큰 폭주 회피)
    skipped_by_id = [p for p in image_files_all if p.stem in SKIP_FILE_IDS]
    image_files = [p for p in image_files_all if p.stem not in SKIP_FILE_IDS]
    if skipped_by_id:
        print(f"[Skip] 자동 트랙 제외 페이지 ({len(skipped_by_id)}장, SKIP_FILE_IDS):")
        for p in skipped_by_id:
            print(f"  - {p.relative_to(INPUT_DIR)}  (메타 보조 트랙으로 별도 처리)")
        print()

    if not image_files:
        return 0

    todo_new = []
    todo_retry = []
    skip_done = []

    for img_path in image_files:
        rel_path = img_path.relative_to(INPUT_DIR)
        json_path = OUTPUT_DIR / rel_path.parent / f"{img_path.stem}.json"

        if not json_path.exists():
            todo_new.append(img_path)
            continue

        try:
            data = json.loads(json_path.read_text(encoding='utf-8'))
            if 'error' in data:
                todo_retry.append(img_path)
            elif 'vl_markdown' in data or 'vl_clean_text' in data:
                skip_done.append(img_path)
            else:
                todo_retry.append(img_path)
        except Exception:
            todo_retry.append(img_path)

    todo_total = todo_new + todo_retry
    if not todo_total:
        return 0

    print(f"[Scan] 입력 폴더 스캔 완료")
    print(f"  - 전체 이미지   : {len(image_files):>5}장")
    print(f"  - 처리 완료(스킵): {len(skip_done):>5}장")
    print(f"  - 신규 처리 대상: {len(todo_new):>5}장")
    if todo_retry:
        print(f"  - 에러 재시도   : {len(todo_retry):>5}장")
    print("-" * 60)

    processed = 0
    for i, img_path in enumerate(todo_total, 1):
        rel_path = img_path.relative_to(INPUT_DIR)
        out_folder = OUTPUT_DIR / rel_path.parent
        out_folder.mkdir(parents=True, exist_ok=True)

        base_name = img_path.stem
        json_path  = out_folder / f"{base_name}.json"
        md_path    = out_folder / f"{base_name}.md"
        txt_path   = out_folder / f"{base_name}.txt"
        clean_path = out_folder / f"{base_name}_clean.txt"

        is_retry = img_path in todo_retry
        marker = "[재시도]" if is_retry else "[신규]"
        print(f"{marker} [{i}/{len(todo_total)}] {rel_path}", flush=True)

        try:
            start_time = time.time()
            extracted = extract_text_from_image(img_path)
            elapsed = time.time() - start_time

            markdown_text = extracted['markdown']
            clean_text = extracted['clean_text']

            result_data = {
                "file_path": str(rel_path),
                "version_tag": VERSION_TAG,
                "vl_markdown": markdown_text,
                "vl_clean_text": clean_text,
                "vl_plain_text": clean_text,
                "processing_time_sec": round(elapsed, 1),
                "preprocessing": "none",
                "postprocessing": "html_cleanup",
                "gemini_corrected": ""
            }

            json_path.write_text(
                json.dumps(result_data, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )

            # ============================================================
            # 옵션 C 출력 분리:
            #   .md      = HTML 있는 원본 마크다운 (구조 보존)
            #   .txt     = HTML 있는 원본 (.md와 동일 내용, 호환용)
            #   _clean.txt = HTML 제거된 깔끔 텍스트 (연구자 본인 검수용)
            # ============================================================
            if markdown_text:
                md_path.write_text(markdown_text, encoding="utf-8")
                txt_path.write_text(markdown_text, encoding="utf-8")
            if clean_text:
                clean_path.write_text(clean_text, encoding="utf-8")

            real_count = count_real_chars(clean_text)
            md_len = len(markdown_text)
            if real_count == 0:
                print(f"   -> 경고: 인식된 텍스트 없음 ({elapsed:.1f}초)")
            else:
                print(f"   -> 완료: 본문 {real_count}자 / MD {md_len}자 ({elapsed:.1f}초)")

            processed += 1

        except Exception as e:
            print(f"   -> 오류: {str(e)[:120]}")
            error_data = {
                "file_path": str(rel_path),
                "version_tag": VERSION_TAG,
                "error": str(e),
                "last_attempt": datetime.now().isoformat()
            }
            json_path.write_text(
                json.dumps(error_data, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )

    if processed > 0:
        completed_ids = set()
        for f in OUTPUT_DIR.rglob('*.json'):
            if f.name.startswith('_'):
                continue
            try:
                data = json.loads(f.read_text(encoding='utf-8'))
                if 'error' not in data and ('vl_markdown' in data or 'vl_clean_text' in data):
                    completed_ids.add(f.stem)
            except Exception:
                pass
        write_version_xlsx(completed_ids)

    return processed


def run_daemon():
    print(f"[System] {VERSION_NAME}")
    print(f"[System] 상시 감시 시작 (주기: {POLL_INTERVAL}초)")
    print(f"[System] 감시 대상: {INPUT_DIR}")
    print(f"[System] 결과 저장: {OUTPUT_DIR}")
    print(f"[System] 종료: Ctrl + C\n" + "-" * 60)

    try:
        while True:
            count = process_pipeline()
            if count > 0:
                print("-" * 60)
                print(f"[System] 대기 중...")
            time.sleep(POLL_INTERVAL)
    except KeyboardInterrupt:
        print("\n[System] 사용자 요청으로 종료합니다.")


if __name__ == '__main__':
    run_daemon()