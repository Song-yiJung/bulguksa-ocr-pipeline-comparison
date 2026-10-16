import os
import json
import time
import re
import csv
import argparse
from pathlib import Path
from datetime import date, datetime
from dotenv import load_dotenv
import google.generativeai as genai
import PIL.Image
import openpyxl

# 파일 버전 식별자 (Sheet1 step2 컬럼에 기록됨)
SCRIPT_VERSION = "v3"

# 1. 환경 변수 및 API 인증
PROJECT_ROOT = Path(__file__).parent.parent
load_dotenv(dotenv_path=PROJECT_ROOT / ".env")

# GEMINI_API_KEYS (복수, 콤마구분) 우선, 없으면 GEMINI_API_KEY (단수) 폴백
_api_keys_raw = os.getenv("GEMINI_API_KEYS", "") or os.getenv("GEMINI_API_KEY", "")
GEMINI_API_KEYS = [k.strip() for k in _api_keys_raw.split(",") if k.strip()]

if not GEMINI_API_KEYS:
    print("[Error] .env 파일에 GEMINI_API_KEY(단수) 또는 GEMINI_API_KEYS(복수, 콤마구분)가 설정되지 않았습니다.")
    exit(1)

current_key_index = 0

# 2. CLI 인자 (환경변수 폴백 → 기본값=v3_VL_html)
parser = argparse.ArgumentParser(description="step2: Paddle 1차 결과 + Gemini 의미론적 교정")
parser.add_argument('--input-dir',  default=os.getenv('OCR_INPUT_DIR',
                    r'G:\jggmk2\OCR_PIPLINE_IMAGES'),
                    help='원본 이미지 디렉토리 (Gemini 입력용)')
parser.add_argument('--results-dir', default=os.getenv('OCR_RESULTS_DIR',
                    r'D:\OCR_PIPLINE\results\version_C_v3_VL_html'),
                    help='1차(step1) 처리 결과 폴더 = step2 입출력 디렉토리')
parser.add_argument('--field', default=None,
                    help='강제로 사용할 OCR 필드명 (지정 안 하면 자동 감지)')
parser.add_argument('--model', default='gemini-2.5-flash',
                    help='Gemini 모델명 (예: gemini-2.5-flash, gemini-2.5-pro)')
args = parser.parse_args()

GEMINI_MODEL_NAME = f'models/{args.model}'

# 키 개수에 비례해 대기 시간 자동 계산 (분당 10회 기준)
SLEEP_INTERVAL = max(2, 60 // (10 * len(GEMINI_API_KEYS)))

# 식민지기 사료엔 통치·하사·검열 등 정치 어휘가 자연스럽게 등장하므로
# 안전 필터를 모두 해제 (학술 OCR 교정 목적)
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT",        "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH",       "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
]


GENERATION_CONFIG = {
    "temperature": 0.0,  # 결정론적 출력 — 분포 꼬리 환각 차단 + 재현성 ↑
}


def get_model():
    global current_key_index
    current_key_index = (current_key_index + 1) % len(GEMINI_API_KEYS)
    genai.configure(api_key=GEMINI_API_KEYS[current_key_index])
    return genai.GenerativeModel(
        GEMINI_MODEL_NAME,
        safety_settings=SAFETY_SETTINGS,
        generation_config=GENERATION_CONFIG,
    )


def _extract_response_text(response):
    """response에서 텍스트와 finish_reason을 안전하게 추출. 실패 시 (None, reason)."""
    try:
        candidates = getattr(response, 'candidates', None) or []
        if not candidates:
            return None, "NO_CANDIDATES"
        cand = candidates[0]
        finish_reason = getattr(cand, 'finish_reason', None)
        finish_name = getattr(finish_reason, 'name', None) or str(finish_reason)
        # 정상 종료 (STOP=1) 또는 MAX_TOKENS(2) 등 — content가 있으면 사용
        content = getattr(cand, 'content', None)
        parts = getattr(content, 'parts', None) if content else None
        if parts:
            text = "".join(getattr(p, 'text', '') for p in parts).strip()
            if text:
                return text, finish_name
        return None, finish_name or "EMPTY_RESPONSE"
    except Exception as e:
        return None, f"PARSE_ERR:{type(e).__name__}:{str(e)[:80]}"

# 3. 물리적 입출력 볼륨 경로 (1단계와 아키텍처 공유)
INPUT_DIR  = Path(args.input_dir)
OUTPUT_DIR = Path(args.results_dir)  # 1차 결과 폴더 = step2 입출력
XLSX_PATH  = INPUT_DIR / 'OCR데이터목록.xlsx'

# OCR 필드 자동 감지 우선순위 (앞쪽이 우선)
OCR_FIELD_CANDIDATES = [
    "vl_clean_text",  # v3_VL_html: HTML 제거된 깔끔본 (Gemini 입력용 ★)
    "vl_markdown",    # v3_VL: 마크다운 (HTML 표 포함)
    "paddle_raw",     # v1_japan, v1_cht, v2_tuned: PP-OCR 원시 라인
    "vision_raw",     # 레거시 (Google Vision 결과)
]

def extract_ocr_text(data: dict, forced_field=None):
    """1차 OCR JSON에서 텍스트와 사용된 필드명을 반환."""
    if forced_field:
        return data.get(forced_field, "") or "", forced_field
    for field in OCR_FIELD_CANDIDATES:
        v = data.get(field)
        if v and isinstance(v, str) and v.strip():
            return v, field
    return "", ""

# --- 비용 로깅 설정 (장당 실측 가능) ---
# 공식 가격: https://ai.google.dev/pricing (요율 변동 시 수치만 조정)
GEMINI_PRICING_USD_PER_1M = {
    'models/gemini-2.5-flash':     {'input': 0.30, 'output': 2.50},
    'models/gemini-2.5-flash-lite':{'input': 0.10, 'output': 0.40},
    'models/gemini-2.5-pro':       {'input': 1.25, 'output': 10.00},
    'models/gemini-2.0-flash':     {'input': 0.10, 'output': 0.40},
    'models/gemini-2.0-flash-lite':{'input': 0.075,'output': 0.30},
}
USD_TO_KRW = 1400  # 환율 (청구서 환산용; 필요시 조정)
COST_LOG_PATH = OUTPUT_DIR / '_step2_cost_log.csv'

# 세션 누적 (프로세스 종료 시 요약 출력)
_session_stats = {
    'calls': 0, 'input_tokens': 0, 'output_tokens': 0,
    'cost_usd': 0.0, 'started_at': datetime.now(),
}

def _log_cost(file_id: str, model_name: str, prompt_tokens: int,
              output_tokens: int, duration_sec: float):
    """단일 Gemini 호출의 토큰·비용을 CSV 로깅하고 세션 누적치에 반영."""
    pricing = GEMINI_PRICING_USD_PER_1M.get(model_name, {'input': 0.0, 'output': 0.0})
    input_usd = prompt_tokens * pricing['input'] / 1_000_000
    output_usd = output_tokens * pricing['output'] / 1_000_000
    total_usd = input_usd + output_usd
    total_krw = total_usd * USD_TO_KRW

    # 세션 누적
    _session_stats['calls'] += 1
    _session_stats['input_tokens'] += prompt_tokens
    _session_stats['output_tokens'] += output_tokens
    _session_stats['cost_usd'] += total_usd

    # CSV 기록 (첫 쓰기 시 헤더 포함)
    is_new = not COST_LOG_PATH.exists()
    try:
        with open(COST_LOG_PATH, 'a', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            if is_new:
                w.writerow(['timestamp', 'file_id', 'model', 'input_tokens',
                            'output_tokens', 'total_tokens', 'input_cost_usd',
                            'output_cost_usd', 'total_cost_usd', 'total_cost_krw',
                            'duration_sec'])
            w.writerow([
                datetime.now().isoformat(timespec='seconds'), file_id, model_name,
                prompt_tokens, output_tokens, prompt_tokens + output_tokens,
                round(input_usd, 6), round(output_usd, 6), round(total_usd, 6),
                round(total_krw, 2), round(duration_sec, 2),
            ])
    except Exception as e:
        print(f"  [비용 로그 기록 실패(무시): {str(e)[:80]}]")

    return total_usd, total_krw

def _maybe_log_cost(response, file_id, model, call_duration):
    """response에서 usage_metadata가 있으면 비용 로깅, 결과 튜플 반환. 없으면 None."""
    try:
        usage = getattr(response, 'usage_metadata', None)
        if not usage:
            return None
        in_tok = getattr(usage, 'prompt_token_count', 0) or 0
        out_tok = getattr(usage, 'candidates_token_count', 0) or 0
        usd, krw = _log_cost(
            file_id=file_id,
            model_name=model.model_name if hasattr(model, 'model_name') else GEMINI_MODEL_NAME,
            prompt_tokens=in_tok,
            output_tokens=out_tok,
            duration_sec=call_duration,
        )
        return in_tok, out_tok, usd, krw
    except Exception as _cost_e:
        print(f"  [비용 로그 실패(무시): {str(_cost_e)[:120]}]")
        return None


def _print_session_summary():
    """프로세스 종료 시 세션 누적 비용 요약 출력."""
    s = _session_stats
    if s['calls'] == 0:
        return
    dur_min = (datetime.now() - s['started_at']).total_seconds() / 60
    cost_krw = s['cost_usd'] * USD_TO_KRW
    avg_usd = s['cost_usd'] / s['calls']
    avg_krw = cost_krw / s['calls']
    print("\n" + "=" * 60)
    print(f"[세션 비용 요약] {dur_min:.1f}분 동안 {s['calls']}호출")
    print(f"  입력 토큰:   {s['input_tokens']:>10,}")
    print(f"  출력 토큰:   {s['output_tokens']:>10,}")
    print(f"  총 토큰:     {s['input_tokens'] + s['output_tokens']:>10,}")
    print(f"  총 비용:     ${s['cost_usd']:.4f}  (≈₩{cost_krw:,.0f})")
    print(f"  장당 평균:   ${avg_usd:.4f}  (≈₩{avg_krw:.1f})")
    print(f"  상세 로그:   {COST_LOG_PATH}")
    print("=" * 60)


def _expand_range(s: str):
    m = re.match(r'(.+-)(\d+)~.+?(\d+)$', s)
    if m:
        p, start, end = m.group(1), int(m.group(2)), int(m.group(3))
        w = len(m.group(2))
        return {f'{p}{str(i).zfill(w)}' for i in range(start, end + 1)}
    return {s}


def update_sheet1_step2(completed_ids: set):
    """Sheet1 step2 컬럼에 '날짜 버전' 형식 기록. 기존 값은 보존."""
    if not XLSX_PATH.exists():
        return
    try:
        stamp = f"{date.today().isoformat()} {SCRIPT_VERSION}"
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb['Sheet1']
        changed = False
        for row in ws.iter_rows(min_row=2):
            files_cell = row[13]
            step2_cell = row[15]
            if not files_cell.value:
                continue
            if step2_cell.value:
                continue  # 이미 기록이 있으면 덮어쓰지 않음
            row_ids = set()
            for part in str(files_cell.value).split('|'):
                row_ids.update(_expand_range(part.strip()))
            if row_ids and row_ids.issubset(completed_ids):
                step2_cell.value = stamp
                changed = True
        if changed:
            wb.save(XLSX_PATH)
    except Exception as e:
        print(f"  [Sheet1 업데이트 실패(무시): {str(e)[:80]}]")

# 3. 도메인 및 매체 특성이 반영된 고문서 해독 특화 프롬프트
SYSTEM_PROMPT = """당신은 1930년대 일제강점기(1910~1945) 조선총독부 행정문서 부터 민간 서신, 전보에 이르는 문화유산 기록물 해독에 통달한 역사학자이자 아키비스트이다.
본 사료는 '불국사(佛國寺)' 및 관련 문화재를 대상으로 한 공문서, 서신, 전보 등으로 일본 한자(구자체), 히라가나, 가타카나, 숫자가 혼용되어 있으며, 인쇄체와 수기체(필기)가 섞여 있다.

첨부된 원본 이미지와 1차 Vision API가 물리적으로 추출한 원시 텍스트(vision_raw)를 교차 검증하여, 다음 학술적 지침에 따라 텍스트를 교정 및 복원한다.

[교정 및 복원 지침]

1. 표·목록 레이아웃 보존 (Tabular Layout Preservation) — 최우선 규칙:
   - 동일 열의 값이 반복될 때는 절대 평탄화하지 말고 반드시 **'〃' 또는 '全'** 기호로 표기한다.
   - 행정 계층(郡/面/里)은 각 행에서 동일 상위가 반복되면 '〃'로 축약하되, **상위 계층을 임의로 생략하거나 압축하지 말 것**.
     * 올바른 예:
       慶州郡 見谷面 羅原里          羅原里五層石塔
       〃     江西面 玉山里          淨惠寺址十三層石塔
       〃     內東面 馬洞里  佛國寺  大雄殿
     * 잘못된 예(평탄화): '慶州羅原里', '江西面玉山里'
   - 열 간 공백(spacing)은 이미지에 보이는 열 구분을 반영하여 탭/다수 공백으로 유지한다.

2. 일본어 법령체 접속사·조사 정밀 판독 (Classical Japanese Legal Grammar):
   - '又ハ'(또는)과 '及'(그리고)를 절대 혼동하지 말 것. 이미지 획 수와 가타카나 'ハ'의 유무로 판별.
     * '遺蹟又ハ景勝ノ地'(O) / '遺蹟及景勝地'(X)
     * '建造物ノ類'의 'ノ'와 'ハ' 분별 철저
   - 법령·공문서 상투어는 시대 용례를 반영한다:
     * '保存上必要ト認ムル事項'(구형) / '保存上必要ト認メラルル事項'(수동형)
     * '其他參考ト爲ルベキ事項'(관용구) / '其他參考ト為スヘキ事項'
     * 'ニ関スル', 'ニ付', 'ニ依ル'(ニヨル), '添附スル'
   - 조동사 활용('タル', 'ベキ', 'セシ', 'ラル')을 임의로 단순화하지 말 것.

3. 표준 조사서 템플릿 인식 (Fixed Survey Form Recognition):
   고적·보물 조사서는 다음 8개 고정 필드를 따른다. **번호를 중복하거나 건너뛰지 말 것**:
     一、名稱
     二、所在地
     三、地域、地番、地目及地積
     四、工作物其ノ他ノ物件ノ名稱、員數、品質、形狀、構造、形式及大サ
     五、現狀
     六、由來又ハ傳說
     七、保存上必要ト認メラルル事項
     八、其他參考ト爲ルベキ事項
   - 이미지에서 각 필드의 시작 번호(一~八)가 보이지 않아도 위 순서로 복원한다.
   - 소유자·관리자 등 중간 항목이 표 밖으로 튀어나오면 번호를 임의로 부여하지 말 것.

4. 숫자·도량형 정확성 (Numerical Fidelity):
   - 자릿수 구분 쉼표/점을 반드시 보존한다. 예: '三,三〇九坪'을 '三九坪'으로 축약 금지.
   - 치수 단위(尺·寸·間·段·坪)는 이미지 판독에 근거; 추측 금지.
   - 연호는 '昭和N年M月D日' 형식으로 풀어쓴다. 숫자가 불명확하면 '昭和□年'으로 표기.

5. 한자 변이자(구자체) 보존:
   - 현대 일본 상용한자(신자체)로 바꾸지 말 것. 원본의 구자체를 그대로 유지:
     * 臺(O) / 台(X), 龜(O) / 亀(X), 佛(O) / 仏(X), 國(O) / 国(X)
     * 關(O) / 関(X), 舊(O) / 旧(X), 藝(O) / 芸(X), 圖(O) / 図(X)
   - 유사 자형 오독 주의: 開/圓, 螭/龜, 奉/春, 栗/栢, 蒙/豪.

6. 고유명사·사찰명·미술사 용어 복원:
   - 인명: 長尾欽彌(나가오 킨야), 藤島亥治郞 등 당대 관련자명 철저 복원.
   - 지명: 慶州郡, 內東面 馬洞里, 陽北面 凡谷里 등 조선 행정구역 정확히.
   - 불교/미술사 용어: 阿彌陀如來, 毘盧舍那佛, 舍利塔, 石窟庵, 多寶塔, 釋迦塔, 白雲橋, 青雲橋, 七寶橋, 蓮華橋 등.
   - 사찰명은 실제 존재한 이름만 사용: 佛國寺, 石窟庵, 芬皇寺, 鳳停寺, 開心寺, 淨惠寺, 栢栗寺 등.

7. 환각(Hallucination) 방지 — 엄수:
   - 원본에 없는 내용을 생성하지 말 것.
   - 한 글자라도 이미지에서 확신할 수 없으면 **해당 글자 뒤에 '[?]'** 표기.
     예: '仁旺里[?]', '瞻星臺[?]'
   - 완전히 판독 불가능한 구간은 '□' 또는 '(판독불가)'로 표기.
   - vision_raw에 있으나 이미지에서 확인되지 않는 글자는 신뢰하지 말 것.
   - **★ 1차 OCR이 매우 짧을 때(30자 미만): 빈 페이지·표지·메모지 가능성이 높다.
     이미지에 罫線(괘선)·빈 결재란·印章 흔적만 보여도 그 양식을 추측으로 채우지 말 것.
     보이지 않는 사찰명·지명·인명·행정구역·결재란 항목을 임의로 생성 금지.
     1차 OCR 텍스트와 이미지에서 명백히 보이는 글자만 출력한다.**

8. 인장·직인·부기(附記) 처리:
   - 붉은 인장이 찍힌 위치에는 '(鈐印)' 또는 '(직인: 내용)' 형태로 표기.
     예: '(직인) 朝鮮總督府 9.5.6.'
   - 접수 스탬프·결재 도장의 날짜는 이미지에서 직접 판독.
   - 관인 테두리로 인한 무의미한 영문자(DES, DID 등)는 삭제.

9. 문서 레이아웃 복원:
   - 세로쓰기(우→좌) 원칙. vision_raw의 배열은 참고만 하고 이미지의 논리적 흐름에 따라 재구성.
   - 공문서 양식은 원본 서식 그대로 유지 (未完結/完結 種/施行月日番號/昭和  年  月  日接受/昭和  年  月  日起案/發送/記帳/鈐印/校合/淨寫/關係/番號/第  號/決裁 등).
   - 서신/전보는 발신인·수신인·본문 순서를 이미지 기준으로 복원.

10. 출력 형식 통제(Output Control):
    - 부가적인 설명, 주석, 마크다운 코드 블록을 절대 포함하지 말고 교정된 순수 텍스트만 반환.
    - 연구자의 한글 주석(예: '공착물', '지')이 vision_raw에 섞여 있어도 원문 한자로 환원.
    - 빈 줄은 원본 문서의 단락 구분을 따라 유지.
"""

def process_gemini_correction():
    """1차 처리된 JSON 산출물을 기반으로 2차 의미론적 교정 파이프라인 구동"""
    if not OUTPUT_DIR.exists():
        print(f"[System] 1차 처리된 결과물 디렉토리가 존재하지 않습니다. ({OUTPUT_DIR})")
        return

    json_files = sorted(OUTPUT_DIR.rglob('*.json'))
    json_files = [f for f in json_files
                  if not f.name.startswith('_')
                  and not any(pt.startswith('_') for pt in f.relative_to(OUTPUT_DIR).parts)]

    processed_count = 0
    print(f"[System] 키 {len(GEMINI_API_KEYS)}개 사용 / 호출 간격: {SLEEP_INTERVAL}초")

    for i, json_path in enumerate(json_files, 1):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # 멱등성 유지: 교정 필드 또는 txt 파일이 존재하면 스킵
            txt_path = json_path.with_name(f"{json_path.stem}_gemini_final.txt")
            if data.get("gemini_corrected", "").strip():
                continue
            if txt_path.exists():
                # txt는 있지만 JSON이 비어있는 경우 JSON만 동기화하고 스킵
                data["gemini_corrected"] = txt_path.read_text(encoding="utf-8")
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                continue

            rel_file_path = data.get("file_path")
            ocr_text, used_field = extract_ocr_text(data, args.field)
            ocr_len = len(ocr_text.strip())
            img_path = INPUT_DIR / rel_file_path

            if not img_path.exists():
                print(f"[Error] 참조할 원본 이미지가 누락되었습니다: {img_path}")
                continue

            if not ocr_text:
                fields_present = [k for k in data.keys() if not k.startswith('_')]
                print(f"  -> [Skip] OCR 텍스트 필드 없음 ({json_path.name}, 후보: {OCR_FIELD_CANDIDATES}, 실제: {fields_present})")
                continue

            # ============================================================
            # 환각 가드 (c): 1차 OCR 길이 기반 분기
            #   < 10자 : Gemini 호출 스킵, 1차 OCR 결과 그대로 사용 (빈 페이지 환각 차단)
            #   10~50자: Gemini 호출하되 프롬프트에 "양식 추측 채우지 말 것" 경고 추가
            #   > 50자 : 일반 처리
            # ============================================================
            if ocr_len < 10:
                data["gemini_corrected"] = ocr_text
                data["gemini_skip_reason"] = f"too_short ({ocr_len}자, threshold=10)"
                data.pop("gemini_error", None)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                txt_path.write_text(ocr_text, encoding="utf-8")
                print(f"[{i}/{len(json_files)}] {rel_file_path} -> [Skip] OCR {ocr_len}자 < 10 (빈페이지 환각가드, 1차 결과 그대로)")
                processed_count += 1
                continue

            short_note = ""
            if ocr_len < 50:
                short_note = (
                    f"\n\n[중요: 1차 OCR이 매우 짧음({ocr_len}자) — 빈 페이지·표지·메모지 가능성. "
                    f"이미지에 罫線/빈 결재란/印章 흔적만 보여도 양식을 추측으로 채우지 말 것. "
                    f"보이는 글자만 출력.]"
                )

            print(f"[{i}/{len(json_files)}] 2차 교정 처리 중: {rel_file_path} [field={used_field}, len={ocr_len}]", flush=True)

            img = PIL.Image.open(img_path)
            combined_prompt = f"{SYSTEM_PROMPT}\n\n[1차 원시 텍스트 ({used_field}, {ocr_len}자)]\n{ocr_text}{short_note}"

            # 키 로테이션으로 모델 호출
            model = get_model()
            call_start = time.time()
            response = model.generate_content([combined_prompt, img])
            call_duration = time.time() - call_start

            corrected_text, finish_reason = _extract_response_text(response)
            if corrected_text is None:
                # 응답 거부/공백 — 사유를 JSON에 기록하여 추적 가능하게
                data["gemini_error"] = f"{finish_reason} ({call_duration:.1f}s)"
                data["gemini_corrected"] = ""
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"  -> [Skip] 응답 없음 | finish_reason={finish_reason} | {call_duration:.1f}s")
                # 비용 로그는 토큰 메타데이터가 있으면 그래도 기록 (호출은 발생함)
                _maybe_log_cost(response, json_path.stem, model, call_duration)
                time.sleep(SLEEP_INTERVAL)
                continue

            # 정상 처리 — 이전에 실패로 기록된 에러 필드는 제거
            data.pop("gemini_error", None)
            data["gemini_corrected"] = corrected_text
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            # 분석용 단일 TXT 파일 생성
            txt_path.write_text(corrected_text, encoding="utf-8")

            # 토큰·비용 로깅
            usd_krw = _maybe_log_cost(response, json_path.stem, model, call_duration)
            if usd_krw:
                in_tok, out_tok, usd, krw = usd_krw
                print(f"  -> 교정 완료 | {in_tok}+{out_tok} 토큰 | ${usd:.4f} (≈₩{krw:.1f}) | {call_duration:.1f}s")
            else:
                print(f"  -> 교정 완료 (토큰 메타데이터 없음) | {call_duration:.1f}s")

            processed_count += 1

            time.sleep(SLEEP_INTERVAL)

        except Exception as e:
            err_msg = str(e)
            if '429' in err_msg:
                print(f"  -> [429] 할당량 초과. 60초 대기 후 재시도...")
                time.sleep(60)
            else:
                # 에러 사유를 JSON에도 기록하여 다음 실행에서 재시도 가능
                try:
                    data["gemini_error"] = f"EXCEPTION:{type(e).__name__}:{err_msg[:300]}"
                    data.setdefault("gemini_corrected", "")
                    with open(json_path, 'w', encoding='utf-8') as _f:
                        json.dump(data, _f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
                print(f"  -> [Error] 처리 실패 ({type(e).__name__}): {err_msg[:300]}")

    # Sheet1 step2 컬럼 일괄 업데이트 (gemini_corrected 채워진 파일 기준)
    completed_ids = set()
    for f in OUTPUT_DIR.rglob('*.json'):
        if f.name.startswith('_'):
            continue
        try:
            d = json.loads(f.read_text(encoding='utf-8'))
            if d.get('gemini_corrected', '').strip():
                completed_ids.add(f.stem)
        except Exception:
            pass
    update_sheet1_step2(completed_ids)

    if processed_count == 0:
        print("[System] 현재 2차 교정을 수행할 신규 데이터가 없습니다.")
    else:
        print(f"[System] 총 {processed_count}건의 2차 교정 작업이 완료되었습니다.")


VALID_IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff'}


def report_missing_step2():
    """step2 누락/완료 현황 리포트"""
    input_rels = {p.relative_to(INPUT_DIR) for p in INPUT_DIR.rglob('*')
                  if p.suffix.lower() in VALID_IMAGE_EXTS
                  and not any(pt.startswith('_') for pt in p.relative_to(INPUT_DIR).parts)}
    done = set()
    for f in OUTPUT_DIR.rglob('*.json'):
        if f.name.startswith('_'):
            continue
        try:
            d = json.loads(f.read_text(encoding='utf-8'))
            if d.get('gemini_corrected', '').strip():
                # JSON 파일과 같은 위치/이름의 이미지를 매칭 (확장자 무관)
                rel_no_ext = f.relative_to(OUTPUT_DIR).with_suffix('')
                for ext in VALID_IMAGE_EXTS:
                    candidate = rel_no_ext.with_suffix(ext)
                    if candidate in input_rels:
                        done.add(candidate)
                        break
        except Exception:
            pass
    missing = input_rels - done
    print(f"[현황] INPUT: {len(input_rels)} / step2 완료: {len(done)} / 누락: {len(missing)}")
    from collections import defaultdict
    by_kwon = defaultdict(int)
    for r in missing:
        by_kwon[r.parts[0]] += 1
    for k in sorted(by_kwon.keys()):
        print(f"  {k}: {by_kwon[k]}개 미완료")


if __name__ == '__main__':
    import atexit
    atexit.register(_print_session_summary)
    print("=" * 60)
    print(f"[Config] step2_gemini SCRIPT_VERSION={SCRIPT_VERSION}")
    print(f"[Config] 이미지 디렉토리:  {INPUT_DIR}")
    print(f"[Config] 1차 결과 디렉토리: {OUTPUT_DIR}")
    print(f"[Config] Gemini 모델:     {GEMINI_MODEL_NAME}")
    print(f"[Config] API 키 개수:     {len(GEMINI_API_KEYS)} (호출 간격 {SLEEP_INTERVAL}s)")
    print(f"[Config] OCR 필드 우선순위: {args.field or OCR_FIELD_CANDIDATES}")
    print("=" * 60)
    report_missing_step2()
    print("-" * 60)
    process_gemini_correction()
    _print_session_summary()

    # 비용분석 시트 자동 갱신 (CSV 로그 기반)
    try:
        print("\n[System] 비용분석 시트 갱신 중...")
        import sys
        _here = os.path.dirname(os.path.abspath(__file__))
        if _here not in sys.path:
            sys.path.insert(0, _here)
        from update_cost_sheet import main as _update_cost_sheet
        _update_cost_sheet()
    except Exception as _e:
        print(f"[System] 비용분석 시트 갱신 건너뜀 (무시): {str(_e)[:100]}")
