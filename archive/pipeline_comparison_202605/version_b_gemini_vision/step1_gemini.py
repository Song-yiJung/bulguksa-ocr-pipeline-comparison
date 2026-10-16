"""version_B/D step1 — Gemini가 이미지에서 직접 텍스트 추출 (LLM-first OCR).

A/C의 (OCR → Gemini 교정) 순서를 뒤집은 hybrid 파이프라인의 1단계.
1차 OCR 텍스트 없이 시스템 프롬프트 + 이미지만 Gemini에 입력.
"""
import os
import json
import time
import sys
import argparse
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import google.generativeai as genai
import PIL.Image

PROJECT_ROOT = Path(__file__).parent.parent
load_dotenv(dotenv_path=PROJECT_ROOT / ".env")

# ============================================================
# 버전 메타데이터
# ============================================================
VERSION_TAG = "v_B_step1_gemini"
VERSION_NAME = "Version B step1 — Gemini 멀티모달 OCR (1차 텍스트 없이 이미지만)"

# ============================================================
# API 키 (단/복수 폴백)
# ============================================================
_api_keys_raw = os.getenv("GEMINI_API_KEYS", "") or os.getenv("GEMINI_API_KEY", "")
GEMINI_API_KEYS = [k.strip() for k in _api_keys_raw.split(",") if k.strip()]
if not GEMINI_API_KEYS:
    print("[Error] .env에 GEMINI_API_KEY(단수) 또는 GEMINI_API_KEYS(복수)가 필요")
    exit(1)

# ============================================================
# CLI 인자
# ============================================================
parser = argparse.ArgumentParser(description="step1: Gemini가 이미지에서 직접 OCR")
parser.add_argument('--input-dir', default=os.getenv('OCR_INPUT_DIR',
                    r'G:\jggmk2\OCR_PIPLINE_IMAGES'))
parser.add_argument('--output-dir', required=True,
                    help='step1 결과 저장 폴더 (예: D:\\OCR_PIPLINE\\results\\version_B_step1)')
parser.add_argument('--model', default='gemini-2.5-flash')
parser.add_argument('--once', action='store_true', help='daemon 없이 한 번만 실행')
args = parser.parse_args()

INPUT_DIR = Path(args.input_dir)
OUTPUT_DIR = Path(args.output_dir)
GEMINI_MODEL_NAME = f'models/{args.model}'
POLL_INTERVAL = 60

# ============================================================
# SYSTEM_PROMPT — step2_gemini.py에서 import (코드 중복 방지)
# ============================================================
_step2_path = PROJECT_ROOT / "version_C_paddle_standard"
sys.path.insert(0, str(_step2_path))
# step2_gemini.py가 모듈 레벨 argparse를 쓰므로 sys.argv 임시 격리
_argv_backup = sys.argv[:]
sys.argv = [sys.argv[0]]
try:
    from step2_gemini import SYSTEM_PROMPT as _BASE_SYSTEM_PROMPT
except Exception as e:
    print(f"[Warning] SYSTEM_PROMPT import 실패: {e}")
    _BASE_SYSTEM_PROMPT = ("당신은 1930년대 일제강점기 조선총독부 행정문서 해독에 통달한 "
                          "역사학자입니다. 이미지에서 보이는 글자만 정확히 추출하세요.")
sys.argv = _argv_backup

# step1은 1차 OCR 텍스트가 없으므로 프롬프트 끝에 한 줄 추가
SYSTEM_PROMPT = (
    _BASE_SYSTEM_PROMPT
    + "\n\n[중요: 이 작업은 1차 OCR 결과 없이 이미지에서 직접 텍스트를 추출합니다. "
    "이미지에서 명백히 보이는 글자만 출력하고, 보이지 않는 양식/사찰명/인명을 추측 생성 금지.]"
)

# ============================================================
# 안전 필터 풀기 + 모델
# ============================================================
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT",        "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH",       "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
]

current_key_index = 0
SLEEP_INTERVAL = max(2, 60 // (10 * len(GEMINI_API_KEYS)))

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
    try:
        candidates = getattr(response, 'candidates', None) or []
        if not candidates:
            return None, "NO_CANDIDATES"
        cand = candidates[0]
        finish_reason = getattr(cand, 'finish_reason', None)
        finish_name = getattr(finish_reason, 'name', None) or str(finish_reason)
        content = getattr(cand, 'content', None)
        parts = getattr(content, 'parts', None) if content else None
        if parts:
            text = "".join(getattr(p, 'text', '') for p in parts).strip()
            if text:
                return text, finish_name
        return None, finish_name or "EMPTY_RESPONSE"
    except Exception as e:
        return None, f"PARSE_ERR:{type(e).__name__}"


# ============================================================
# 메인
# ============================================================
VALID_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff'}


def process_pipeline():
    if not INPUT_DIR.exists():
        print(f"[Error] 입력 폴더 없음: {INPUT_DIR}")
        return 0

    image_files = sorted([
        p for p in INPUT_DIR.rglob('*')
        if p.suffix.lower() in VALID_EXTS
        and not any(pt.startswith('_') for pt in p.relative_to(INPUT_DIR).parts)
    ])

    todo = []
    for img_path in image_files:
        rel = img_path.relative_to(INPUT_DIR)
        json_path = OUTPUT_DIR / rel.parent / f"{img_path.stem}.json"
        if json_path.exists():
            try:
                d = json.loads(json_path.read_text(encoding='utf-8'))
                if d.get('gemini_raw', '').strip():
                    continue  # 이미 완료
            except Exception:
                pass
        todo.append(img_path)

    if not todo:
        return 0

    print(f"[step1] 신규/재시도 대상: {len(todo)}장")
    processed = 0

    for i, img_path in enumerate(todo, 1):
        rel = img_path.relative_to(INPUT_DIR)
        out_folder = OUTPUT_DIR / rel.parent
        out_folder.mkdir(parents=True, exist_ok=True)

        base_name = img_path.stem
        json_path = out_folder / f"{base_name}.json"
        txt_path = out_folder / f"{base_name}.txt"

        print(f"[{i}/{len(todo)}] {rel}", flush=True)

        try:
            img = PIL.Image.open(img_path)
            model = get_model()
            start = time.time()
            response = model.generate_content([SYSTEM_PROMPT, img])
            elapsed = time.time() - start

            text, finish = _extract_response_text(response)
            usage = getattr(response, 'usage_metadata', None)
            in_tok = getattr(usage, 'prompt_token_count', 0) or 0 if usage else 0
            out_tok = getattr(usage, 'candidates_token_count', 0) or 0 if usage else 0

            if text is None:
                result = {
                    "file_path": str(rel),
                    "version_tag": VERSION_TAG,
                    "gemini_raw": "",
                    "gemini_error": finish,
                    "processing_time_sec": round(elapsed, 1),
                    "prompt_tokens": in_tok,
                    "output_tokens": out_tok,
                }
                print(f"  -> [Skip] finish_reason={finish} / {elapsed:.1f}s")
            else:
                result = {
                    "file_path": str(rel),
                    "version_tag": VERSION_TAG,
                    "gemini_raw": text,
                    "processing_time_sec": round(elapsed, 1),
                    "prompt_tokens": in_tok,
                    "output_tokens": out_tok,
                }
                txt_path.write_text(text, encoding="utf-8")
                print(f"  -> {len(text)}자 / {in_tok}+{out_tok} 토큰 / {elapsed:.1f}s")

            json_path.write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            processed += 1
            time.sleep(SLEEP_INTERVAL)

        except Exception as e:
            err_msg = str(e)
            if '429' in err_msg:
                print(f"  -> [429] 60초 대기 후 재시도...")
                time.sleep(60)
            else:
                error_data = {
                    "file_path": str(rel),
                    "version_tag": VERSION_TAG,
                    "error": f"{type(e).__name__}: {err_msg[:300]}",
                    "last_attempt": datetime.now().isoformat()
                }
                json_path.write_text(
                    json.dumps(error_data, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )
                print(f"  -> [Error] {type(e).__name__}: {err_msg[:200]}")

    return processed


def run_daemon():
    print("=" * 60)
    print(f"[System] {VERSION_NAME}")
    print(f"[Config] 입력: {INPUT_DIR}")
    print(f"[Config] 출력: {OUTPUT_DIR}")
    print(f"[Config] 모델: {GEMINI_MODEL_NAME}")
    print(f"[Config] 키 {len(GEMINI_API_KEYS)}개 / 호출 간격 {SLEEP_INTERVAL}s")
    print("=" * 60)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if args.once:
        process_pipeline()
        return

    print(f"[System] daemon 모드 (폴링 {POLL_INTERVAL}s, Ctrl+C 종료)")
    print("-" * 60)
    try:
        while True:
            count = process_pipeline()
            if count > 0:
                print(f"[System] {count}건 처리, 대기 중...")
            time.sleep(POLL_INTERVAL)
    except KeyboardInterrupt:
        print("\n[System] 종료")


if __name__ == '__main__':
    run_daemon()
