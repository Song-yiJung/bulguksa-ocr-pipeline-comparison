"""2단계 비교 실험: 지시문 3판(v1·v2·v3) × 모델 2종 × 3회 재독해.

논문 Ⅴ장의 실험을 실행한 코드(run_experiment.py)의 공개본이다.
모델 호출과 입력 구성은 원래 코드와 같고, 로컬 경로와 비용 추정용 내부 기록만 덜어 냈다.

호출 조건
  - 입력: f"{지시문}\\n\\n[1차 원시 텍스트 (vision_raw)]\\n{vision_raw}" 문자열 + 원본 이미지
  - gemini-2.5-flash: google-generativeai(구형 SDK). temperature·top_p·top_k·안전 설정 지정 안 함
  - gemini-3.8-flash: google-genai(신형 SDK). temperature 등은 지정 안 함.
      사고 수준 thinking_level=MEDIUM(이 모델의 기본값을 명시), 이미지 해상도 MEDIA_RESOLUTION_HIGH
  - 각 조합을 3회 실행. 호출 순서는 SEED로 섞는다. 이미 결과가 있는 호출은 건너뛴다.

지시문 이름(공개명 ↔ 실험 당시 내부명)
  v1 = v1  (초기 지시문)
  v2 = v3  (수정 지시문, 이후 실제 판독에 사용하는 확정본)
  v3 = v4b (상세화 지시문)
지시문 파일은 실험에 쓴 원문 그대로이며, 실행 전에 MD5로 원문과 같은지 확인한다.

사용 예:
  python run_experiment.py --sample 표본목록.csv --image-dir 이미지폴더 --vision-dir 실험데이터/비전결과 --dry-run
  python run_experiment.py --sample 표본목록.csv --image-dir 이미지폴더 --vision-dir 실험데이터/비전결과 --model gemini-2.5-flash
  python run_experiment.py --sample 표본목록.csv --image-dir 이미지폴더 --vision-dir 실험데이터/비전결과 --model gemini-3.8-flash

표본목록(CSV 또는 XLSX 시트 '표본목록') 열: 파일id, 상대경로(--image-dir 기준), 분석집합('본 분석' 또는 '짧은 면')
출력: <out-dir>/<모델>/<지시문>/run<n>/<파일id>.txt, 호출 기록 <out-dir>/<모델>/_호출기록.csv
API 키: .env 의 GEMINI_API_KEYS (쉼표로 여러 개를 적으면 첫 번째를 쓴다)
"""
import argparse
import csv
import hashlib
import json
import os
import random
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import PIL.Image
from dotenv import load_dotenv

HERE = Path(__file__).parent
PROMPT_DIR = HERE / 'prompts'

# 공개명 → (파일명, 실험 당시 내부명, 원문 MD5 앞 8자리)
# MD5는 파일을 UTF-8 텍스트로 읽은 값(줄바꿈은 읽을 때 LF로 정규화됨) 기준이다.
PROMPTS = {
    'v1': ('v1_initial.txt',  'v1',  '6bf22ffd'),
    'v2': ('v2_revised.txt',  'v3',  '57e7db87'),
    'v3': ('v3_detailed.txt', 'v4b', '938632bc'),
}

DEFAULT_MODEL = 'gemini-2.5-flash'
MODELS = ['gemini-2.5-flash', 'gemini-3.8-flash']
SEED = 20260925

# 모델별 단가 (USD / 1M 토큰, 사고 토큰은 출력 단가). ai.google.dev/pricing, 2026-10-06 확인
PRICES = {'gemini-2.5-flash': (0.30, 2.50), 'gemini-3.8-flash': (0.75, 3.75)}
USD_TO_KRW = 1400

RETRY_WAITS = [10, 30, 60, 120]   # 429·5xx·시간 초과 시 재시도 대기(초)

# Gemini 3.x 설정
THINKING_LEVEL = 'MEDIUM'                    # 3.8-flash 기본값(medium)을 명시
MEDIA_RESOLUTION = 'MEDIA_RESOLUTION_HIGH'   # 이미지당 약 1,120토큰

_log_lock = threading.Lock()
_print_lock = threading.Lock()
_stop = threading.Event()   # 월 지출 한도 도달 시 남은 호출을 건너뛴다
_client = None              # google-genai 클라이언트


def load_prompts(versions):
    prompts = {}
    for v in versions:
        fname, internal, expected = PROMPTS[v]
        text = (PROMPT_DIR / fname).read_text(encoding='utf-8')
        md5 = hashlib.md5(text.encode('utf-8')).hexdigest()[:8]
        if md5 != expected:
            raise SystemExit(f'[중단] {v}({internal}) 지시문이 실험 원문과 다릅니다 (MD5 {md5} ≠ {expected})')
        prompts[v] = text
    return prompts


def load_sample(path: Path, image_dir: Path, subset):
    if path.suffix.lower() == '.csv':
        rows = list(csv.DictReader(open(path, encoding='utf-8-sig')))
    else:
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True)['표본목록']
        it = ws.iter_rows(values_only=True)
        header = list(next(it))
        rows = [dict(zip(header, r)) for r in it if r and r[0]]
    pages = []
    for r in rows:
        is_main = r['분석집합'] == '본 분석'
        if (subset == 'main' and not is_main) or (subset == 'short' and is_main):
            continue
        img = Path(r['상대경로'])
        pages.append({'fid': r['파일id'], 'image': img if img.is_absolute() else image_dir / img})
    return pages


def call_gemini25(model, prompt_text, vision_raw, image_path):
    """gemini-2.5-flash: 운영 코드와 같은 구형 SDK(google.generativeai) 호출."""
    combined = f"{prompt_text}\n\n[1차 원시 텍스트 (vision_raw)]\n{vision_raw}"
    with PIL.Image.open(image_path) as img:
        img.load()
        return model.generate_content([combined, img])


def call_gemini3(model_name, prompt_text, vision_raw, image_path):
    """gemini-3.x: 입력 구성은 2.5와 같다. 사고 수준과 이미지 해상도만 명시한다."""
    from google.genai import types
    combined = f"{prompt_text}\n\n[1차 원시 텍스트 (vision_raw)]\n{vision_raw}"
    image = types.Part.from_bytes(data=Path(image_path).read_bytes(), mime_type='image/jpeg')
    image.media_resolution = types.PartMediaResolution(level=MEDIA_RESOLUTION)
    config = types.GenerateContentConfig(thinking_config=types.ThinkingConfig(thinking_level=THINKING_LEVEL))
    resp = _client.models.generate_content(model=model_name, contents=[combined, image], config=config)
    if not (resp.text or '').strip():
        reason = resp.candidates[0].finish_reason if resp.candidates else 'NO_CANDIDATES'
        raise ValueError(f'빈 응답 (finish_reason={reason})')
    return resp


def run_task(task, prompts, cfg, total, counter):
    version, run, page = task['version'], task['run'], task['page']
    fid = page['fid']
    if _stop.is_set():
        return 'skipped', 0.0
    target = cfg['out'] / version / f'run{run}' / f'{fid}.txt'
    vision_raw = json.loads((cfg['vision_dir'] / f'{fid}.json').read_text(encoding='utf-8')).get('vision_raw', '')

    if cfg['gemini3']:
        model = None
    else:
        import google.generativeai as genai
        model = genai.GenerativeModel(cfg['model_name'])

    attempts, note, resp, text = 0, '', None, None
    t0 = time.time()
    for wait in [0] + RETRY_WAITS:
        if wait:
            time.sleep(wait)
        attempts += 1
        try:
            if cfg['gemini3']:
                resp = call_gemini3(cfg['model_name'], prompts[version], vision_raw, page['image'])
            else:
                resp = call_gemini25(model, prompts[version], vision_raw, page['image'])
            text = resp.text.strip()
            break
        except Exception as e:
            note = f'{type(e).__name__}: {str(e)[:200]}'
            resp, text = None, None
            if 'spending cap' in str(e):
                _stop.set()
                break
            if not any(s in str(e) for s in ('429', '500', '503', 'Deadline', 'timed out', 'UNAVAILABLE')):
                break
    dur = time.time() - t0

    usage = getattr(resp, 'usage_metadata', None) if resp is not None else None
    in_tok = (getattr(usage, 'prompt_token_count', 0) or 0) if usage else 0
    out_tok = (getattr(usage, 'candidates_token_count', 0) or 0) if usage else 0
    total_tok = (getattr(usage, 'total_token_count', 0) or 0) if usage else 0
    think_tok = getattr(usage, 'thoughts_token_count', None) if usage else None
    if think_tok is None:   # 구형 SDK에는 사고 토큰 항목이 없어 차이로 구한다
        think_tok = max(0, total_tok - in_tok - out_tok)
    p_in, p_out = PRICES[cfg['model']]
    cost = (in_tok * p_in + (out_tok + think_tok) * p_out) / 1e6 * USD_TO_KRW

    if text is not None:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding='utf-8')
        status = 'ok'
    else:
        status = 'error'   # 결과 파일을 만들지 않으므로 다시 실행하면 이 호출만 재시도된다

    with _log_lock:
        log = cfg['out'] / '_호출기록.csv'
        is_new = not log.exists()
        with open(log, 'a', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            if is_new:
                w.writerow(['timestamp', 'version', 'run', 'file_id', 'status', 'input_tokens',
                            'output_tokens', 'thinking_tokens', 'cost_krw', 'duration_sec', 'attempts', 'note'])
            w.writerow([datetime.now().isoformat(timespec='seconds'), version, run, fid, status,
                        in_tok, out_tok, think_tok, round(cost, 2), round(dur, 1), attempts, note])
    with _print_lock:
        counter[0] += 1
        mark = '완료' if status == 'ok' else f'오류 ({note[:60]})'
        print(f'[{counter[0]}/{total}] {version} run{run} {fid}  {mark}  ₩{cost:.1f}  {dur:.1f}s', flush=True)
    return status, cost


def main():
    global _client
    ap = argparse.ArgumentParser(description='지시문 3판 비교 실험 (2단계 Gemini 재독해)')
    ap.add_argument('--sample', required=True, type=Path, help='표본목록 CSV/XLSX')
    ap.add_argument('--image-dir', required=True, type=Path, help='원본 이미지 폴더')
    ap.add_argument('--vision-dir', required=True, type=Path, help='1단계 Vision 결과 JSON 폴더')
    ap.add_argument('--out-dir', type=Path, default=HERE / '실험데이터' / '출력')
    ap.add_argument('--model', default=DEFAULT_MODEL, choices=MODELS)
    ap.add_argument('--versions', nargs='+', default=list(PROMPTS), choices=list(PROMPTS))
    ap.add_argument('--runs', type=int, default=3)
    ap.add_argument('--subset', default='all', choices=['all', 'main', 'short'],
                    help='all=본 분석+짧은 면, main=본 분석만, short=기준 판독문 20자 미만 면만')
    ap.add_argument('--limit', type=int, default=None, help='앞에서부터 N면만 (시험 실행용)')
    ap.add_argument('--workers', type=int, default=5)
    ap.add_argument('--dry-run', action='store_true', help='호출하지 않고 대상 수만 출력')
    args = ap.parse_args()

    cfg = {'model': args.model, 'model_name': f'models/{args.model}',
           'gemini3': not args.model.startswith('gemini-2.'),
           'out': args.out_dir / args.model, 'vision_dir': args.vision_dir}

    prompts = load_prompts(args.versions)
    pages = load_sample(args.sample, args.image_dir, args.subset)
    if args.limit:
        pages = pages[:args.limit]
    missing = [p['fid'] for p in pages
               if not p['image'].exists() or not (args.vision_dir / f"{p['fid']}.json").exists()]
    if missing:
        raise SystemExit(f'[중단] 이미지 또는 Vision 결과가 없는 면 {len(missing)}건: {missing[:5]}')

    tasks = [{'version': v, 'run': r, 'page': p}
             for p in pages for v in args.versions for r in range(1, args.runs + 1)]
    random.Random(SEED).shuffle(tasks)
    todo = [t for t in tasks
            if not (cfg['out'] / t['version'] / f"run{t['run']}" / f"{t['page']['fid']}.txt").exists()]

    print('=' * 70)
    print(f"모델 {cfg['model_name']} / 지시문 " + ', '.join(f'{v}({PROMPTS[v][1]})' for v in args.versions))
    print(f'대상 {len(pages)}면 × {len(args.versions)}판 × {args.runs}회 = {len(tasks)}회, 남은 호출 {len(todo)}회')
    print('=' * 70)
    if args.dry_run or not todo:
        return

    load_dotenv(HERE / '.env')
    keys = [k.strip() for k in os.getenv('GEMINI_API_KEYS', '').split(',') if k.strip()]
    if not keys:
        raise SystemExit('[중단] .env 에 GEMINI_API_KEYS 가 없습니다.')
    cfg['out'].mkdir(parents=True, exist_ok=True)
    if cfg['gemini3']:
        import google.genai as genai3
        _client = genai3.Client(api_key=keys[0])
        sdk = f'google-genai {genai3.__version__}'
        extra = {'thinking_level': THINKING_LEVEL, 'media_resolution(이미지별)': MEDIA_RESOLUTION}
    else:
        import google.generativeai as genai
        genai.configure(api_key=keys[0])
        sdk = f'google-generativeai {genai.__version__}'
        extra = {}
    settings = {'model': cfg['model_name'], 'sdk': sdk, **extra,
                'temperature·top_p·top_k·안전 설정': '지정 안 함',
                '지시문': {v: {'파일': PROMPTS[v][0], '내부명': PROMPTS[v][1], 'MD5': PROMPTS[v][2]} for v in args.versions},
                'runs': args.runs, 'seed': SEED, '기록일': datetime.now().isoformat(timespec='seconds')}
    (cfg['out'] / '_실행설정.json').write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')

    counter, total_cost = [0], 0.0
    count = {'ok': 0, 'error': 0, 'skipped': 0}
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = [ex.submit(run_task, t, prompts, cfg, len(todo), counter) for t in todo]
        for f in as_completed(futures):
            status, cost = f.result()
            total_cost += cost
            count[status] += 1
    print(f"완료 {count['ok']} / 오류 {count['error']} / 건너뜀 {count['skipped']} / 비용 약 ₩{total_cost:,.0f}")
    if count['error'] or count['skipped']:
        print('결과 파일이 없는 호출은 같은 명령을 다시 실행하면 그 호출만 다시 시도한다.')


if __name__ == '__main__':
    main()
