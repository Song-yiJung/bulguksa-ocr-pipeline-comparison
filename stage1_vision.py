"""1단계: Google Cloud Vision 문자 인식.

실험에 쓴 1단계 결과는 2026-04-20에 아래와 같은 호출로 한 번 산출한 것이며,
이후 모든 2단계 실험에 같은 결과를 입력하였다.

  - document_text_detection (DOCUMENT_TEXT_DETECTION)
  - language_hints = ["ja", "zh-Hant"]

원래 운영 코드(step1_vision_20260420.py)는 폴더를 상시 감시하며 작업 목록 엑셀을 갱신하는
기능을 함께 가졌다. 공개본에서는 그 부분을 덜어 내고, Vision 호출과 저장 형식은 그대로 두었다.

사용 예:
  python stage1_vision.py --image-dir 이미지폴더 --out-dir 실험데이터/비전결과

출력: <out-dir>/<파일id>.json  {"file_path": ..., "vision_raw": ..., "gemini_corrected": ""}
인증: 환경 변수 GOOGLE_APPLICATION_CREDENTIALS 에 서비스 계정 키 파일 경로를 지정한다(.env 가능).
"""
import argparse
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from google.cloud import vision

LANGUAGE_HINTS = ["ja", "zh-Hant"]
IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff'}


def extract_text_from_image(client, img_path: Path) -> str:
    """운영 코드와 같은 호출."""
    image = vision.Image(content=img_path.read_bytes())
    image_context = vision.ImageContext(language_hints=LANGUAGE_HINTS)
    response = client.document_text_detection(image=image, image_context=image_context)
    if response.error.message:
        raise Exception(f"Vision API 통신 오류: {response.error.message}")
    if response.full_text_annotation:
        return response.full_text_annotation.text.strip()
    return ""


def main():
    ap = argparse.ArgumentParser(description='1단계: Google Cloud Vision 문자 인식')
    ap.add_argument('--image-dir', required=True, type=Path, help='원본 이미지 폴더(하위 폴더 포함)')
    ap.add_argument('--out-dir', required=True, type=Path, help='결과 JSON 저장 폴더')
    args = ap.parse_args()

    load_dotenv()
    if not os.getenv('GOOGLE_APPLICATION_CREDENTIALS'):
        raise SystemExit('[중단] GOOGLE_APPLICATION_CREDENTIALS 가 설정되지 않았습니다.')
    client = vision.ImageAnnotatorClient()

    images = sorted(p for p in args.image_dir.rglob('*')
                    if p.suffix.lower() in IMAGE_EXTS
                    and not any(pt.startswith('_') for pt in p.relative_to(args.image_dir).parts))
    args.out_dir.mkdir(parents=True, exist_ok=True)

    done = 0
    for img_path in images:
        json_path = args.out_dir / f'{img_path.stem}.json'
        if json_path.exists():          # 이미 처리한 면은 건너뛴다
            continue
        rel = img_path.relative_to(args.image_dir)
        try:
            text = extract_text_from_image(client, img_path)
            data = {"file_path": str(rel), "vision_raw": text, "gemini_corrected": ""}
            json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
            done += 1
            print(f'[완료] {rel}  {len(text)}자', flush=True)
        except Exception as e:
            print(f'[오류] {rel}: {str(e)[:120]}', flush=True)
    print(f'신규 처리 {done}면')


if __name__ == '__main__':
    main()
