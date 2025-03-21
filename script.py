from googleapiclient.discovery import build
from google.oauth2 import service_account
import time

from dotenv import load_dotenv
import os

load_dotenv()

SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
PRESENTATION_ID = os.getenv("PRESENTATION_ID")

def create_slide(presentation_id, text_list):
    SCOPES = ['https://www.googleapis.com/auth/presentations']
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    service = build('slides', 'v1', credentials=creds)

    # スライド削除（既存スライドがある場合）
    presentation = service.presentations().get(presentationId=presentation_id).execute()
    slides = presentation.get('slides', [])
    if slides:
        first_slide_id = slides[0]['objectId']
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': [{
                'deleteObject': {'objectId': first_slide_id}
            }]}
        ).execute()

    # スライド作成リクエスト
    requests = []
    for i in range(len(text_list)):
        slide_id = f'slide_{i}'
        requests.append({
            'createSlide': {
                'objectId': slide_id,
                'slideLayoutReference': {
                    'predefinedLayout': 'TITLE_AND_BODY'
                }
            }
        })

    # スライドを一括作成
    if requests:
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        print(f"{len(text_list)} slides created.")

    # ★ 少し待つ（重要）
    time.sleep(1.5)

    # 作成後にスライドを再取得
    presentation = service.presentations().get(presentationId=presentation_id).execute()
    slides = presentation.get('slides', [])
    print(f"取得したスライド枚数: {len(slides)}")

    for i, text in enumerate(text_list):
        if i >= len(slides):
            print(f"スライド {i} が見つかりません。スキップ。")
            continue

        slide = slides[i]
        found = False
        for element in slide.get('pageElements', []):
            shape = element.get('shape')
            if shape and shape.get('shapeType') == 'TEXT_BOX':
                text_id = element.get('objectId')
                service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [{
                        'insertText': {
                            'objectId': text_id,
                            'text': text
                        }
                    }]}
                ).execute()
                found = True
                break
        if not found:
            print(f"スライド {i} にテキストボックスが見つかりませんでした。")

# --- サンプル字幕 ---
youtube_subtitles = [
    "Introduction to AI and its impact on society",
    "The history of artificial intelligence",
    "How machine learning differs from traditional programming",
    "Challenges in AI development",
    "Future prospects and ethical considerations"
]

create_slide(PRESENTATION_ID, youtube_subtitles)
