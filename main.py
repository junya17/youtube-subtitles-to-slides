from googleapiclient.discovery import build
from google.oauth2 import service_account

def create_slide(presentation_id, text_list):
    """
    Googleスライドに字幕を1スライドずつ追加する関数
    """
    SCOPES = ['https://www.googleapis.com/auth/presentations']
    SERVICE_ACCOUNT_FILE = 'path/to/your/service-account.json'  # 自分のサービスアカウントキーに変更

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    service = build('slides', 'v1', credentials=creds)
    
    requests = []
    
    for i, text in enumerate(text_list):
        slide_id = f'slide_{i}'
        requests.append({
            'createSlide': {
                'objectId': slide_id,
                'insertionIndex': '1',
                'slideLayoutReference': {'predefinedLayout': 'TITLE'}
            }
        })
        requests.append({
            'insertText': {
                'objectId': slide_id,
                'text': text
            }
        })
    
    body = {'requests': requests}
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    print(f"{len(text_list)} slides created successfully!")

# サンプル字幕データ
youtube_subtitles = [
    "Introduction to AI and its impact on society",
    "The history of artificial intelligence",
    "How machine learning differs from traditional programming",
    "Challenges in AI development",
    "Future prospects and ethical considerations"
]

# GoogleスライドのプレゼンテーションIDを指定
PRESENTATION_ID = "your_presentation_id_here"  # ここにGoogleスライドのIDを設定

# スライド作成関数を実行
create_slide(PRESENTATION_ID, youtube_subtitles)
