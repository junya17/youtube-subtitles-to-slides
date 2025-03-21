from flask import Flask, render_template, request, redirect, url_for, flash
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # required for flashing messages

# Load environment variables
from dotenv import load_dotenv
load_dotenv()

SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
PRESENTATION_ID = os.getenv("PRESENTATION_ID")
SCOPES = ['https://www.googleapis.com/auth/presentations']

# Set up Google Slides API client
def get_slides_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build('slides', 'v1', credentials=creds)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        subtitles = request.form['subtitles'].strip().split('\n')
        if not subtitles:
            flash('Please enter at least one subtitle.')
            return redirect(url_for('index'))

        try:
            service = get_slides_service()
            # Delete existing slides (except the first one)
            presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
            slides = presentation.get('slides', [])
            if len(slides) > 0:
                for slide in slides:
                    slide_id = slide['objectId']
                    service.presentations().batchUpdate(
                        presentationId=PRESENTATION_ID,
                        body={'requests': [{'deleteObject': {'objectId': slide_id}}]}
                    ).execute()

            # Add new slides
            for i, text in enumerate(subtitles):
                slide_id = f'slide_{i}'
                service.presentations().batchUpdate(
                    presentationId=PRESENTATION_ID,
                    body={
                        'requests': [
                            {
                                'createSlide': {
                                    'objectId': slide_id,
                                    'slideLayoutReference': {
                                        'predefinedLayout': 'TITLE_AND_BODY'
                                    }
                                }
                            }
                        ]
                    }
                ).execute()

            # Re-fetch to insert text
            presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()
            slides = presentation.get('slides', [])

            for i, text in enumerate(subtitles):
                slide = slides[i]
                for element in slide.get('pageElements', []):
                    shape = element.get('shape')
                    if shape and shape.get('shapeType') == 'TEXT_BOX':
                        text_id = element.get('objectId')
                        service.presentations().batchUpdate(
                            presentationId=PRESENTATION_ID,
                            body={
                                'requests': [
                                    {
                                        'insertText': {
                                            'objectId': text_id,
                                            'text': text
                                        }
                                    }
                                ]
                            }
                        ).execute()
                        break

            flash('Slides updated successfully!')
            return redirect(url_for('index'))

        except Exception as e:
            flash(f'An error occurred: {e}')
            return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
    app.debug = True

