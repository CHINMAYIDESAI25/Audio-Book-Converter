from flask import Flask, render_template, request, send_file
import PyPDF2
import docx
import pyttsx3

app = Flask(__name__)

@app.route('/', methods = ['GET', 'POST'])
def index():
    if (request.method == 'POST'):
        uploaded_file = request.files['uploaded_file']
        voice = request.form['voice']
        speed = int(request.form['speed'])
        if (uploaded_file.filename.endswith(".pdf")):
            text = extract_text_from_pdf(uploaded_file)
        elif (uploaded_file.filename.endswith(".docx")):
            text = extract_text_from_docx(uploaded_file)
        else:
            return ("Unsupported File Type")
        convert_text_to_audio(text, voice, speed)
        return send_file('output.mp3', as_attachment = True)
    return render_template('index.html')

def extract_text_from_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page_num in range (len(pdf_reader.pages)):
        text += pdf_reader.pages[page_num].extract_text()
    return text

def extract_text_from_docx(file):
    doc_reader = docx.Document(file)
    text = ""
    for paragraph in doc_reader.paragraphs:
        text += paragraph.text + '\n'
    return text

def convert_text_to_audio(text, voice, speed):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    if (voice == 'male'):
        engine.setProperty('voice', voices[0].id)
    else:
        engine.setProperty('voice', voices[1].id)
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate * speed / 100)
    engine.save_to_file(text, 'output.mp3')
    engine.runAndWait()

if __name__ == '__main__':
    app.run(debug = True)