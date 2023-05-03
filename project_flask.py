from flask import Flask, render_template, request
from flask import send_file
import os
import openpyxl
from googletrans import Translator
from datetime import datetime
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication 


app = Flask(__name__)

# @app.route('/')
# def index():
#     return render_template('trans_file.html')
###########################################################################
@app.route('/')
def index():
    return render_template('login.html')

@app.route('/login_check', methods=['GET'])
def login_check():
    username = request.args.get('username')
    password = request.args.get('password')

    if username == 'admin' and password == 'admin123':
        message = '환영합니다.'
        return render_template('trans_file.html')
    else:
        message = '로그인 실패.아이디와 비밀번호를 다시 확인하세요.'
        return render_template('login_check.html', message=message)

############################################################################
@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]
    file.save(os.path.join("uploads", file.filename))
    
    # 엑셀 파일을 불러온 뒤 활성화된 시트를 선택
    workbook = openpyxl.load_workbook(os.path.join("uploads", file.filename))
    sheet = workbook.active

    # 구글 번역 기능
    translator = Translator()

    # 각 cell을 선택하여 value 값을 번역
    for row in sheet.iter_rows():
        for cell in row:
            translated_text = translator.translate(cell.value, dest='en').text
            cell.value = translated_text

    # 새로운 엑셀 파일로 저장
    workbook.save('translated_download.xlsx')
    
    return render_template('result_trans.html', file_name=file.filename)


@app.route("/compress", methods=["POST"])
def compress():
    uploads_dir = "uploads"
    files = request.form.getlist("files")
    zip_path = os.path.join(uploads_dir, "compressed_to_email.zip")
    with zipfile.ZipFile(zip_path, "w") as zip_file:
        for file in files:
            file_path = os.path.join(uploads_dir, file),
            zip_file.write(file_path, file)

    compressed_file = os.path.basename(zip_path)

    return render_template("result_compress.html", compressed_file=compressed_file)

@app.route('/download_report')
def download_report():
    return send_file('translated_download.xlsx', as_attachment=True)

@app.route('/send_report')
def send_report():
    smtp = smtplib.SMTP('smtp.naver.com', 587)
    smtp.ehlo()
    smtp.starttls()

    smtp.login('<--id-->','<--pw-->')    ####################################비번 입력#####################################

    myemail = '<-->@naver.com'
    youremail = '<-->@naver.com'

    msg = MIMEMultipart()

    msg['Subject'] ="파일 압축 첨부합니다."
    msg['From'] = myemail
    msg['To'] = youremail

    text="""
    엑셀 파일 압축해서 첨부하였습니다.
    감사합니다.
    """
    contentPart=MIMEText(text)
    msg.attach(contentPart)

    compressed_file = 'D:\\project\\uploads\\compressed_to_email.zip'
    with open(compressed_file, 'rb') as f:
        etc_part=MIMEApplication(f.read())
        etc_part.add_header('Content-Disposition','attachment', filename='compressed_to_email.zip')
        msg.attach(etc_part)

    smtp.sendmail(myemail, youremail, msg.as_string())
    smtp.quit()

    return "이메일이 성공적으로 전송되었습니다."

if __name__ == '__main__':
    app.run(debug=True)