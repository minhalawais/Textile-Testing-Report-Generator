from web_app import Flask, request, redirect, render_template
from docx2pdf import convert
from generate_qrcode import get_file_image
app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/create', methods=['POST'])
def create():
    req_no = request.form['req_no']

    # Perform any necessary processing with the entered request number
    # For demonstration purposes, we'll just redirect to a success page
    return redirect('/success')

@app.route('/success')
def success():
    get_file_image(page_no=0)
    return render_template('collageScreen.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0',port=40000, debug=True)
