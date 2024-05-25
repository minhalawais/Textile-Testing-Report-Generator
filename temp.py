from flask import Flask, render_template
from global_func import get_image_links
app = Flask(__name__)

# Sample image URLs

@app.route('/')
def index():
    image_urls = get_image_links()
    print(image_urls)
    return render_template('images.html', image_urls=image_urls)

if __name__ == '__main__':
    app.run(debug=True)
