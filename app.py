"""
Created by: Clement
"""

from flask import Flask

UPLOAD_FOLDER = r'site\uploads'
OUTPUT_FOLDER = r'E:\Users\Clement\Documents\GitHub\shopify-to-dhl-orders\site\output'
app = Flask(__name__, template_folder='./site/templates')
app.secret_key = "secretkey"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024