from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
import os
import uuid

import main


app = Flask(__name__)
# app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # Limit file size to 10MB

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Handle file upload
        f = request.files['file']
        if f:
            # Generate a unique filename
            filename, extension = os.path.splitext(f.filename)
            unique_id = str(uuid.uuid4())  # Use uuid for unique ID generation
            new_filename = f"{unique_id}{filename}{extension}"
            
            
            f.save(os.path.join('uploads', new_filename))
            # Handle successful upload logic here (e.g., display success message)
            
            try:
                if main.check_cols(os.path.join('uploads', new_filename)):
                    main.import_excel(os.path.join('uploads', new_filename))
            except Exception as e:
                return render_template('index.html', status = f"Columns do not match {e}")
                

            return render_template('index.html', status = "File uploaded successfully!")
        
        return redirect(url_for('index'))  # Redirect to avoid duplicate submissions

    return render_template('index.html', status = "")

if __name__ == '__main__':
    import subprocess, sys
    if sys.platform == 'darwin':
        subprocess.Popen('open http://127.0.0.1:8080', shell=True)
    elif sys.platform == 'win32':
        subprocess.Popen('chrome http://127.0.0.1:8080', shell=True)
    app.run(debug=True, port=8080)
    
