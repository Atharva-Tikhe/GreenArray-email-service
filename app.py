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
            
            f.save(f'uploads/{new_filename}')  # Save file with unique name
            # Handle successful upload logic here (e.g., display success message)
            
            try:
                if main.check_cols(f'uploads/{new_filename}'):
                    main.import_excel(f'uploads/{new_filename}')
            except Exception as e:
                return render_template('index.html', status = f"Columns do not match {e}")
                

            return render_template('index.html', status = "File uploaded successfully!")
        
        return redirect(url_for('index'))  # Redirect to avoid duplicate submissions

    return render_template('index.html', status = "Error with file!")

if __name__ == '__main__':
    app.run(debug=True)
