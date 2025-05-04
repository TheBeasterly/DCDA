# app.py
import os
import uuid
import time
import json       # <-- ADDED IMPORT
import traceback  # <-- ADDED IMPORT
from flask import (Flask, request, render_template, redirect,
                   url_for, send_from_directory, flash, session,
                   Response) # Added Response for SSE
from werkzeug.utils import secure_filename

# Import the refactored analysis function (now a generator)
from RVToolsAnalysis_web import process_rvtools_file

# --- Configuration ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
PROCESSED_FOLDER = os.path.join(BASE_DIR, 'processed')
ALLOWED_EXTENSIONS = {'xlsx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Initialize Flask App
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.secret_key = 'your secret key here' # CHANGE THIS!

# --- Temporary Task Storage (for demonstration) ---
# WARNING: This is NOT suitable for production with multiple users/workers.
# A proper task queue (Celery) or database would be needed.
TASK_INFO = {}
# --- --- --- --- --- --- --- --- --- --- --- --- ---

# --- Helper Function ---
def allowed_file(filename):
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- Routes ---
@app.route('/', methods=['GET'])
def index():
    """Renders the main page with the upload form."""
    # No longer needs to handle processed_filename from session here
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles file uploads, stores task info, and redirects to processing page."""
    if 'rvtools_file' not in request.files:
        flash('No file part in the request.')
        return redirect(url_for('index'))

    file = request.files['rvtools_file']

    if file.filename == '':
        flash('No selected file.')
        return redirect(url_for('index'))

    if file and allowed_file(file.filename):
        original_filename = file.filename
        # Use secure_filename on the original name just in case, though we don't pass it in URL now
        secure_original_filename = secure_filename(original_filename)

        # Create a unique internal filename for saving
        unique_id = uuid.uuid4().hex[:16] # Slightly longer UUID part
        input_filename = f"input_{unique_id}.xlsx"
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)

        try:
            file.save(input_filepath)
            print(f"File saved internally as: {input_filepath}") # Server log

            # --- Store task info ---
            # Key: unique internal filename. Value: dict with original name and full path
            TASK_INFO[input_filename] = {
                'original_filename': original_filename, # Store the actual original name
                'input_filepath': input_filepath,
                'timestamp': time.time() # Add timestamp for potential cleanup
            }
            print(f"Stored task info for {input_filename}")
            # --- --- --- --- --- ---

            # Redirect to the processing page, passing the unique internal name
            return redirect(url_for('processing_page', input_filename=input_filename))

        except Exception as e:
            flash(f'An unexpected error occurred during upload: {str(e)}')
            print(f"Error during upload for {original_filename}: {e}") # Server log
            # Clean up if file was saved before error
            if os.path.exists(input_filepath):
                 try: os.remove(input_filepath)
                 except OSError as rm_err: print(f"Error removing failed upload {input_filepath}: {rm_err}")
            return redirect(url_for('index'))

    else:
        flash('Invalid file type. Please upload an .xlsx file.')
        return redirect(url_for('index'))

# --- New Route for Processing Page ---
@app.route('/processing/<input_filename>')
def processing_page(input_filename):
    """Displays the page that will show progress updates."""
    # We only need to pass the input_filename to the template.
    # The template's JavaScript will use this to make the EventSource connection.
    if input_filename not in TASK_INFO:
         flash("Error: Processing task not found or expired.")
         return redirect(url_for('index'))

    # Get the original filename to display on the page initially
    original_filename = TASK_INFO[input_filename].get('original_filename', 'Unknown File')

    return render_template('processing.html',
                           input_filename=input_filename,
                           original_filename=original_filename)


# --- New Route for SSE Stream ---
@app.route('/stream/<input_filename>')
def stream(input_filename):
    """Streams progress events for a given task."""
    print(f"SSE stream requested for: {input_filename}") # Server log

    # --- Retrieve task details ---
    task_details = TASK_INFO.get(input_filename)
    if not task_details:
        # Function to generate an immediate SSE error if task not found
        def error_stream():
             error_payload = json.dumps({"success": False, "message": "Error: Task details not found or expired."})
             yield f"event: result\ndata: {error_payload}\n\n"
        return Response(error_stream(), mimetype='text/event-stream')
    # --- --- --- --- --- --- ---

    input_filepath = task_details['input_filepath']
    original_basename = task_details['original_filename']
    output_folder = app.config['PROCESSED_FOLDER']

    # --- Create the generator and stream response ---
    # process_rvtools_file is now a generator function
    def event_stream():
        try:
            # Call the generator function
            generator = process_rvtools_file(input_filepath, output_folder, original_basename)
            # Yield each message from the generator
            for message in generator:
                yield message
            print(f"SSE stream finished for: {input_filename}")
        except Exception as e:
             # Log unexpected errors during streaming
             print(f"FATAL ERROR during SSE generation for {input_filename}: {e}")
             traceback.print_exc()
             # Yield a final error message to the client if possible
             error_payload = json.dumps({"success": False, "message": f"Unexpected server error during processing: {e}"})
             yield f"event: result\ndata: {error_payload}\n\n"
        finally:
             # --- Cleanup (Optional but recommended) ---
             # Remove task info after streaming is complete or errored
             if input_filename in TASK_INFO:
                 removed_info = TASK_INFO.pop(input_filename)
                 print(f"Removed task info for {input_filename}")
             # Optionally remove the original uploaded input file
             if os.path.exists(input_filepath):
                 try:
                     os.remove(input_filepath)
                     print(f"Removed uploaded file: {input_filepath}")
                 except OSError as e:
                     print(f"Error removing uploaded file {input_filepath} after processing: {e}")
             # --- --- --- --- --- --- --- --- --- ---

    response = Response(event_stream(), mimetype='text/event-stream')
    # Prevent caching of the stream
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response
    # --- --- --- --- --- --- --- --- --- --- --- ---

# --- Download Route (No changes needed from previous version) ---
@app.route('/download/<filename>')
def download_file(filename):
    """Provides the processed file for download."""
    safe_filename = os.path.basename(filename)
    if ".." in safe_filename or safe_filename.startswith(("/", "\\")):
        flash("Invalid filename.")
        return redirect(url_for('index'))

    processed_file_path = os.path.join(app.config['PROCESSED_FOLDER'], safe_filename)

    if not os.path.isfile(processed_file_path):
        flash(f"Error: File '{safe_filename}' not found on server.")
        print(f"Download attempt failed: File not found at {processed_file_path}")
        return redirect(url_for('index'))

    try:
        return send_from_directory(app.config['PROCESSED_FOLDER'],
                                     safe_filename, as_attachment=True)
    except Exception as e:
        flash(f"An error occurred while trying to download the file.")
        print(f"Error during download of {safe_filename}: {e}")
        return redirect(url_for('index'))

# --- Run the App ---
if __name__ == '__main__':
    # REMOVE debug=True for production!
    app.run(debug=True)