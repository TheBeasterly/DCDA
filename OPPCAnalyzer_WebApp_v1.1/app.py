# app.py
import os
import uuid
import time
import json
import traceback
from flask import (Flask, request, render_template, redirect,
                   url_for, send_from_directory, flash, session,
                   Response, jsonify) # Added jsonify
from werkzeug.utils import secure_filename
# from werkzeug.exceptions import NotFound # Optional

# Assume RVToolsAnalysis_web.py contains the process_rvtools_file generator
# Make sure it's in the same directory or accessible via Python path
try:
    # Ensure RVToolsAnalysis_web.py is updated to use "-ANALYZED" suffix
    from RVToolsAnalysis_web import process_rvtools_file
except ImportError:
    print("ERROR: Could not import 'process_rvtools_file' from 'RVToolsAnalysis_web'.")
    print("Ensure the file 'RVToolsAnalysis_web.py' exists and is updated.")
    # Dummy function to allow Flask to start, yields an error via SSE
    def process_rvtools_file(*args, **kwargs):
        yield f"event: result\ndata: {json.dumps({'success': False, 'message': 'Error: Backend processing script not found.'})}\n\n"


# --- Configuration ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_FOLDER = os.path.join(BASE_DIR, 'client_data')
DEFAULT_PROJECT_NAME = "General" # Default project folder name
ALLOWED_EXTENSIONS = {'xlsx'}

os.makedirs(DATA_FOLDER, exist_ok=True) # Ensure base data folder exists

# Initialize Flask App
app = Flask(__name__)
app.config['DATA_FOLDER'] = DATA_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024 # 32 MB upload limit
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'change-this-dev-secret-key-immediately') # Use env var!


# --- Temporary Task Storage (In-memory) ---
TASK_INFO = {}

# --- Helper Function ---
def allowed_file(filename):
    """Checks if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ================================================================
# --- Core Application Routes ---
# ================================================================

@app.route('/', methods=['GET'])
def index():
    """Renders the main landing/welcome page."""
    return render_template('landing.html')


@app.route('/rvtools_analysis', methods=['GET'])
def rvtools_analysis():
    """Renders the upload form page for RVTools analysis."""
    client_list = []
    data_folder_path = app.config['DATA_FOLDER']
    try:
        if os.path.exists(data_folder_path):
            items = os.listdir(data_folder_path)
            client_list = [i for i in items if os.path.isdir(os.path.join(data_folder_path, i)) and not i.startswith('.')]
            client_list.sort(key=str.lower)
    except Exception as e:
        print(f"Error scanning clients: {e}")
        flash("Error retrieving client list.", "error")
    # Renders templates/rvtools_upload.html
    return render_template('rvtools_upload.html', client_list=client_list)


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles the actual file upload POST request."""
    form_redirect_target = 'rvtools_analysis' # Redirect here on error

    if 'rvtools_file' not in request.files: flash('No file part.', 'error'); return redirect(url_for(form_redirect_target))
    file = request.files['rvtools_file']; original_filename = file.filename
    if not original_filename or not file or not allowed_file(original_filename): flash('No file or invalid type.', 'error'); return redirect(url_for(form_redirect_target))

    # Get Client Info
    selected_client = request.form.get('selected_client'); new_client_name = request.form.get('new_client_name', '').strip(); client_name_to_use = ""
    if selected_client == '_new' and new_client_name: client_name_to_use = new_client_name
    elif selected_client and selected_client != '_new': client_name_to_use = selected_client
    else: flash('Select or add client.', 'error'); return redirect(url_for(form_redirect_target))
    safe_client_name = secure_filename(client_name_to_use)
    if not safe_client_name: flash(f'Invalid client name: "{client_name_to_use}".', 'error'); return redirect(url_for(form_redirect_target))

    # Get Project Info (Handles Default)
    selected_project = request.form.get('selected_project'); new_project_name = request.form.get('new_project_name', '').strip(); project_name_to_use = ""
    if selected_project == '_new' and new_project_name: project_name_to_use = new_project_name
    elif selected_project and selected_project not in ['_new', '', '_none_selected', '_general_select', DEFAULT_PROJECT_NAME]: project_name_to_use = selected_project
    elif selected_project == DEFAULT_PROJECT_NAME: project_name_to_use = DEFAULT_PROJECT_NAME
    else: project_name_to_use = DEFAULT_PROJECT_NAME
    safe_project_name = secure_filename(project_name_to_use)
    if not safe_project_name: safe_project_name = DEFAULT_PROJECT_NAME; print(f"Warning: Project name invalid, using default '{DEFAULT_PROJECT_NAME}'.")
    print(f"Upload: Client='{safe_client_name}', Project='{safe_project_name}'")

    # Construct Paths & Create Dirs (Using Originals/Analyzed)
    try:
        client_project_path = os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name)
        upload_dir = os.path.join(client_project_path, 'Originals'); analyzed_dir = os.path.join(client_project_path, 'Analyzed')
        input_filepath = os.path.join(upload_dir, original_filename); os.makedirs(upload_dir, exist_ok=True); os.makedirs(analyzed_dir, exist_ok=True)
    except Exception as e: flash(f"Path Error: {e}", 'error'); print(f"Upload Path Error: {e}"); return redirect(url_for(form_redirect_target))

    # Save File & Store Task
    if os.path.exists(input_filepath): flash(f'File "{original_filename}" exists in Originals.', 'error'); return redirect(url_for(form_redirect_target))
    try:
        file.save(input_filepath); task_id = uuid.uuid4().hex[:16]
        TASK_INFO[task_id] = {'original_filename': original_filename, 'input_filepath': input_filepath,'output_folder': analyzed_dir, 'client_name': safe_client_name,'project_name': safe_project_name,'timestamp': time.time()}
        print(f"Upload: Task {task_id} stored."); return redirect(url_for('processing_page', task_id=task_id))
    except Exception as e: flash(f'Save Error: {e}', 'error'); print(f"Upload Save Error: {e}"); return redirect(url_for(form_redirect_target))


@app.route('/processing/<task_id>')
def processing_page(task_id):
    """Displays the page that shows progress updates."""
    task_details = TASK_INFO.get(task_id);
    if not task_details: flash("Task not found.", "error"); return redirect(url_for('index'))
    return render_template('processing.html', task_id=task_id, display_filename=task_details.get('original_filename', '?'), client_name=task_details.get('client_name', '?'), project_name=task_details.get('project_name', '?'))


@app.route('/stream/<task_id>')
def stream(task_id):
    """Streams progress events (SSE) for a given task."""
    # --- This complex function remains the same as the last fully corrected version ---
    print(f"SSE Stream: Request for task_id: {task_id}")
    task_details = TASK_INFO.get(task_id)
    if not task_details: # Check Task Exists
        def e_404(): yield f"event:result\ndata:{json.dumps({'success':False,'message':'Task not found'})}\n\n"; print(f"SSE 404: Task {task_id}")
        return Response(e_404(), mimetype='text/event-stream')
    input_filepath=task_details['input_filepath']; original_basename=task_details['original_filename']; output_folder=task_details['output_folder']
    safe_client_name=task_details['client_name']; safe_project_name=task_details['project_name']
    if not os.path.exists(input_filepath): # Check Input File Exists
        def e_file(): yield f"event:result\ndata:{json.dumps({'success':False,'message':f'Input file {original_basename} missing.'})}\n\n"; print(f"SSE Error: Input missing {task_id}")
        if task_id in TASK_INFO: TASK_INFO.pop(task_id)
        return Response(e_file(), mimetype='text/event-stream')
    # Event Stream Generator
    def event_stream():
        try:
            generator=process_rvtools_file(input_filepath,output_folder,original_basename); final_payload_str=None; print(f"SSE Stream: Starting generator {task_id}")
            for msg in generator: # Intercept result message
                 if msg.startswith("event: result"):
                     try: data_part=msg.split("data: ", 1)[1].rsplit("\n\n", 1)[0]; final_payload_str=data_part; print(f"SSE Stream: Got result {task_id}")
                     except IndexError: err_data=json.dumps({"success": False,"message":"Malformed result."}); yield f"event: result\ndata: {err_data}\n\n"; final_payload_str=None; break
                 else: yield msg # Pass progress
            if final_payload_str: # Process final result
                try:
                    payload=json.loads(final_payload_str);
                    if payload.get("success"):
                        fname=payload.get("message");
                        if not fname: raise ValueError("Filename missing.")
                        rel_path=os.path.join(safe_client_name,safe_project_name,'Analyzed',fname).replace("\\","/")
                        yield f"event:result\ndata:{json.dumps({'success':True,'message':rel_path})}\n\n"; print(f"SSE Stream: Success {task_id}. Path: {rel_path}")
                    else: yield f"event:result\ndata:{final_payload_str}\n\n"; print(f"SSE Stream: Failure {task_id}: {payload.get('message')}")
                except Exception as e: err_data=json.dumps({"success": False,"message":f"Finalizing error: {e}"}); yield f"event:result\ndata:{err_data}\n\n"; print(f"SSE Finalizing Error {task_id}: {e}")
            else: err_data=json.dumps({'success': False,'message':'Script finished unexpectedly.'}); yield f"event:result\ndata:{err_data}\n\n"; print(f"SSE Error: Result missing {task_id}")
        except Exception as e: error_payload = json.dumps({"success": False,"message":f"Server error: {e}"}); yield f"event:result\ndata:{error_payload}\n\n"; traceback.print_exc(); print(f"SSE FATAL {task_id}: {e}")
        finally: TASK_INFO.pop(task_id,None); print(f"SSE Stream: Cleaned task {task_id}. Input '{input_filepath}' preserved.")
    response=Response(event_stream(),mimetype='text/event-stream'); response.headers['Cache-Control']='no-cache'; return response


@app.route('/download/<path:filepath>')
def download_file(filepath):
    """Provides files for download using a relative path within DATA_FOLDER."""
    # This function remains the same as the previously corrected version
    print(f"Download Request: Relative path: '{filepath}'")
    data_folder_abs=os.path.abspath(app.config['DATA_FOLDER']); req_path_abs=os.path.normpath(os.path.join(data_folder_abs,filepath))
    if not req_path_abs.startswith(data_folder_abs): print(f"Download Error: Path traversal: '{filepath}'"); flash("Invalid path.", "error"); return redirect(url_for('index'))
    try: return send_from_directory(directory=data_folder_abs, path=filepath, as_attachment=True)
    except FileNotFoundError: print(f"Download Error: File not found: '{req_path_abs}'"); flash("File not found.", "error"); return redirect(url_for('index'))
    except Exception as e: print(f"Download Error: Unexpected: {e}"); traceback.print_exc(); flash("Download error.", "error"); return redirect(url_for('index'))

# ================================================================
# --- Browse Routes (Single Page Approach) ---
# ================================================================

# --- Route for the main single-page browse UI ---
@app.route('/browse')
def browse_existing(): # Renamed from browse_clients
    """Scans for client folders and renders the single-page browse interface."""
    print("Browse Request: Loading single browse page.")
    client_list = []
    data_folder_path = app.config['DATA_FOLDER']
    try:
        if os.path.exists(data_folder_path):
            items = os.listdir(data_folder_path)
            client_list = [ item for item in items if os.path.isdir(os.path.join(data_folder_path, item)) and not item.startswith('.') ]
            client_list.sort(key=str.lower)
        else: print(f"Browse Info: Data folder does not exist: {data_folder_path}")
        print(f"Browse: Found clients for initial list: {client_list}")
    except Exception as e:
        print(f"Browse Error: Scanning clients: {e}")
        traceback.print_exc(); flash("Error retrieving client list.", "error")
        client_list = [] # Ensure empty list on error

    # Renders templates/browse_existing.html (This template needs to be created next)
    return render_template('browse_existing.html', client_list=client_list)


# --- API Route to get projects for a client (used by browse JS) ---
# This route remains the same as defined previously
@app.route('/get_projects/<client_name>')
def get_projects(client_name):
    """API endpoint (called by JS) to get project folders for a client."""
    print(f"API Request: Get projects for client '{client_name}'")
    safe_client_name = secure_filename(client_name)
    if not safe_client_name: return jsonify({"error": "Invalid client name"}), 400
    client_dir = os.path.join(app.config['DATA_FOLDER'], safe_client_name); project_list = []
    try:
        if not os.path.isdir(client_dir): return jsonify({"projects": []}) # OK if dir doesn't exist
        items = os.listdir(client_dir)
        project_list = [item for item in items if os.path.isdir(os.path.join(client_dir, item)) and not item.startswith('.') and item not in ['Originals', 'Analyzed']]
        project_list.sort(key=str.lower); return jsonify({"projects": project_list})
    except Exception as e: print(f"API Error scan projects '{safe_client_name}': {e}"); return jsonify({"error": "Server error scanning projects."}), 500


# --- *** NEW API Route: /get_files *** ---
# This route is called by the JavaScript on the single browse page
@app.route('/get_files/<client_name>/<project_name>')
def get_files(client_name, project_name):
    """API endpoint (called by JS) to get file lists for a client/project."""
    print(f"API Request: Get files for '{client_name}' / '{project_name}'")
    safe_client_name = secure_filename(client_name)
    safe_project_name = secure_filename(project_name)

    if not safe_client_name or not safe_project_name:
        print("API Error: Invalid client or project name format.")
        return jsonify({"error": "Invalid client or project name"}), 400

    # Construct paths to Originals and Analyzed folders
    project_dir = os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name)
    originals_dir = os.path.join(project_dir, 'Originals')
    analyzed_dir = os.path.join(project_dir, 'Analyzed')

    originals_list = []
    analyzed_list = []

    # Scan Originals safely
    try:
        if os.path.isdir(originals_dir): # Check if directory exists first
            items = os.listdir(originals_dir)
            originals_list = [f for f in items if os.path.isfile(os.path.join(originals_dir, f)) and not f.startswith('.')]
            originals_list.sort(key=str.lower)
    except Exception as e: print(f"API Error: Scanning Originals '{originals_dir}': {e}") # Log error but continue

    # Scan Analyzed safely
    try:
        if os.path.isdir(analyzed_dir): # Check if directory exists first
            items = os.listdir(analyzed_dir)
            analyzed_list = [f for f in items if os.path.isfile(os.path.join(analyzed_dir, f)) and not f.startswith('.')]
            analyzed_list.sort(key=str.lower)
    except Exception as e: print(f"API Error: Scanning Analyzed '{analyzed_dir}': {e}") # Log error but continue

    print(f"API Success: Files found - Originals: {len(originals_list)}, Analyzed: {len(analyzed_list)}")
    # Return both lists as JSON
    return jsonify({
        "originals": originals_list,
        "analyzed": analyzed_list
    })


# --- *** REMOVED Old Browse Routes *** ---
# The old @app.route('/browse/<client_name>') [browse_projects] and
# @app.route('/browse/<client_name>/<project_name>') [browse_files]
# that rendered separate HTML pages have been removed.

# ================================================================
# --- Run the App ---
# ================================================================
if __name__ == '__main__':
    print(f"--- Starting Datacenter Data Analyzer ---")
    print(f"Data Folder: {app.config['DATA_FOLDER']}")
    app.run(host='0.0.0.0', port=5000, debug=True) # Set debug=False for production