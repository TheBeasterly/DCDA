# app.py
import os
import uuid
import time
import json
import traceback
import pandas as pd # Make sure pandas is imported
from flask import (Flask, request, render_template, redirect,
                   url_for, send_from_directory, flash, session,
                   Response, jsonify)
from werkzeug.utils import secure_filename
# from werkzeug.exceptions import NotFound # Optional

# Assume RVToolsAnalysis_web.py contains the process_rvtools_file generator
try:
    # Ensure RVToolsAnalysis_web.py uses "-ANALYZED" suffix
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
DEFAULT_PROJECT_NAME = "General"
ALLOWED_EXTENSIONS = {'xlsx'}
os.makedirs(DATA_FOLDER, exist_ok=True)

# Initialize Flask App
app = Flask(__name__)
app.config['DATA_FOLDER'] = DATA_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'change-this-dev-secret-key-immediately-please-again-again')


# --- Temporary Task Storage ---
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
    client_list = []; data_folder_path = app.config['DATA_FOLDER']
    try:
        if os.path.exists(data_folder_path): items = os.listdir(data_folder_path); client_list = [i for i in items if os.path.isdir(os.path.join(data_folder_path, i)) and not i.startswith('.')]; client_list.sort(key=str.lower)
    except Exception as e: print(f"Error scanning clients: {e}"); flash("Error retrieving client list.", "error")
    return render_template('rvtools_upload.html', client_list=client_list)


@app.route('/get_projects/<client_name>')
def get_projects(client_name):
    """API endpoint (called by JS) to get project folders for a client."""
    safe_client_name = secure_filename(client_name)
    if not safe_client_name: return jsonify({"error": "Invalid client name"}), 400
    client_dir = os.path.join(app.config['DATA_FOLDER'], safe_client_name); project_list = []
    try:
        if not os.path.isdir(client_dir): return jsonify({"projects": []})
        items = os.listdir(client_dir)
        project_list = [item for item in items if os.path.isdir(os.path.join(client_dir, item)) and not item.startswith('.') and item not in ['Originals', 'Analyzed']]
        project_list.sort(key=str.lower); return jsonify({"projects": project_list})
    except Exception as e: print(f"API Error scan projects '{safe_client_name}': {e}"); return jsonify({"error": "Server error scanning projects."}), 500


@app.route('/check_file_exists/<client_name>/<project_name>/<path:original_filename>')
def check_file_exists(client_name, project_name, original_filename):
    """API endpoint to check if original or analyzed file exists."""
    safe_client_name = secure_filename(client_name); safe_project_name = secure_filename(project_name)
    if not safe_client_name or not safe_project_name or not original_filename: return jsonify({"error": "Invalid input"}), 400
    try:
        client_project_path = os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name)
        original_filepath = os.path.join(client_project_path, 'Originals', original_filename)
        base_name, ext = os.path.splitext(original_filename); analyzed_filename = f"{base_name}-ANALYZED{ext}"; analyzed_filepath = os.path.join(client_project_path, 'Analyzed', analyzed_filename)
        original_exists = os.path.exists(original_filepath) and os.path.isfile(original_filepath); analyzed_exists = os.path.exists(analyzed_filepath) and os.path.isfile(analyzed_filepath)
        print(f"API Check: C='{safe_client_name}', P='{safe_project_name}', F='{original_filename}' -> OrigExists: {original_exists}, AnalyzedExists: {analyzed_exists}")
        return jsonify({"original_exists": original_exists, "analyzed_exists": analyzed_exists})
    except Exception as e: print(f"API Check Error: {e}"); traceback.print_exc(); return jsonify({"error": "Server error checking file."}), 500


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles upload, checks collision action, saves file, starts task."""
    form_redirect_target = 'rvtools_analysis'
    if 'rvtools_file' not in request.files: flash('No file part.', 'error'); return redirect(url_for(form_redirect_target))
    file = request.files['rvtools_file']; original_filename = file.filename
    if not original_filename or not file or not allowed_file(original_filename): flash('No file or invalid type.', 'error'); return redirect(url_for(form_redirect_target))
    selected_client = request.form.get('selected_client'); new_client_name = request.form.get('new_client_name', '').strip(); client_name_to_use = ""
    if selected_client == '_new' and new_client_name: client_name_to_use = new_client_name
    elif selected_client and selected_client != '_new': client_name_to_use = selected_client
    else: flash('Select or add client.', 'error'); return redirect(url_for(form_redirect_target))
    safe_client_name = secure_filename(client_name_to_use)
    if not safe_client_name: flash(f'Invalid client name: "{client_name_to_use}".', 'error'); return redirect(url_for(form_redirect_target))
    selected_project = request.form.get('selected_project'); new_project_name = request.form.get('new_project_name', '').strip(); project_name_to_use = ""
    if selected_project == '_new' and new_project_name: project_name_to_use = new_project_name
    elif selected_project and selected_project not in ['_new', '', '_none_selected', '_general_select', DEFAULT_PROJECT_NAME]: project_name_to_use = selected_project
    elif selected_project == DEFAULT_PROJECT_NAME: project_name_to_use = DEFAULT_PROJECT_NAME
    else: project_name_to_use = DEFAULT_PROJECT_NAME
    safe_project_name = secure_filename(project_name_to_use)
    if not safe_project_name: safe_project_name = DEFAULT_PROJECT_NAME; print(f"Warning: Project name invalid, using default '{DEFAULT_PROJECT_NAME}'.")
    collision_action = request.form.get('collision_action')
    print(f"Upload: Client='{safe_client_name}', Project='{safe_project_name}', File='{original_filename}', Action='{collision_action}'")
    try: # Construct paths
        client_project_path = os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name)
        upload_dir = os.path.join(client_project_path, 'Originals'); analyzed_dir = os.path.join(client_project_path, 'Analyzed')
        input_filepath = os.path.join(upload_dir, original_filename);
        analyzed_base, analyzed_ext = os.path.splitext(original_filename); analyzed_filename = f"{analyzed_base}-ANALYZED{analyzed_ext}"; analyzed_filepath = os.path.join(analyzed_dir, analyzed_filename)
        os.makedirs(upload_dir, exist_ok=True); os.makedirs(analyzed_dir, exist_ok=True)
    except Exception as e: flash(f"Path Error: {e}", 'error'); print(f"Upload Path Error: {e}"); return redirect(url_for(form_redirect_target))
    # Server-side collision check
    original_exists = os.path.exists(input_filepath) and os.path.isfile(input_filepath); analyzed_exists = os.path.exists(analyzed_filepath) and os.path.isfile(analyzed_filepath); collision_detected = original_exists or analyzed_exists
    if collision_detected and collision_action != 'overwrite': flash(f'Error: File exists. Upload cancelled.', 'error'); print(f"Upload Error: Collision detected, overwrite denied."); return redirect(url_for(form_redirect_target))
    elif collision_detected and collision_action == 'overwrite': print(f"Upload Info: Collision detected, proceeding with OVERWRITE.")
    try: # Save File & Store Task
        file.save(input_filepath); task_id = uuid.uuid4().hex[:16]
        TASK_INFO[task_id] = {'original_filename': original_filename, 'input_filepath': input_filepath,'output_folder': analyzed_dir, 'client_name': safe_client_name,'project_name': safe_project_name,'timestamp': time.time()}
        print(f"Upload: Task {task_id} stored."); return redirect(url_for('processing_page', task_id=task_id))
    except Exception as e: flash(f'Save Error: {e}', 'error'); print(f"Upload Save Error: {e}"); return redirect(url_for(form_redirect_target))


@app.route('/processing/<task_id>')
def processing_page(task_id):
    task_details = TASK_INFO.get(task_id);
    if not task_details: flash("Task not found.", "error"); return redirect(url_for('index'))
    return render_template('processing.html', task_id=task_id, display_filename=task_details.get('original_filename', '?'), client_name=task_details.get('client_name', '?'), project_name=task_details.get('project_name', '?'))


@app.route('/stream/<task_id>')
def stream(task_id):
    # --- Stream route uses corrected indentation ---
    print(f"SSE Stream: Request for task_id: {task_id}")
    task_details = TASK_INFO.get(task_id)
    if not task_details:
        print(f"SSE Stream Error: Task details not found task {task_id}.")
        def error_stream_task_not_found(): yield f"event:result\ndata:{json.dumps({'success':False,'message':'Task details not found.'})}\n\n"
        return Response(error_stream_task_not_found(), mimetype='text/event-stream')
    input_filepath = task_details['input_filepath']; original_basename = task_details['original_filename']
    output_folder = task_details['output_folder']; safe_client_name = task_details['client_name']; safe_project_name = task_details['project_name']
    print(f"SSE Stream: Task {task_id} details retrieved.")
    if not os.path.exists(input_filepath):
        print(f"SSE Stream Error: Input file '{input_filepath}' missing task {task_id}.")
        def error_stream_file_missing(): yield f"event:result\ndata:{json.dumps({'success':False,'message':f'Input file {original_basename} missing.'})}\n\n"
        if task_id in TASK_INFO: TASK_INFO.pop(task_id)
        return Response(error_stream_file_missing(), mimetype='text/event-stream')
    # Event Stream Generator
    def event_stream():
        try:
            generator=process_rvtools_file(input_filepath,output_folder,original_basename); final_payload_str=None; print(f"SSE Stream: Starting generator {task_id}")
            for msg in generator: # Intercept result
                 if msg.startswith("event: result"):
                     try: data_part = msg.split("data: ", 1)[1].rsplit("\n\n", 1)[0]; final_payload_str = data_part; print(f"SSE Stream: Intercepted result {task_id}")
                     except IndexError: err_data = json.dumps({"success": False, "message": "Malformed result."}); yield f"event: result\ndata: {err_data}\n\n"; final_payload_str = None; break
                 else: yield msg
            if final_payload_str: # Process final result
                try:
                    payload=json.loads(final_payload_str);
                    if payload.get("success"):
                        fname=payload.get("message"); # Should be "File-ANALYZED.ext"
                        if not fname: raise ValueError("Filename missing.")
                        rel_path=os.path.join(safe_client_name,safe_project_name,'Analyzed',fname).replace("\\","/")
                        # Create ENHANCED success payload
                        success_data = {"success": True, "download_path": rel_path, "client_name": safe_client_name, "project_name": safe_project_name, "analyzed_filename": fname }
                        yield f"event:result\ndata:{json.dumps(success_data)}\n\n"; print(f"SSE Stream: Success task {task_id}. Path: '{rel_path}'")
                    else: yield f"event:result\ndata:{final_payload_str}\n\n"; print(f"SSE Stream: Failure {task_id}: {payload.get('message')}")
                except Exception as e: err_data = json.dumps({"success": False, "message": f"Error finalizing: {e}"}); yield f"event:result\ndata: {err_data}\n\n"; print(f"SSE Finalizing Error {task_id}: {e}")
            else: err_data = json.dumps({'success': False, 'message': 'Script finished unexpectedly.'}); yield f"event:result\ndata: {err_data}\n\n"; print(f"SSE Error: Result missing {task_id}")
        except Exception as e: error_payload = json.dumps({"success": False, "message": f"Server error: {e}"}); yield f"event:result\ndata: {error_payload}\n\n"; traceback.print_exc(); print(f"SSE FATAL {task_id}: {e}")
        finally: TASK_INFO.pop(task_id,None); print(f"SSE Stream: Cleaned task {task_id}. Input '{input_filepath}' preserved.")
    response=Response(event_stream(),mimetype='text/event-stream'); response.headers['Cache-Control']='no-cache'; return response


@app.route('/download/<path:filepath>')
def download_file(filepath):
    data_folder_abs=os.path.abspath(app.config['DATA_FOLDER']); req_path_abs=os.path.normpath(os.path.join(data_folder_abs,filepath))
    if not req_path_abs.startswith(data_folder_abs): flash("Invalid path.", "error"); return redirect(url_for('index'))
    try: return send_from_directory(directory=data_folder_abs, path=filepath, as_attachment=True)
    except FileNotFoundError: flash("File not found.", "error"); return redirect(url_for('index'))
    except Exception as e: flash("Download error.", "error"); traceback.print_exc(); return redirect(url_for('index'))

# ================================================================
# --- Results & Browse Routes ---
# ================================================================

@app.route('/results/<client_name>/<project_name>/<path:analyzed_filename>')
def view_results(client_name, project_name, analyzed_filename):
    """Reads summary sheets from an analyzed Excel file and displays them."""
    print(f"Results Request: C='{client_name}', P='{project_name}', F='{analyzed_filename}'")
    safe_client_name = secure_filename(client_name); safe_project_name = secure_filename(project_name)
    if not safe_client_name or not safe_project_name or not analyzed_filename:
        flash("Invalid client, project, or filename.", "error"); return redirect(url_for('browse_page'))

    summary_data_json = {}; filepath_valid = False; analyzed_filepath = ""
    expected_sheets = [ 'Overall Summary Totals', 'Overall Powerstate Counts', 'Datacenter Summary Combined' ]

    try: # Construct and validate path
        analyzed_filepath = os.path.normpath(os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name, 'Analyzed', analyzed_filename))
        data_folder_abs = os.path.abspath(app.config['DATA_FOLDER'])
        if not os.path.abspath(analyzed_filepath).startswith(data_folder_abs): raise ValueError("Invalid path.")
        if not os.path.isfile(analyzed_filepath): raise FileNotFoundError(f"File not found: {analyzed_filepath}")
        filepath_valid = True
    except FileNotFoundError as e: print(f"Results Error: {e}"); flash(f"File not found: {analyzed_filename}.", "error"); return redirect(url_for('browse_page'))
    except Exception as path_err: print(f"Results Path Error: {path_err}"); flash("Path error.", "error"); return redirect(url_for('browse_page'))

    # Read sheets and convert to JSON
    if filepath_valid:
        try:
            excel_data = pd.read_excel(analyzed_filepath, sheet_name=expected_sheets, engine='openpyxl')
            for sheet_name in expected_sheets:
                if sheet_name in excel_data:
                    df = excel_data[sheet_name]
                    # --- ADDED DEBUG LOG HERE ---
                    print(f"DEBUG: Columns for sheet '{sheet_name}': {df.columns.tolist()}")
                    summary_data_json[sheet_name] = df.to_json(orient='records', date_format='iso', default_handler=str)
                else:
                    summary_data_json[sheet_name] = "[]"; print(f"DEBUG: Sheet '{sheet_name}' not found.")
        except ValueError as ve: print(f"Results Warn: Missing sheets: {ve}"); flash("Warn: Some summaries missing.", "warning");
        except Exception as e: print(f"Results Error: Reading Excel: {e}"); traceback.print_exc(); flash("Error reading results.", "error");
        finally: # Ensure dict has keys even if reading failed partially
            for sheet_name in expected_sheets:
                 if sheet_name not in summary_data_json: summary_data_json[sheet_name] = "[]"

    return render_template('results.html',
                           client_name=safe_client_name, project_name=safe_project_name,
                           analyzed_filename=analyzed_filename, summary_data_json=summary_data_json)


@app.route('/browse')
def browse_existing(): # Renamed function to match user fix
    """Renders the single-page browse interface."""
    print("Browse Request: Loading single browse page.")
    client_list = []; data_folder_path = app.config['DATA_FOLDER']
    try:
        if os.path.exists(data_folder_path): items = os.listdir(data_folder_path); client_list = [ i for i in items if os.path.isdir(os.path.join(data_folder_path, i)) and not i.startswith('.') ]; client_list.sort(key=str.lower)
    except Exception as e: print(f"Browse Error: Clients: {e}"); flash("Error listing clients.", "error"); client_list = []
    return render_template('browse_existing.html', client_list=client_list)


@app.route('/get_files/<client_name>/<project_name>') # API
def get_files(client_name, project_name):
    safe_client_name = secure_filename(client_name); safe_project_name = secure_filename(project_name)
    if not safe_client_name or not safe_project_name: return jsonify({"error": "Invalid client/project name"}), 400
    project_dir = os.path.join(app.config['DATA_FOLDER'], safe_client_name, safe_project_name)
    originals_dir = os.path.join(project_dir, 'Originals'); analyzed_dir = os.path.join(project_dir, 'Analyzed')
    originals_list = []; analyzed_list = []
    try: # Scan Originals
        if os.path.isdir(originals_dir): items = os.listdir(originals_dir); originals_list = [f for f in items if os.path.isfile(os.path.join(originals_dir, f)) and not f.startswith('.')]; originals_list.sort(key=str.lower)
    except Exception as e: print(f"API Error: Originals: {e}")
    try: # Scan Analyzed
        if os.path.isdir(analyzed_dir): items = os.listdir(analyzed_dir); analyzed_list = [f for f in items if os.path.isfile(os.path.join(analyzed_dir, f)) and not f.startswith('.')]; analyzed_list.sort(key=str.lower)
    except Exception as e: print(f"API Error: Analyzed: {e}")
    return jsonify({"originals": originals_list, "analyzed": analyzed_list})


# --- Removed Old Multi-Page Browse Routes ---

# ================================================================
# --- Run the App ---
# ================================================================
if __name__ == '__main__':
    print(f"--- Starting Datacenter Data Analyzer ---")
    print(f"Data Folder: {app.config['DATA_FOLDER']}")
    app.run(host='0.0.0.0', port=5000, debug=True) # Set debug=False for production