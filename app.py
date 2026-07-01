from flask import Flask, request, render_template, jsonify, send_file, session
import os
from datetime import datetime
from excel_to_json_anak import process_excel_to_json, validate_template_compliance
from export_analisis import export_analisis_from_json
import uuid

app = Flask(__name__)

# Configuration for both development and production
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'sitrek_stunting_secret_key_2024')

# Platform-specific configurations
is_railway = os.environ.get('RAILWAY_ENVIRONMENT', '') != ''
is_vercel = os.environ.get('VERCEL_ENV', '') != ''
is_serverless = is_railway or is_vercel

# Use /tmp for serverless platforms (read-only filesystem otherwise)
if is_vercel or is_railway:
    app.config['UPLOAD_FOLDER'] = '/tmp/uploads'
    app.config['EXPORT_FOLDER'] = '/tmp/exports'
else:
    app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', 'uploads')
    app.config['EXPORT_FOLDER'] = 'exports'

# Server-side storage for export data (in-memory, works on serverless within same instance)
export_data_store = {}

# Initialize session
if is_vercel:
    pass
else:
    from flask_session import Session
    if is_railway:
        app.config['SESSION_FILE_DIR'] = '/tmp/flask_sessions'
    else:
        app.config['SESSION_FILE_DIR'] = 'flask_sessions'
    app.config['SESSION_TYPE'] = 'filesystem'
    app.config['SESSION_FILE_THRESHOLD'] = 500
    app.config['SESSION_PERMANENT'] = False
    os.makedirs(app.config['SESSION_FILE_DIR'], exist_ok=True)
    sess = Session(app)

# Ensure writable directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
if 'EXPORT_FOLDER' in app.config:
    os.makedirs(app.config['EXPORT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if file and file.filename.endswith(('.xlsx', '.xls')):
            filename = file.filename
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Validate template compliance first
            is_valid, validation_result = validate_template_compliance(filepath)

            if not is_valid:
                # Template validation failed - return detailed error
                return jsonify({
                    'success': False,
                    'message': 'Template validation failed',
                    'validation_error': True,
                    'validation': validation_result,
                    'error': f'Template tidak sesuai: {"; ".join(validation_result.get("errors", ["Unknown error"]))}'
                }), 400

            # Process Excel file to JSON
            result = process_excel_to_json(filepath)

            # Add validation information to the result
            result['validation'] = validation_result

            # Store processed data for export functionality
            export_id = str(uuid.uuid4())
            export_data_store[export_id] = {
                'data': result,
                'upload_timestamp': datetime.now().isoformat(),
                'created_at': datetime.now().isoformat()
            }
            # Use session to store export_id reference (small data, safe for cookies)
            session['export_id'] = export_id
            if not is_vercel:
                # On non-serverless, also store full data in session as fallback
                session['processed_data'] = result
                session['upload_timestamp'] = datetime.now().isoformat()

            # Add export_id to result for frontend
            result['export_id'] = export_id
            # Add has_export_data flag inside result for frontend to pass to display function
            result['has_export_data'] = True

            # Build appropriate message based on validation results
            message = 'File uploaded and processed successfully'
            if validation_result.get('warnings'):
                message += f' (dengan {len(validation_result["warnings"])} peringatan)'

            return jsonify({
                'success': True,
                'message': message,
                'data': result,
                'has_export_data': True,
                'export_id': export_id  # Send export_id to frontend
            })
        else:
            return jsonify({'error': 'Please upload an Excel file (.xlsx or .xls)'}), 400

    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/files')
def list_files():
    try:
        files = []
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            if filename.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                files.append({
                    'name': filename,
                    'size': os.path.getsize(file_path)
                })
        return jsonify({'files': files})
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500

@app.route('/health')
def health():
    """Health check endpoint for monitoring"""
    return jsonify({
        'status': 'healthy',
        'service': 'SiTrack Stunting',
        'version': '1.0.0'
    })

@app.route('/download-template')
def download_template():
    """
    Download template reference file
    """
    try:
        template_path = os.path.join('data_template', 'Data Test.xlsx')
        if os.path.exists(template_path):
            return send_file(template_path,
                           as_attachment=True,
                           download_name='Template_Data_Anak.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return jsonify({'error': 'Template file not found'}), 404
    except Exception as e:
        return jsonify({'error': f'Error downloading template: {str(e)}'}), 500

@app.route('/export-analisis')
def export_analisis():
    """
    Export analisis data pertumbuhan anak ke Excel dengan format analisis
    Accepts optional 'periods' query parameter for filtering by periode
    Example: /export-analisis?periods=JAN%202025&periods=FEB%202025
    """
    try:
        # Get selected periods from query parameters
        selected_periods = request.args.getlist('periods')
        
        # Try to get export_id from session first (priority)
        export_id = session.get('export_id')

        # Get processed data with proper priority
        processed_data = None

        # Priority 1: Use session export_id to get server storage data
        if export_id and export_id in export_data_store:
            processed_data = export_data_store[export_id]['data']
        # Priority 2: Use session data directly (compatibility)
        elif 'processed_data' in session:
            processed_data = session['processed_data']
        # Priority 3: Only use most recent server data as last resort
        elif export_data_store:
            export_id = max(export_data_store.keys(),
                       key=lambda k: export_data_store[k].get('created_at', ''))
            processed_data = export_data_store[export_id]['data']
        else:
            return jsonify({'error': 'Tidak ada data untuk di-export. Silakan upload file terlebih dahulu.'}), 400

        # Double-check that we have valid data structure
        if not processed_data or not isinstance(processed_data, dict):
            return jsonify({'error': 'Data tidak valid. Silakan upload file kembali.'}), 400

        # Check for essential data
        if 'children' not in processed_data or not processed_data['children']:
            return jsonify({'error': 'Tidak ada data anak untuk di-export.'}), 400

        # Filter data by selected periods if specified
        if selected_periods:
            filtered_data = filter_data_by_periods(processed_data, selected_periods)
        else:
            filtered_data = processed_data

        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Analisis_Pertumbuhan_Anak_{timestamp}.xlsx"

        # Export to Excel using the export function
        success, result = export_analisis_from_json(filtered_data, filename, app.config['EXPORT_FOLDER'])

        if success:
            # Return the generated file for download
            return send_file(result,
                           as_attachment=True,
                           download_name=filename,
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return jsonify({'error': f'Gagal membuat file export: {result}'}), 500

    except Exception as e:
        return jsonify({'error': f'Error during export: {str(e)}'}), 500

def filter_data_by_periods(data, selected_periods):
    """
    Filter data to only include selected periods
    """
    if not selected_periods:
        return data
    
    filtered_data = data.copy()
    filtered_children = []
    
    for child in data.get('children', []):
        filtered_child = child.copy()
        if 'measurements' in child:
            # Filter measurements by selected periods
            filtered_measurements = [
                m for m in child['measurements'] 
                if m.get('periode') in selected_periods
            ]
            filtered_child['measurements'] = filtered_measurements
        
        # Only include child if they have measurements after filtering
        if filtered_child.get('measurements'):
            filtered_children.append(filtered_child)
    
    filtered_data['children'] = filtered_children
    # Update totals
    filtered_data['total_children'] = len(filtered_children)
    if filtered_children:
        total_measurements = sum(len(child.get('measurements', [])) for child in filtered_children)
        filtered_data['total_periods'] = len(selected_periods)
    
    return filtered_data

@app.route('/check-export-data')
def check_export_data():
    """
    Check if there's data available for export
    """
    try:
        export_id = session.get('export_id')
        stats = {}
        upload_time = None
        data = None

        # Priority 1: Use session export_id to get server storage data
        if export_id and export_id in export_data_store:
            data = export_data_store[export_id]['data']
            upload_time = export_data_store[export_id]['upload_timestamp']
        # Priority 2: Use session data directly
        elif 'processed_data' in session:
            data = session['processed_data']
            upload_time = session.get('upload_timestamp')
        # Priority 3: Use most recent server data as last resort
        elif export_data_store:
            latest_export_id = max(export_data_store.keys(),
                                 key=lambda k: export_data_store[k].get('created_at', ''))
            if latest_export_id in export_data_store:
                data = export_data_store[latest_export_id]['data']
                upload_time = export_data_store[latest_export_id]['upload_timestamp']

        # Extract stats if we have data
        if data and isinstance(data, dict) and 'children' in data and data['children']:
            stats = {
                'total_children': data.get('total_children', len(data['children'])),
                'total_periods': data.get('total_periods', 0),
                'file_name': data.get('file_name', 'Unknown'),
                'format_type': data.get('format_type', 'Unknown')
            }

        return jsonify({
            'has_export_data': data is not None,
            'upload_timestamp': upload_time,
            'stats': stats,
            'server_storage_count': len(export_data_store),
            'session_export_id': export_id
        })

    except Exception as e:
        return jsonify({'error': f'Error checking export data: {str(e)}'}), 500

@app.route('/debug-session')
def debug_session():
    """
    Debug endpoint to check session and server storage data
    """
    try:
        session_info = {
            'session_keys': list(session.keys()),
            'has_processed_data': 'processed_data' in session,
            'has_export_id': 'export_id' in session,
            'session_id': getattr(session, '_sid', 'unknown'),
            'session_type': str(type(session)),
            'debug_info': {}
        }

        # Check session data
        if 'processed_data' in session:
            data = session['processed_data']
            session_info['debug_info']['session_data'] = {
                'data_type': str(type(data)),
                'is_dict': isinstance(data, dict),
                'data_keys': list(data.keys()) if isinstance(data, dict) else 'not_dict',
                'has_children': isinstance(data, dict) and 'children' in data,
                'children_count': len(data.get('children', [])) if isinstance(data, dict) and 'children' in data else 0,
                'total_children_field': data.get('total_children', 'missing') if isinstance(data, dict) else 'missing'
            }

        # Check server storage
        session_info['debug_info']['server_storage'] = {
            'total_stored_exports': len(export_data_store),
            'storage_keys': list(export_data_store.keys()),
            'storage_details': {}
        }

        if export_data_store:
            for export_id, storage_data in export_data_store.items():
                session_info['debug_info']['server_storage']['storage_details'][export_id] = {
                    'created_at': storage_data.get('created_at'),
                    'has_children': isinstance(storage_data.get('data', {}), dict) and 'children' in storage_data.get('data', {}),
                    'children_count': len(storage_data.get('data', {}).get('children', [])),
                    'upload_timestamp': storage_data.get('upload_timestamp')
                }

        return jsonify(session_info)
    except Exception as e:
        return jsonify({'error': f'Debug session error: {str(e)}'}), 500

@app.route('/debug-export-section')
def debug_export_section():
    """
    Debug endpoint to test export section visibility logic
    """
    try:
        # Simulate what happens after upload
        debug_info = {
            'step_1_upload_response': {
                'description': 'What /upload returns after successful file upload',
                'has_export_data_field': True,
                'data_structure_check': 'Success -> should trigger showExportSection()'
            },
            'step_2_frontend_logic': {
                'description': 'Frontend JavaScript logic from index.html',
                'displayBalitaGrowthResult_calls': 'showExportSection(data) if data.has_export_data',
                'line_in_template': 'Line 948-950 in displayBalitaGrowthResult()'
            },
            'step_3_showExportSection': {
                'description': 'showExportSection() function makes export section visible',
                'export_section_id': 'document.getElementById("exportSection")',
                'display_style': 'Should change from "none" to "block"'
            },
            'step_4_alternative_logic': {
                'description': 'Alternative check via /check-export-data on page load',
                'calls_checkExportData': 'window.addEventListener("load", () => { loadFiles(); checkExportData(); })',
                'page_load_logic': 'Line 683-686 in index.html'
            }
        }

        return jsonify(debug_info)
    except Exception as e:
        return jsonify({'error': f'Debug export section error: {str(e)}'}), 500

@app.route('/debug-create-test-data')
def debug_create_test_data():
    """
    Create test data in session for debugging
    """
    try:
        # Create minimal test data
        test_data = {
            'file_name': 'Test_File_Debug.xlsx',
            'format_type': 'Debug Format',
            'total_children': 2,
            'total_periods': 3,
            'validation': {'valid': True, 'errors': [], 'warnings': []},
            'children': [
                {
                    'nama_anak': 'Test Child 1',
                    'nik': '1234567890123456',
                    'measurements': [
                        {'periode': 'JAN 2024', 'berat_kg': 5.5, 'tinggi_cm': 65.0},
                        {'periode': 'FEB 2024', 'berat_kg': 5.8, 'tinggi_cm': 66.5}
                    ]
                },
                {
                    'nama_anak': 'Test Child 2',
                    'nik': '2345678901234567',
                    'measurements': [
                        {'periode': 'JAN 2024', 'berat_kg': 6.2, 'tinggi_cm': 68.0}
                    ]
                }
            ]
        }

        # Store in session
        session['processed_data'] = test_data
        session['upload_timestamp'] = datetime.now().isoformat()

        # Also store in server-side storage for testing
        export_id = str(uuid.uuid4())
        export_data_store[export_id] = {
            'data': test_data,
            'upload_timestamp': datetime.now().isoformat(),
            'created_at': datetime.now().isoformat()
        }
        session['export_id'] = export_id

        return jsonify({
            'success': True,
            'message': 'Test data created successfully',
            'test_data_created': test_data,
            'session_keys': list(session.keys()),
            'export_id': export_id,
            'debug_info': {
                'processed_data_stored': True,
                'server_storage_stored': True,
                'children_count': len(test_data['children']),
                'total_measurements': sum(len(child['measurements']) for child in test_data['children'])
            }
        })
    except Exception as e:
        return jsonify({'error': f'Debug create test data error: {str(e)}'}), 500

if __name__ == '__main__':
    # Debug logging for startup
    print("=== SiTrack Stunting Startup Debug ===")
    print(f"PORT environment variable: {os.environ.get('PORT', 'NOT_SET')}")
    print(f"HOST environment variable: {os.environ.get('HOST', 'NOT_SET')}")
    print(f"FLASK_ENV environment variable: {os.environ.get('FLASK_ENV', 'NOT_SET')}")
    print(f"RAILWAY_ENVIRONMENT environment variable: {os.environ.get('RAILWAY_ENVIRONMENT', 'NOT_SET')}")

    # Test directories
    print(f"Current working directory: {os.getcwd()}")
    print(f"Upload folder exists: {os.path.exists(app.config['UPLOAD_FOLDER'])}")
    print(f"Session folder exists: {os.path.exists(app.config['SESSION_FILE_DIR'])}")
    print(f"Data template exists: {os.path.exists('data_template')}")
    if os.path.exists('data_template'):
        print(f"Data template contents: {os.listdir('data_template')}")

    port = int(os.environ.get('PORT', 8080))  # Updated for Railway
    host = os.environ.get('HOST', '0.0.0.0')
    debug_mode = os.environ.get('FLASK_ENV', 'production') == 'development'

    print(f"Starting SiTrack Stunting on {host}:{port}")
    print(f"Debug mode: {debug_mode}")
    print(f"Environment: {os.environ.get('RAILWAY_ENVIRONMENT', 'development')}")
    print("=== Starting Flask App ===")
    app.run(debug=debug_mode, host=host, port=port)