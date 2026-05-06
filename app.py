"""
DD1750 Converter Web Application

A Flask application that converts GCSS-Army BOM PDFs to DD Form 1750 Packing Lists.
Includes a review/edit interface for users to verify and modify extracted data.
"""

import os
import json
import tempfile
import uuid
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify, session
from werkzeug.utils import secure_filename

from dd1750_core import (
    extract_items_from_pdf,
    generate_dd1750_from_items,
    BomItem,
    BomFormat,
    HeaderInfo
)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

# Configuration
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# In-memory storage for extraction results (use Redis/DB in production)
extraction_cache = {}


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_template_path():
    """Get path to DD1750 template."""
    # Look for template in multiple locations
    possible_paths = [
        os.path.join(os.path.dirname(__file__), 'templates', 'blank_1750.pdf'),
        os.path.join(os.path.dirname(__file__), 'static', 'blank_1750.pdf'),
        os.path.join(os.path.dirname(__file__), 'blank_1750.pdf'),
        '/app/blank_1750.pdf',
        '/app/templates/blank_1750.pdf',
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    raise FileNotFoundError("DD1750 template not found")


@app.route('/')
def index():
    """Main page with upload form."""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_bom():
    """
    Upload and extract items from BOM PDF.
    Returns extraction results for review.
    """
    if 'bom_file' not in request.files:
        return jsonify({'error': 'No BOM file provided'}), 400
    
    bom_file = request.files['bom_file']
    
    if bom_file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(bom_file.filename):
        return jsonify({'error': 'Invalid file type. Please upload a PDF.'}), 400
    
    try:
        # Save uploaded file temporarily
        filename = secure_filename(bom_file.filename)
        session_id = str(uuid.uuid4())
        bom_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{session_id}_bom.pdf")
        bom_file.save(bom_path)
        
        # Get start page from form
        start_page = int(request.form.get('start_page', 0))
        
        # Extract items
        result = extract_items_from_pdf(bom_path, start_page)
        
        # Store in cache for later use
        extraction_cache[session_id] = {
            'bom_path': bom_path,
            'result': result,
            'created_at': datetime.now().isoformat()
        }
        
        # Prepare response data
        items_data = [
            {
                'line_no': item.line_no,
                'description': item.description,
                'nsn': item.nsn,
                'qty': item.qty,
                'unit_of_issue': item.unit_of_issue,
                'oh_qty': item.oh_qty,  # -1 = not specified, 0 = zero, >0 = has qty
            }
            for item in result.items
        ]
        
        metadata_data = {
            'end_item_niin': result.metadata.end_item_niin,
            'end_item_description': result.metadata.end_item_description,
            'lin': result.metadata.lin,
            'serial_equip_no': result.metadata.serial_equip_no,
            'uic': result.metadata.uic,
            'format_detected': result.format_detected.value,
        }
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'items': items_data,
            'metadata': metadata_data,
            'item_count': len(result.items),
            'pages_processed': result.pages_processed,
            'warnings': result.warnings,
            'errors': result.errors,
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/generate', methods=['POST'])
def generate_dd1750():
    """
    Generate DD1750 PDF from reviewed/edited items.
    """
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        session_id = data.get('session_id')
        items_data = data.get('items', [])
        header_data = data.get('header', {})
        
        if not items_data:
            return jsonify({'error': 'No items to generate'}), 400
        
        # Convert JSON data back to BomItem objects
        items = []
        for i, item_data in enumerate(items_data):
            item = BomItem(
                line_no=i + 1,
                description=item_data.get('description', ''),
                nsn=item_data.get('nsn', ''),
                qty=int(item_data.get('qty', 1)),
                unit_of_issue=item_data.get('unit_of_issue', 'EA'),
            )
            items.append(item)
        
        # Create HeaderInfo from form data
        header = HeaderInfo(
            packed_by=header_data.get('packed_by', ''),
            num_boxes=header_data.get('num_boxes', '1'),
            requisition_no=header_data.get('requisition_no', ''),
            order_no=header_data.get('order_no', ''),
            end_item=header_data.get('end_item', ''),
            date=header_data.get('date', ''),
        )
        
        # Generate PDF
        template_path = get_template_path()
        output_path = os.path.join(
            app.config['UPLOAD_FOLDER'],
            f"{session_id or uuid.uuid4()}_dd1750.pdf"
        )
        
        output_path, count = generate_dd1750_from_items(items, template_path, output_path, header)
        
        # Clean up cache if session exists
        if session_id and session_id in extraction_cache:
            cache_data = extraction_cache.pop(session_id)
            if os.path.exists(cache_data.get('bom_path', '')):
                try:
                    os.remove(cache_data['bom_path'])
                except:
                    pass
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name='DD1750.pdf',
            mimetype='application/pdf'
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/quick-generate', methods=['POST'])
def quick_generate():
    """
    Quick generation without review step.
    For users who want to skip the review interface.
    """
    if 'bom_file' not in request.files:
        return "No BOM file provided", 400
    
    bom_file = request.files['bom_file']
    template_file = request.files.get('template_file')
    
    if bom_file.filename == '':
        return "No file selected", 400
    
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save BOM
            bom_path = os.path.join(tmpdir, 'bom.pdf')
            bom_file.save(bom_path)
            
            # Get or use default template
            if template_file and template_file.filename:
                tpl_path = os.path.join(tmpdir, 'template.pdf')
                template_file.save(tpl_path)
            else:
                tpl_path = get_template_path()
            
            # Output path
            out_path = os.path.join(tmpdir, 'dd1750.pdf')
            
            # Get start page
            start_page = int(request.form.get('start_page', 0))
            
            # Extract and generate
            result = extract_items_from_pdf(bom_path, start_page)
            
            if not result.items:
                return "No items found in BOM. Ensure this is a GCSS-Army format BOM.", 400
            
            out_path, count = generate_dd1750_from_items(result.items, tpl_path, out_path)
            
            return send_file(
                out_path,
                as_attachment=True,
                download_name='DD1750.pdf',
                mimetype='application/pdf'
            )
            
    except Exception as e:
        return f"Error: {str(e)}", 500


@app.route('/api/formats')
def get_supported_formats():
    """Return information about supported BOM formats."""
    return jsonify({
        'supported_formats': [
            {
                'name': 'GCSS-Army Component Listing',
                'description': 'Standard Component Listing / Hand Receipt with LV column',
                'identifier': BomFormat.GCSS_ARMY_STANDARD.value,
            },
            {
                'name': 'Equipment Property Record',
                'description': 'EPP format BOM',
                'identifier': BomFormat.EPP_FORMAT.value,
            }
        ],
        'note': 'Handwritten BOMs are not supported. Please obtain clean digital BOMs from GCSS-Army.',
    })


@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    return jsonify({'error': 'File too large. Maximum size is 16MB.'}), 413


@app.errorhandler(500)
def server_error(e):
    """Handle internal server errors."""
    return jsonify({'error': 'Internal server error'}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
