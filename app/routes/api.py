"""API endpoints for LCP generation."""
import os
import uuid
import tempfile
from flask import Blueprint, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename

from app.services.workbook_parser import (
    parse_workbook,
    WorkbookParseError,
    NoItemsSelectedError,
    MissingPatientInfoError
)
from app.services.cost_calculator import calculate_all_costs
from app.services.document_generator import generate_lcp_document
from app.services.supabase_client import (
    save_case,
    save_case_items,
    save_document_metadata,
    upload_file_to_storage,
    get_cases,
    get_case,
    get_case_items,
    get_documents,
    get_download_url
)

api_bp = Blueprint('api', __name__, url_prefix='/api')

ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}


def allowed_file(filename):
    """Check if file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@api_bp.route('/generate', methods=['POST'])
def generate_lcp():
    """
    Generate LCP recommendations from uploaded workbook.

    Expects multipart/form-data with:
        - file: Master Workbook (.xlsm/.xlsx)
        - medical_summary: Optional medical summary (.docx)

    Returns:
        JSON with case_id, document details, and summary
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xlsm file'}), 400

    # Save uploaded file temporarily
    temp_dir = tempfile.mkdtemp()
    filename = secure_filename(file.filename)
    workbook_path = os.path.join(temp_dir, filename)
    file.save(workbook_path)

    try:
        # Parse workbook
        workbook_data = parse_workbook(workbook_path)

        # Calculate costs
        cost_data = calculate_all_costs(workbook_data)

        # Generate document
        patient_name = workbook_data['patient_info'].get('patient_name', 'Unknown')
        safe_name = "".join(c for c in patient_name if c.isalnum() or c in ' -_').strip()
        doc_filename = f"LCP_Recommendations_{safe_name}_{uuid.uuid4().hex[:8]}.docx"
        doc_path = os.path.join(temp_dir, doc_filename)

        generate_lcp_document(
            workbook_data['patient_info'],
            cost_data,
            doc_path
        )

        # Try to save to Supabase (optional - works without it)
        case_id = None
        try:
            case = save_case(workbook_data['patient_info'], cost_data['totals'])
            if case:
                case_id = case['id']
                save_case_items(case_id, cost_data['items'])

                # Upload document to storage
                storage_path = f"cases/{case_id}/{doc_filename}"
                upload_file_to_storage(doc_path, storage_path)
                save_document_metadata(
                    case_id,
                    doc_filename,
                    storage_path,
                    os.path.getsize(doc_path)
                )
        except Exception as e:
            # Supabase save failed, but we can still return the document
            current_app.logger.warning(f"Supabase save failed: {str(e)}")

        # Return the generated document
        return send_file(
            doc_path,
            as_attachment=True,
            download_name=doc_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except NoItemsSelectedError as e:
        return jsonify({'error': str(e)}), 400
    except MissingPatientInfoError as e:
        return jsonify({'error': str(e)}), 400
    except WorkbookParseError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        current_app.logger.error(f"Error generating LCP: {str(e)}")
        return jsonify({'error': f'Error processing workbook: {str(e)}'}), 500
    finally:
        # Cleanup temp files
        try:
            if os.path.exists(workbook_path):
                os.remove(workbook_path)
        except Exception:
            pass


@api_bp.route('/preview', methods=['POST'])
def preview_lcp():
    """
    Preview LCP data without generating document.

    Returns cost calculations and summary without saving.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']

    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file'}), 400

    temp_dir = tempfile.mkdtemp()
    filename = secure_filename(file.filename)
    workbook_path = os.path.join(temp_dir, filename)
    file.save(workbook_path)

    try:
        workbook_data = parse_workbook(workbook_path)
        cost_data = calculate_all_costs(workbook_data)

        return jsonify({
            'patient_info': {
                'patient_name': workbook_data['patient_info'].get('patient_name'),
                'date_of_birth': str(workbook_data['patient_info'].get('date_of_birth', '')),
                'date_of_injury': str(workbook_data['patient_info'].get('date_of_injury', '')),
                'life_expectancy': workbook_data['patient_info'].get('life_expectancy'),
            },
            'totals': cost_data['totals'],
            'categories': {
                cat: {
                    'annual_cost': data['annual_cost'],
                    'one_time_cost': data['one_time_cost'],
                    'item_count': len(data['items'])
                }
                for cat, data in cost_data['category_totals'].items()
            },
            'item_count': len(cost_data['items']),
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 400
    finally:
        try:
            if os.path.exists(workbook_path):
                os.remove(workbook_path)
        except Exception:
            pass


@api_bp.route('/cases', methods=['GET'])
def list_cases():
    """List all cases."""
    try:
        limit = request.args.get('limit', 50, type=int)
        offset = request.args.get('offset', 0, type=int)
        cases = get_cases(limit, offset)
        return jsonify({'cases': cases})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@api_bp.route('/cases/<case_id>', methods=['GET'])
def get_case_detail(case_id):
    """Get case details by ID."""
    try:
        case = get_case(case_id)
        if not case:
            return jsonify({'error': 'Case not found'}), 404

        items = get_case_items(case_id)
        documents = get_documents(case_id)

        return jsonify({
            'case': case,
            'items': items,
            'documents': documents
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@api_bp.route('/documents/<case_id>/download', methods=['GET'])
def download_document(case_id):
    """Get download URL for case document."""
    try:
        documents = get_documents(case_id)
        if not documents:
            return jsonify({'error': 'No documents found'}), 404

        # Get most recent document
        doc = documents[0]
        download_url = get_download_url(doc['storage_path'])

        return jsonify({
            'download_url': download_url,
            'file_name': doc['file_name']
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
