"""Supabase client service."""
import os
from supabase import create_client, Client
from app.config import Config


_supabase_client = None


def get_supabase_client() -> Client:
    """Get or create Supabase client."""
    global _supabase_client

    if _supabase_client is None:
        url = Config.SUPABASE_URL
        key = Config.SUPABASE_KEY

        if not url or not key:
            raise ValueError("Supabase URL and Key must be configured")

        _supabase_client = create_client(url, key)

    return _supabase_client


def save_case(patient_info, totals):
    """
    Save case to Supabase.

    Args:
        patient_info: Patient information dictionary
        totals: Cost totals dictionary

    Returns:
        Created case record
    """
    client = get_supabase_client()

    case_data = {
        'patient_name': patient_info.get('patient_name'),
        'date_of_birth': str(patient_info.get('date_of_birth')) if patient_info.get('date_of_birth') else None,
        'date_of_injury': str(patient_info.get('date_of_injury')) if patient_info.get('date_of_injury') else None,
        'date_of_report': str(patient_info.get('date_of_report')) if patient_info.get('date_of_report') else None,
        'life_expectancy': float(patient_info.get('life_expectancy', 0) or 0),
        'age_initiated': int(patient_info.get('age_initiated', 0) or 0) if patient_info.get('age_initiated') else None,
        'geographic_multiplier': float(patient_info.get('geographic_multiplier', 1.0) or 1.0),
        'city_state': patient_info.get('city_state'),
        'zipcode': patient_info.get('zipcode'),
        'referring_attorney': patient_info.get('referring_attorney'),
        'total_annual_cost': totals['total_annual'],
        'total_one_time_cost': totals['total_one_time'],
        'lifetime_annual_cost': totals['lifetime_annual'],
        'grand_total': totals['grand_total'],
        'status': 'completed',
    }

    result = client.table('cases').insert(case_data).execute()
    return result.data[0] if result.data else None


def save_case_items(case_id, items):
    """
    Save case items to Supabase.

    Args:
        case_id: UUID of the parent case
        items: List of item dictionaries

    Returns:
        Created item records
    """
    client = get_supabase_client()

    items_data = []
    for idx, item in enumerate(items):
        items_data.append({
            'case_id': case_id,
            'category': item.get('category', ''),
            'item_name': item.get('item', '') or item.get('service_description', ''),
            'subcategory': item.get('subcategory'),
            'service_description': item.get('service_description'),
            'code_type': item.get('code_type'),
            'cpt_codes': item.get('code'),
            'unit_cost': item.get('unit_cost', 0),
            'frequency': item.get('frequency'),
            'annual_cost': item.get('annual_cost', 0),
            'one_time_cost': item.get('one_time_cost', 0),
            'rationale': item.get('rationale'),
            'source': item.get('source'),
            'sort_order': idx,
        })

    if items_data:
        result = client.table('case_items').insert(items_data).execute()
        return result.data
    return []


def save_document_metadata(case_id, file_name, storage_path, file_size, doc_type='lcp_recommendations'):
    """
    Save document metadata to Supabase.

    Args:
        case_id: UUID of the parent case
        file_name: Name of the file
        storage_path: Path in Supabase storage
        file_size: Size of file in bytes
        doc_type: Type of document

    Returns:
        Created document record
    """
    client = get_supabase_client()

    doc_data = {
        'case_id': case_id,
        'document_type': doc_type,
        'file_name': file_name,
        'storage_path': storage_path,
        'file_size': file_size,
        'mime_type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }

    result = client.table('documents').insert(doc_data).execute()
    return result.data[0] if result.data else None


def upload_file_to_storage(file_path, storage_path):
    """
    Upload file to Supabase Storage.

    Args:
        file_path: Local file path
        storage_path: Destination path in storage bucket

    Returns:
        Storage path on success
    """
    client = get_supabase_client()

    with open(file_path, 'rb') as f:
        file_data = f.read()

    result = client.storage.from_('lcp-documents').upload(
        storage_path,
        file_data,
        {'content-type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}
    )

    return storage_path


def get_cases(limit=50, offset=0):
    """Get list of cases."""
    client = get_supabase_client()

    result = client.table('cases').select('*').order('created_at', desc=True).range(offset, offset + limit - 1).execute()
    return result.data


def get_case(case_id):
    """Get case by ID."""
    client = get_supabase_client()

    result = client.table('cases').select('*').eq('id', case_id).single().execute()
    return result.data


def get_case_items(case_id):
    """Get items for a case."""
    client = get_supabase_client()

    result = client.table('case_items').select('*').eq('case_id', case_id).order('sort_order').execute()
    return result.data


def get_documents(case_id):
    """Get documents for a case."""
    client = get_supabase_client()

    result = client.table('documents').select('*').eq('case_id', case_id).order('created_at', desc=True).execute()
    return result.data


def get_download_url(storage_path):
    """Get signed URL for file download."""
    client = get_supabase_client()

    result = client.storage.from_('lcp-documents').create_signed_url(storage_path, 3600)  # 1 hour expiry
    return result.get('signedURL')
