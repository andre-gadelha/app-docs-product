from flask import Blueprint, render_template, request, send_file, current_app, jsonify
from datetime import datetime
import os
from app.services.docx_service import DocxService

documentos_bp = Blueprint('documentos', __name__, url_prefix='/documentos')

@documentos_bp.route('/proposta_os', methods=['GET', 'POST'])
def proposta_os():
    if request.method == 'POST':
        try:
            data = request.json
            service = DocxService()
            filepath = service.generate_proposta_os(data)
            return jsonify({'success': True, 'filename': os.path.basename(filepath)})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400
            
    now = datetime.now().strftime('%d/%m/%Y')
    return render_template('proposta_os.html', now=now)

@documentos_bp.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "Arquivo não encontrado", 404

@documentos_bp.route('/relatorio_entrega')
def relatorio_entrega():
    return render_template('relatorio_entrega.html')
