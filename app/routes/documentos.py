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

@documentos_bp.route('/relatorio_entrega', methods=['GET', 'POST'])
def relatorio_entrega():
    if request.method == 'POST':
        try:
            autor = request.form.get('autor', '').strip()
            data_servidor = request.form.get('data_servidor', '').strip()
            proposta_file = request.files.get('proposta_file')
            hu_files = request.files.getlist('hu_files')

            if not autor:
                return jsonify({'success': False, 'error': 'Campo Autor é obrigatório.'}), 400
            if not proposta_file or not proposta_file.filename:
                return jsonify({'success': False, 'error': 'Upload da proposta é obrigatório.'}), 400
            if not proposta_file.filename.lower().endswith('.docx'):
                return jsonify({'success': False, 'error': 'A proposta deve ser um arquivo .docx.'}), 400
            if not hu_files or all(not file.filename for file in hu_files):
                return jsonify({'success': False, 'error': 'É necessário enviar pelo menos uma HU (PDF).'}), 400
            invalid_hu = [file.filename for file in hu_files if file.filename and not file.filename.lower().endswith('.pdf')]
            if invalid_hu:
                return jsonify({'success': False, 'error': 'As HUs devem ser arquivos .pdf.'}), 400

            service = DocxService()
            filepath, _, _ = service.generate_relatorio_entrega(autor, data_servidor, proposta_file, hu_files)
            return jsonify({'success': True, 'filename': os.path.basename(filepath)})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400

    now = datetime.now().strftime('%d/%m/%Y')
    return render_template('relatorio_entrega.html', now=now)
