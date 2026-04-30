from datetime import datetime
import os
import shutil
import uuid
from pathlib import Path

import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from flask import current_app
from werkzeug.utils import secure_filename

class DocxService:
    def _resolve_template_path(self, primary_key, legacy_key):
        primary = current_app.config.get(primary_key)
        legacy = current_app.config.get(legacy_key)
        if primary and os.path.exists(primary):
            return primary
        if legacy and os.path.exists(legacy):
            return legacy
        missing = [p for p in [primary, legacy] if p]
        raise FileNotFoundError(
            f"Template não encontrado. Verifique um destes caminhos: {', '.join(missing)}"
        )

    def _insert_after_paragraph(self, paragraph, text, align=None):
        new_p_elm = OxmlElement('w:p')
        paragraph._p.addnext(new_p_elm)
        new_p = Paragraph(new_p_elm, paragraph._parent)
        if text:
            new_p.add_run(text)
        if align is not None:
            new_p.alignment = align
        return new_p

    def _ensure_runtime_folders(self):
        folders = [
            current_app.config['UPLOAD_FOLDER'],
            current_app.config['UPLOAD_PROPOSTAS_FOLDER'],
            current_app.config['UPLOAD_HUS_FOLDER'],
            current_app.config['UPLOAD_RELATORIOS_FOLDER'],
        ]
        for folder in folders:
            Path(folder).mkdir(parents=True, exist_ok=True)

    def _new_execution_paths(self):
        execution_id = datetime.now().strftime('%Y%m%d%H%M%S') + '_' + uuid.uuid4().hex[:8]
        proposta_dir = Path(current_app.config['UPLOAD_PROPOSTAS_FOLDER']) / execution_id
        hu_dir = Path(current_app.config['UPLOAD_HUS_FOLDER']) / execution_id
        relatorio_dir = Path(current_app.config['UPLOAD_RELATORIOS_FOLDER']) / execution_id
        proposta_dir.mkdir(parents=True, exist_ok=True)
        hu_dir.mkdir(parents=True, exist_ok=True)
        relatorio_dir.mkdir(parents=True, exist_ok=True)
        return proposta_dir, hu_dir, relatorio_dir

    def _replace_text_preserve_format(self, paragraph, key, value):
        if key not in paragraph.text:
            return
        if paragraph.runs:
            full_text = paragraph.text.replace(key, value)
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ''
            paragraph.runs[0].font.name = 'Inter'
        else:
            paragraph.text = paragraph.text.replace(key, value)

    def _replace_placeholders_everywhere(self, doc, replacements):
        for p in doc.paragraphs:
            for key, value in replacements.items():
                self._replace_text_preserve_format(p, key, value)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for key, value in replacements.items():
                            self._replace_text_preserve_format(p, key, value)

        for section in doc.sections:
            for header in [section.header, section.footer]:
                for p in header.paragraphs:
                    for key, value in replacements.items():
                        self._replace_text_preserve_format(p, key, value)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for key, value in replacements.items():
                                    self._replace_text_preserve_format(p, key, value)

    def _format_brl_value(self, value):
        """Formata valor numérico no padrão brasileiro sem prefixo de moeda."""
        formatted = f"{value:,.2f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")

    def _extract_intro_from_proposta_docx(self, proposta_path):
        proposta_doc = Document(str(proposta_path))
        paragraphs = [p.text.strip() for p in proposta_doc.paragraphs if p.text and p.text.strip()]
        if not paragraphs:
            return ''
        for idx, text in enumerate(paragraphs):
            lowered = text.casefold()
            if lowered == 'introducao' or lowered == 'introdução':
                if idx + 1 < len(paragraphs):
                    return paragraphs[idx + 1]
        return paragraphs[0]

    def _extract_pdf_text(self, pdf_path):
        full_text = []
        with fitz.open(str(pdf_path)) as pdf:
            for page in pdf:
                text = page.get_text('text').strip()
                if text:
                    full_text.append(text)
        return '\n'.join(full_text).strip()

    def _write_hu_item_9(self, doc, hu_files):
        anchor = None
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if '<Conteudo HUs Item 9>' in text or '{{HU_DETALHE_ITEM9}}' in text:
                paragraph.text = text.replace('<Conteudo HUs Item 9>', '').replace('{{HU_DETALHE_ITEM9}}', '')
                anchor = paragraph
                break
            if text.startswith('9. Anexos'):
                anchor = paragraph
                break

        if anchor is None:
            doc.add_page_break()
            anchor = doc.add_paragraph('9. Anexos')

        current = self._insert_after_paragraph(anchor, '')
        for index, hu_file in enumerate(hu_files, start=1):
            current = self._insert_after_paragraph(current, f'Anexo {index}: {hu_file.stem}')
            content = self._extract_pdf_text(hu_file)
            if content:
                for line in content.splitlines():
                    cleaned = line.strip()
                    if cleaned:
                        current = self._insert_after_paragraph(current, cleaned, WD_ALIGN_PARAGRAPH.JUSTIFY)
            else:
                current = self._insert_after_paragraph(current, 'Conteúdo não extraído do PDF.')
            if index < len(hu_files):
                current = self._insert_after_paragraph(current, '')
                run = current.add_run()
                run.add_break(WD_BREAK.PAGE)

    def _write_hu_list_item_41(self, doc, hu_files):
        hu_lines = [f"Anexo {i}: {hu_file.stem}" for i, hu_file in enumerate(hu_files, start=1)]
        text = '\n'.join(hu_lines) if hu_lines else 'Sem HUs enviadas.'
        for paragraph in doc.paragraphs:
            if '<Lista HUs>' in paragraph.text or '{{HU_LISTA}}' in paragraph.text:
                paragraph.text = paragraph.text.replace('<Lista HUs>', text).replace('{{HU_LISTA}}', text)
                for run in paragraph.runs:
                    run.font.name = 'Inter'
                return

        for paragraph in doc.paragraphs:
            if paragraph.text.strip().startswith('4.1.'):
                self._insert_after_paragraph(paragraph, text)
                return

    def _sanitize_upload_name(self, filename, fallback):
        clean = secure_filename(filename or '')
        return clean or fallback

    def generate_proposta_os(self, data):
        self._ensure_runtime_folders()
        template_path = self._resolve_template_path('TEMPLATE_DOCX', 'LEGACY_TEMPLATE_DOCX')

        doc = Document(template_path)

        # Dados para substituição
        ano_os = data.get('ano_os', '')
        num_os = data.get('num_os', '')
        titulo_os = data.get('titulo_os', '')
        nome_os = data.get('nome_os', '')
        tipo_os = data.get('tipo_os', '')
        nome_autor = data.get('nome_autor', '')
        nome_solicitante = data.get('nome_solicitante', '')
        descricao_geral = data.get('descricao_geral', '')
        qtd_hst = int(float(data.get('qtd_hst', 0))) # Formatação sem casas decimais
        itens = data.get('itens', [])
        data_atual = datetime.now().strftime('%d/%m/%Y')
        valor_calculado = self._format_brl_value(qtd_hst * 200)

        replacements = {
            '<Autor>': nome_autor,
            '<Nome da OS>': nome_os,
            '<Nome da Os>': nome_os,
            '<Tipo da OS>': tipo_os,
            '<Solicitante>': nome_solicitante,
            '<Nome do solicitante>': nome_solicitante,
            '<Descrição Geral da OS>': descricao_geral,
            '<Quantidade de HST>': str(qtd_hst),
            '<Data atual do servidor>': data_atual,
            '<Cálculo em R$ do valor de HST x 200,00>': valor_calculado
        }

        self._replace_placeholders_everywhere(doc, replacements)

        # Tratamento especial para <Itens da OS>
        for p in doc.paragraphs:
            if '<Itens da OS>' in p.text:
                itens_text = ""
                for i, item in enumerate(itens, 1):
                    itens_text += f"{i}. {item}\n"
                # Para itens, substituímos o texto e forçamos alinhamento, mas mantemos o estilo do parágrafo
                p.text = p.text.replace('<Itens da OS>', itens_text.strip())
                for run in p.runs:
                    run.font.name = 'Inter'
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if '<Itens da OS>' in cell.text:
                        itens_text = ""
                        for i, item in enumerate(itens, 1):
                            itens_text += f"{i}. {item}\n"
                        # Substituição na célula preservando parágrafos internos
                        for p in cell.paragraphs:
                            if '<Itens da OS>' in p.text:
                                p.text = p.text.replace('<Itens da OS>', itens_text.strip())
                                for run in p.runs:
                                    run.font.name = 'Inter'
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Nome do arquivo de saída
        filename = f"Proposta OS {ano_os}-{num_os} - {titulo_os}.docx"
        output_path = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        doc.save(output_path)

        return output_path

    def generate_relatorio_entrega(self, autor, proposta_file, hu_files):
        self._ensure_runtime_folders()
        template_path = self._resolve_template_path(
            'TEMPLATE_RELATORIO_ENTREGA_DOCX',
            'LEGACY_TEMPLATE_RELATORIO_ENTREGA_DOCX'
        )

        proposta_dir, hu_dir, relatorio_dir = self._new_execution_paths()
        proposta_filename = self._sanitize_upload_name(proposta_file.filename, 'proposta.docx')
        proposta_target = proposta_dir / proposta_filename
        proposta_file.save(str(proposta_target))

        saved_hus = []
        for hu_file in hu_files:
            if not hu_file or not hu_file.filename:
                continue
            hu_name = self._sanitize_upload_name(hu_file.filename, f'hu_{len(saved_hus)+1}.pdf')
            hu_target = hu_dir / hu_name
            hu_file.save(str(hu_target))
            saved_hus.append(hu_target)

        if not saved_hus:
            raise ValueError('Nenhuma HU válida foi enviada para gerar o relatório.')

        introducao = self._extract_intro_from_proposta_docx(proposta_target)
        data_atual = datetime.now().strftime('%d/%m/%Y')

        doc = Document(template_path)
        replacements = {
            '<Autor>': autor,
            '{{AUTOR}}': autor,
            '<Data atual do servidor>': data_atual,
            '{{DATA_ATUAL_SERVIDOR}}': data_atual,
            '<Introdução da Proposta>': introducao,
            '<Introducao da Proposta>': introducao,
            '{{INTRODUCAO_PROPOSTA}}': introducao,
        }
        self._replace_placeholders_everywhere(doc, replacements)
        self._write_hu_list_item_41(doc, saved_hus)
        self._write_hu_item_9(doc, saved_hus)

        report_title = f"Relatorio de Entrega - {datetime.now().strftime('%Y%m%d_%H%M%S')}"
        output_docx = relatorio_dir / f'{report_title}.docx'
        doc.save(str(output_docx))

        final_output = Path(current_app.config['UPLOAD_FOLDER']) / output_docx.name
        shutil.copy2(output_docx, final_output)
        return str(final_output), str(relatorio_dir), str(hu_dir)
