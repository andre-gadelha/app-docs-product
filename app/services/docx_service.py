from datetime import datetime, timedelta
import os
import re
import shutil
import uuid
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from flask import current_app
from werkzeug.utils import secure_filename

class DocxService:
    XLSX_NS = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    def _resolve_template_path(self, config_key):
        template_path = current_app.config.get(config_key)
        if template_path and os.path.exists(template_path):
            return template_path
        raise FileNotFoundError(
            f"Template não encontrado no caminho configurado: {template_path}"
        )

    def _resolve_data_source_file(self):
        base_dir = Path(current_app.config['FONTE_DADOS_DIR'])
        if not base_dir.exists():
            raise FileNotFoundError(f"Pasta de fonte de dados não encontrada: {base_dir}")
        xlsx_files = sorted(base_dir.glob('*.xlsx'))
        if not xlsx_files:
            raise FileNotFoundError(f"Nenhuma planilha .xlsx encontrada em: {base_dir}")
        return xlsx_files[0]

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

    def _extract_proposta_identity(self, proposta_path, original_filename):
        filename = Path(original_filename or proposta_path.name).name
        filename_match = re.search(
            r"(?i)proposta\s*os\s*(\d{4})\s*-\s*([0-9]+)\s*-\s*(.+?)\.docx$",
            filename
        )
        if filename_match:
            return {
                'ano_os': filename_match.group(1),
                'os': filename_match.group(2).zfill(4),
                'assunto': filename_match.group(3).strip(),
            }

        proposta_doc = Document(str(proposta_path))
        all_text = '\n'.join([p.text for p in proposta_doc.paragraphs if p.text]).strip()
        text_match = re.search(r"(?i)(\d{4})\s*-\s*([0-9]{1,4})", all_text)
        if text_match:
            return {
                'ano_os': text_match.group(1),
                'os': text_match.group(2).zfill(4),
                'assunto': '',
            }
        raise ValueError(
            "Não foi possível identificar Ano OS e Número OS a partir da proposta enviada. "
            "Use um arquivo no padrão: Proposta OS 2026-0032 - Assunto.docx"
        )

    def _normalize_header(self, header):
        key = (header or '').strip().casefold()
        aliases = {
            'ano os': 'ano_os',
            'ano_os': 'ano_os',
            'os': 'os',
            'assunto': 'assunto',
            'início': 'inicio',
            'inicio': 'inicio',
            'fim': 'fim',
            'hsts': 'hst',
            'hst': 'hst',
            'valor': 'valor',
        }
        return aliases.get(key, key)

    def _read_xlsx_row_by_os(self, workbook_path, ano_os, os_number):
        with ZipFile(workbook_path) as archive:
            shared = []
            if 'xl/sharedStrings.xml' in archive.namelist():
                root = ET.fromstring(archive.read('xl/sharedStrings.xml'))
                for item in root.findall('a:si', self.XLSX_NS):
                    shared.append(''.join((t.text or '') for t in item.findall('.//a:t', self.XLSX_NS)))

            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            rows = []
            for row in sheet.findall('a:sheetData/a:row', self.XLSX_NS):
                values = {}
                for cell in row.findall('a:c', self.XLSX_NS):
                    ref = cell.attrib.get('r', 'A1')
                    col_letters = ''.join(ch for ch in ref if ch.isalpha())
                    index = 0
                    for ch in col_letters:
                        index = index * 26 + ord(ch.upper()) - 64
                    col_index = index - 1
                    cell_type = cell.attrib.get('t')
                    raw_node = cell.find('a:v', self.XLSX_NS)
                    raw = '' if raw_node is None or raw_node.text is None else raw_node.text
                    if cell_type == 'inlineStr':
                        value = ''.join((t.text or '') for t in cell.findall('.//a:t', self.XLSX_NS))
                    elif cell_type == 's' and raw:
                        value = shared[int(raw)]
                    else:
                        value = raw
                    values[col_index] = value
                rows.append(values)

        if not rows:
            raise ValueError("Planilha sem conteúdo.")

        max_col = max(rows[0].keys()) if rows[0] else -1
        headers = []
        for i in range(max_col + 1):
            headers.append(rows[0].get(i, '').strip())
        normalized = [self._normalize_header(h) for h in headers]

        for values in rows[1:]:
            row_data = {}
            for i, key in enumerate(normalized):
                row_data[key] = str(values.get(i, '')).strip()
            row_ano = row_data.get('ano_os', '')
            row_os = row_data.get('os', '').zfill(4)
            if row_ano == str(ano_os) and row_os == str(os_number).zfill(4):
                return row_data
        raise ValueError(f"OS {ano_os}-{str(os_number).zfill(4)} não encontrada na planilha.")

    def _excel_serial_to_date(self, value):
        if not value:
            return ''
        try:
            serial = float(str(value).replace(',', '.'))
            dt = datetime(1899, 12, 30) + timedelta(days=serial)
            return dt.strftime('%d/%m/%Y')
        except Exception:
            return str(value)

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

    def _sanitize_output_filename(self, text):
        sanitized = re.sub(r'[<>:"/\\|?*]', '-', text).strip().rstrip('.')
        return sanitized or 'Relatório de Entrega'

    def generate_proposta_os(self, data):
        self._ensure_runtime_folders()
        template_path = self._resolve_template_path('TEMPLATE_DOCX')

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

    def generate_relatorio_entrega(self, autor, data_servidor, proposta_file, hu_files):
        self._ensure_runtime_folders()
        template_path = self._resolve_template_path('TEMPLATE_RELATORIO_ENTREGA_DOCX')

        proposta_dir, hu_dir, relatorio_dir = self._new_execution_paths()
        proposta_filename = self._sanitize_upload_name(proposta_file.filename, 'proposta.docx')
        proposta_target = proposta_dir / proposta_filename
        proposta_file.save(str(proposta_target))

        identity = self._extract_proposta_identity(proposta_target, proposta_file.filename)
        workbook = self._resolve_data_source_file()
        sheet_row = self._read_xlsx_row_by_os(workbook, identity['ano_os'], identity['os'])

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
        data_atual = data_servidor or datetime.now().strftime('%d/%m/%Y')
        assunto = identity['assunto'] or sheet_row.get('assunto', '')
        inicio = self._excel_serial_to_date(sheet_row.get('inicio', ''))
        fim = self._excel_serial_to_date(sheet_row.get('fim', ''))
        hst = sheet_row.get('hst', '')
        valor = self._format_brl_value(float(sheet_row.get('valor', '0') or 0))

        doc = Document(template_path)
        hu_lines = [f"Anexo {i}: {hu_file.stem}" for i, hu_file in enumerate(saved_hus, start=1)]
        replacements = {
            '<Autor>': autor,
            '{{AUTOR}}': autor,
            '{{Autor}}': autor,
            '<Data atual do servidor>': data_atual,
            '{{DATA_ATUAL_SERVIDOR}}': data_atual,
            '{{Data do Servidor}}': data_atual,
            '<Introdução da Proposta>': introducao,
            '<Introducao da Proposta>': introducao,
            '{{INTRODUCAO_PROPOSTA}}': introducao,
            '{{Replicar Introdução da Proposta da Ordem de Serviço}}': introducao,
            '{{Replicar Introducao da Proposta da Ordem de Servico}}': introducao,
            '{{ANO_OS}}': identity['ano_os'],
            '{{Ano_OS}}': identity['ano_os'],
            '{{OS}}': identity['os'],
            '{{Assunto}}': assunto,
            '{{Início}}': inicio,
            '{{Inicio}}': inicio,
            '{{Fim}}': fim,
            '{{HST}}': str(hst),
            '{{Valor}}': valor,
            '{{HU N}}': '\n'.join(hu_lines),
        }
        for idx in range(1, 21):
            replacements[f'{{{{HU {idx}}}}}'] = ''
        for idx, line in enumerate(hu_lines, start=1):
            replacements[f'{{{{HU {idx}}}}}'] = line
        self._replace_placeholders_everywhere(doc, replacements)
        self._write_hu_list_item_41(doc, saved_hus)
        self._write_hu_item_9(doc, saved_hus)

        report_title = self._sanitize_output_filename(
            f"Relatório de Entrega OS {identity['ano_os']}-{identity['os']} - {assunto}".strip()
        )
        output_docx = relatorio_dir / f'{report_title}.docx'
        doc.save(str(output_docx))

        final_output = Path(current_app.config['UPLOAD_FOLDER']) / output_docx.name
        shutil.copy2(output_docx, final_output)
        return str(final_output), str(relatorio_dir), str(hu_dir)
