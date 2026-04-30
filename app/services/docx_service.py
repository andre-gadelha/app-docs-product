from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from flask import current_app
from datetime import datetime
import os

class DocxService:
    def _replace_text_preserve_format(self, paragraph, key, value):
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)
                    run.font.name = 'Inter'

    def _format_brl_value(self, value):
        """Formata valor numérico no padrão brasileiro sem prefixo de moeda."""
        formatted = f"{value:,.2f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")

    def generate_proposta_os(self, data):
        template_path = current_app.config['TEMPLATE_DOCX']
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template não encontrado: {template_path}")

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

        # Substituição nos parágrafos (Preservando formatação via runs)
        for p in doc.paragraphs:
            for key, value in replacements.items():
                self._replace_text_preserve_format(p, key, value)

        # Substituição em tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for key, value in replacements.items():
                            self._replace_text_preserve_format(p, key, value)

        # Substituição em Cabeçalhos e Rodapés
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
