from flask import Flask, request, send_file, render_template
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from io import BytesIO
import zipfile

app = Flask(__name__)

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/generate_pdfs', methods=['POST'])
def generate_pdfs():
    if 'file' not in request.files:
        return 'No file uploaded', 400

    file = request.files['file']
    df = pd.read_excel(file)

    # Garantir que os nomes das colunas estão no formato esperado
    df.columns = df.columns.str.strip()  # Remover espaços extras
    df.columns = df.columns.str.upper()  # Transformar para maiúsculas

    # Verificar se as colunas obrigatórias existem
    required_columns = ['BENEFICIARIO', 'EMPREENDIMENTO', 'UNIDADE', 'VALOR TOTAL']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return f"As seguintes colunas estão faltando no arquivo: {', '.join(missing_columns)}", 400

    pdf_files = []

    # Agrupar o DataFrame por 'BENEFICIARIO'
    grouped = df.groupby('BENEFICIARIO')

    for beneficiario, group in grouped:
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=landscape(A4))

        y_position = 550

        # Título do PDF (centralizado)
        c.setFont("Helvetica-Bold", 16)
        c.setFillColor(colors.black)
        title_width = c.stringWidth(f"Relatório do Beneficiário: {beneficiario}", "Helvetica-Bold", 16)
        c.drawString((A4[0] - title_width) / 2, y_position, f"Relatório do Beneficiário: {beneficiario}")
        y_position -= 40

        # Cabeçalho da tabela (com fundo cinza claro e bordas)
        c.setFont("Helvetica-Bold", 10)  # Aumentei um pouco o tamanho da fonte
        headers = ["Empreendimento", "Unidade", "Valor Total"]
        col_widths = [200, 150, 150]  # Ajustei as larguras das colunas
        x_position = 50

        c.setStrokeColor(colors.grey) # Cor da borda da tabela
        c.setLineWidth(1) # Largura da borda da tabela

        for i, header in enumerate(headers):
            c.setFillColor(colors.HexColor("#f0f0f0"))  # Cinza mais claro
            c.rect(x_position, y_position - 20, col_widths[i], 30, fill=1, stroke=1) # Adicionado stroke para a borda
            c.setFillColor(colors.black)
            c.drawCentredString(x_position + col_widths[i] / 2, y_position - 10, header) # Centralizado o texto no cabeçalho
            x_position += col_widths[i]

        y_position -= 30

        # Desenhando as linhas da tabela
        for idx, row in group.iterrows():
            if y_position < 100:
                c.showPage()
                y_position = 550
                c.setFont("Helvetica-Bold", 16)
                c.drawString(50, y_position, f"Relatório do Beneficiário: {beneficiario}")
                y_position -= 40
                c.setFont("Helvetica-Bold", 10)
                x_position = 50
                for i, header in enumerate(headers):
                    c.setFillColor(colors.HexColor("#f0f0f0"))
                    c.rect(x_position, y_position - 20, col_widths[i], 30, fill=1, stroke=1)
                    c.setFillColor(colors.black)
                    c.drawCentredString(x_position + col_widths[i] / 2, y_position - 10, header)
                    x_position += col_widths[i]
                y_position -= 30

            x_position = 50
            c.setFont("Helvetica", 10) # Aumentei um pouco o tamanho da fonte

            highlight_row = "TOTAL" in str(row['UNIDADE']).upper()

            # Colorindo a linha "TOTAL" (apenas a área da tabela)
            if highlight_row:
                c.setFillColor(colors.yellow)
                c.rect(50, y_position - 20, sum(col_widths), 30, fill=1, stroke=0) # Removido o stroke para não sobrepor a borda da tabela

            for i, value in enumerate([
                row['EMPREENDIMENTO'],
                row['UNIDADE'],
                row['VALOR TOTAL']
            ]):
                if isinstance(value, (float, int)):
                    value = f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                value = str(value)[:30]
                c.setFillColor(colors.black)
                c.drawCentredString(x_position + col_widths[i] / 2, y_position - 10, value) # Centralizado o texto nas células
                c.rect(x_position, y_position - 20, col_widths[i], 30, fill=0, stroke=1) # Adicionado stroke para a borda da célula
                x_position += col_widths[i]
            y_position -= 30

        c.showPage()
        c.save()

        buffer.seek(0)
        pdf_files.append((buffer, f"{beneficiario}.pdf"))

    # Criar um arquivo ZIP com todos os PDFs
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for pdf_buffer, filename in pdf_files:
            zip_file.writestr(filename, pdf_buffer.getvalue())
    zip_buffer.seek(0)

    return send_file(zip_buffer, as_attachment=True, download_name="pdfs.zip", mimetype='application/zip')

if __name__ == "__main__":
    app.run(debug=True)