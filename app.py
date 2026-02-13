from flask import Flask, render_template, request, send_file
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

def limpar_tag(tag):
    return tag.split('}')[-1]

def extrair_dados_da_tag(xml_content, filename, target_tag):
    root = ET.fromstring(xml_content)
    dados_extraidos = []

    for elem in root.iter():
        if limpar_tag(elem.tag).lower() == target_tag.lower():
            registro = {"Arquivo_Origem": filename}
            for filho in elem.iter():
                if filho != elem:
                    registro[limpar_tag(filho.tag)] = filho.text.strip() if filho.text else ""
            dados_extraidos.append(registro)

    return dados_extraidos

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        tag_alvo = request.form["tag"]
        arquivos = request.files.getlist("arquivo")

        base_dados = []

        for arquivo in arquivos:
            nome = secure_filename(arquivo.filename)
            conteudo = arquivo.read()

            if nome.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
                    for info in z.infolist():
                        if info.filename.lower().endswith(".xml"):
                            with z.open(info) as f:
                                base_dados.extend(
                                    extrair_dados_da_tag(f.read(), info.filename, tag_alvo)
                                )
            elif nome.lower().endswith(".xml"):
                base_dados.extend(
                    extrair_dados_da_tag(conteudo, nome, tag_alvo)
                )

        if base_dados:
            df = pd.DataFrame(base_dados)
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            return send_file(
                output,
                download_name=f"consolidado_{tag_alvo}.xlsx",
                as_attachment=True
            )

    return render_template("index.html")

if __name__ == "__main__":
    app.run()
