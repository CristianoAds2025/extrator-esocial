from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "sua_chave_secreta_aqui"
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024


def limpar_tag(tag):
    return tag.split('}')[-1]


def extrair_dados_da_tag(xml_content, filename, target_tag):
    dados_extraidos = []

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        flash(f"Erro ao processar o XML do arquivo '{filename}'. Arquivo inválido.", "danger")
        return []

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

        tag_alvo = request.form.get("tag")
        arquivos = request.files.getlist("arquivo")

        if not tag_alvo:
            flash("Informe a tag que deseja extrair.", "warning")
            return redirect(url_for("index"))

        if not arquivos or arquivos[0].filename == "":
            flash("Nenhum arquivo foi enviado.", "warning")
            return redirect(url_for("index"))

        base_dados = []

        for arquivo in arquivos:
            nome = secure_filename(arquivo.filename)
            conteudo = arquivo.read()

            try:
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

                else:
                    flash(f"O arquivo '{nome}' não é XML nem ZIP válido.", "danger")

            except zipfile.BadZipFile:
                flash(f"O arquivo '{nome}' está corrompido ou não é um ZIP válido.", "danger")

            except Exception as e:
                flash(f"Erro inesperado ao processar '{nome}': {str(e)}", "danger")

        if base_dados:
            df = pd.DataFrame(base_dados)
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            flash(f"{len(base_dados)} registro(s) encontrado(s) para a tag '{tag_alvo}'.", "success")

            return send_file(
                output,
                download_name=f"consolidado_{tag_alvo}.xlsx",
                as_attachment=True
            )

        else:
            flash(f"Nenhum dado encontrado para a tag '{tag_alvo}' nos arquivos enviados.", "warning")
            return redirect(url_for("index"))

    return render_template("index.html")


if __name__ == "__main__":
    app.run()

