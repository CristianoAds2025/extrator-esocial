from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import pandas as pd
import xml.etree.ElementTree as ET
import zipfile
import io
import os
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
        return []

    for elem in root.iter():
        if limpar_tag(elem.tag).lower() == target_tag.lower():

            filhos = list(elem)
            nomes_filhos = [limpar_tag(f.tag) for f in filhos]
            repetidos = {n for n in nomes_filhos if nomes_filhos.count(n) > 1}

            if repetidos:
                for filho in filhos:
                    registro = {"Arquivo_Origem": filename}
                    registro["Tag_Pai"] = target_tag

                    for sub in filho.iter():
                        if sub != filho:
                            registro[limpar_tag(sub.tag)] = sub.text.strip() if sub.text else ""

                    dados_extraidos.append(registro)
            else:
                registro = {"Arquivo_Origem": filename}
                for filho in filhos:
                    registro[limpar_tag(filho.tag)] = filho.text.strip() if filho.text else ""

                dados_extraidos.append(registro)

    return dados_extraidos


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        tag_alvo = request.form.get("tag")
        arquivos = request.files.getlist("arquivo")

        if not tag_alvo:
            return jsonify({"status": "erro", "mensagem": "Informe a tag que deseja extrair."})

        if not arquivos or arquivos[0].filename == "":
            return jsonify({"status": "erro", "mensagem": "Nenhum arquivo foi enviado."})

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
                    return jsonify({"status": "erro", "mensagem": f"O arquivo '{nome}' não é XML nem ZIP válido."})

            except zipfile.BadZipFile:
                return jsonify({"status": "erro", "mensagem": f"O arquivo '{nome}' está corrompido."})

            except Exception as e:
                return jsonify({"status": "erro", "mensagem": f"Erro inesperado ao processar '{nome}': {str(e)}"})

        if base_dados:

            df = pd.DataFrame(base_dados)

            arquivo_nome = f"consolidado_{tag_alvo}.xlsx"
            caminho = f"/tmp/{arquivo_nome}"

            df.to_excel(caminho, index=False)

            return jsonify({
                "status": "ok",
                "quantidade": len(base_dados),
                "arquivo": arquivo_nome
            })

        else:
            return jsonify({
                "status": "vazio",
                "mensagem": f"Nenhum dado encontrado para a tag '{tag_alvo}'."
            })

    return render_template("index.html")


@app.route("/download/<nome>")
def download(nome):
    caminho = f"/tmp/{nome}"
    if os.path.exists(caminho):
        return send_file(caminho, as_attachment=True)
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run()








