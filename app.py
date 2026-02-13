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

    def flatten_element(element, parent_path="", base_data=None):
        if base_data is None:
            base_data = {}

        children = list(element)

        # Se não tiver filhos → é folha
        if not children:
            coluna = parent_path
            base_data[coluna] = element.text.strip() if element.text else ""
            return [base_data]

        resultados = []

        # Agrupa filhos por nome para detectar repetição
        filhos_por_nome = {}
        for child in children:
            nome = limpar_tag(child.tag)
            filhos_por_nome.setdefault(nome, []).append(child)

        for nome, lista_filhos in filhos_por_nome.items():
            if len(lista_filhos) == 1:
                novo_path = f"{parent_path}_{nome}" if parent_path else nome
                resultados.extend(
                    flatten_element(lista_filhos[0], novo_path, base_data.copy())
                )
            else:
                # Caso repetido → explode linhas
                for item in lista_filhos:
                    novo_path = f"{parent_path}_{nome}" if parent_path else nome
                    resultados.extend(
                        flatten_element(item, novo_path, base_data.copy())
                    )

        return resultados

    for elem in root.iter():
        if limpar_tag(elem.tag).lower() == target_tag.lower():

            linhas = flatten_element(elem)

            for linha in linhas:
                linha["Arquivo_Origem"] = filename
                dados_extraidos.append(linha)

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



