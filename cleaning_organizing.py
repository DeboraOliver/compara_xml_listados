from collections import Counter
from datetime import datetime
from pathlib import Path
import re
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


PASTA_XML = Path(__file__).resolve().parent / "downloads_xml"
ARQUIVO_EXCEL = Path(__file__).resolve().parent / "compara_dados.xlsx"


def nome_local(tag):
    """Remove o namespace de uma tag XML."""
    return tag.rsplit("}", 1)[-1]


def extrair_data_geracao(nome_arquivo):
    correspondencia = re.search(r"ID\d{15}(\d{8})", nome_arquivo)
    if correspondencia is None:
        raise ValueError(f"Não foi possível encontrar a data no ID do arquivo {nome_arquivo}")
    try:
        return datetime.strptime(correspondencia.group(1), "%Y%m%d").date()
    except ValueError as erro:
        raise ValueError(f"Data inválida no nome do arquivo {nome_arquivo}") from erro


def extrair_campos(elemento, caminho=(), segmento=None):
    """Retorna valores e atributos usando caminhos completos como nomes de coluna."""
    campos = {}
    nome = segmento or nome_local(elemento.tag)
    caminho_atual = caminho + (nome,)
    prefixo = ".".join(caminho_atual)

    for atributo, valor in elemento.attrib.items():
        campos[f"{prefixo}.@{nome_local(atributo)}"] = valor

    filhos = list(elemento)
    if not filhos:
        campos[prefixo] = (elemento.text or "").strip()
        return campos

    texto = (elemento.text or "").strip()
    if texto:
        campos[prefixo] = texto

    ocorrencias = Counter(nome_local(filho.tag) for filho in filhos)
    vistos = Counter()
    for filho in filhos:
        nome_filho = nome_local(filho.tag)
        vistos[nome_filho] += 1
        sufixo = f"[{vistos[nome_filho]}]" if ocorrencias[nome_filho] > 1 else ""
        campos.update(extrair_campos(filho, caminho_atual, f"{nome_filho}{sufixo}"))

    return campos


def extrair_evt_tab_rubrica(caminho_xml):
    raiz = ET.parse(caminho_xml).getroot()
    elemento = next(
        (elemento for elemento in raiz.iter() if nome_local(elemento.tag) == "evtTabRubrica"),
        None,
    )
    if elemento is None:
        raise ValueError(f"O elemento <evtTabRubrica> não foi encontrado em {caminho_xml.name}")

    campos = {}
    for atributo, valor in elemento.attrib.items():
        campos[f"evtTabRubrica.@{nome_local(atributo)}"] = valor

    for filho in elemento:
        campos.update(extrair_campos(filho))
    return campos


def criar_planilha():
    arquivos_xml = sorted(PASTA_XML.glob("*.xml"))
    if not arquivos_xml:
        raise FileNotFoundError(f"Nenhum arquivo XML encontrado em: {PASTA_XML}")

    linhas = []
    colunas = ["arquivo", "data_geracao"]
    for caminho_xml in arquivos_xml:
        campos = extrair_evt_tab_rubrica(caminho_xml)
        linhas.append({
            "arquivo": caminho_xml.name,
            "data_geracao": extrair_data_geracao(caminho_xml.name),
            **campos,
        })
        for coluna in campos:
            if coluna not in colunas:
                colunas.append(coluna)

    workbook = Workbook()
    planilha = workbook.active
    planilha.title = "evtTabRubrica"
    planilha.append(colunas)
    for celula in planilha[1]:
        celula.font = Font(bold=True)

    for linha in linhas:
        planilha.append([linha.get(coluna, "") for coluna in colunas])

    for celula in planilha["B"][1:]:
        celula.number_format = "dd/mm/yyyy"

    planilha.freeze_panes = "A2"
    planilha.auto_filter.ref = planilha.dimensions
    for indice, coluna in enumerate(colunas, start=1):
        largura = min(max(len(coluna), *(len(str(linha.get(coluna, ""))) for linha in linhas)) + 2, 60)
        planilha.column_dimensions[get_column_letter(indice)].width = largura

    workbook.save(ARQUIVO_EXCEL)
    print(f"Planilha criada: {ARQUIVO_EXCEL}")
    print(f"Arquivos processados: {len(linhas)}")


if __name__ == "__main__":
    criar_planilha()