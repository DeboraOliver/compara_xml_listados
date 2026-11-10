import os
from playwright.sync_api import sync_playwright

ARQUIVO_IDS = "ids.txt"
PASTA_DOWNLOADS = "downloads_xml"
ARQUIVO_SESSAO = "state.json"
CNPJ_PADRAO = "45358058000140"

os.makedirs(PASTA_DOWNLOADS, exist_ok=True)

def carregar_ids(caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return []
    with open(caminho_arquivo, "r", encoding="utf-8") as f:
        return [linha.strip() for linha in f if linha.strip()]

def main():
    lista_ids = carregar_ids(ARQUIVO_IDS)
    if not lista_ids:
        print("Nenhum ID para processar.")
        return

    # if not os.path.exists(ARQUIVO_SESSAO):
    #     print(f"Erro: O arquivo de sessão '{ARQUIVO_SESSAO}' não existe.")
    #     print("Execute primeiro o script 'salvar_sessao.py' para autenticar e gerar a sessão.")
    #     return

    with sync_playwright() as p:
        # Abre o navegador carregando a sessão salva (pode rodar em headless=True se desejar)
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            # storage_state=ARQUIVO_SESSAO,
            accept_downloads=True
        )
        page = context.new_page()

        print(f"Iniciando o download de {len(lista_ids)} arquivo(s)...")

        for idx, id_evento in enumerate(lista_ids, start=1):
            url_consulta = (
                f"https://gerenciadoresocial.sigepe.gov.br/"
                f"sigepe-gerenciadoresocial-web/private/consultar-eventos/consultar/"
                f"{CNPJ_PADRAO}/{id_evento}"
            )

            print(f"[{idx}/{len(lista_ids)}] Processando ID: {id_evento}")

            try:
                # Dispara a navegação para a URL direta aguardando o download
                with page.expect_download(timeout=15000) as download_info:
                    page.goto(url_consulta)
                
                download = download_info.value
                caminho_salvar = os.path.join(PASTA_DOWNLOADS, f"{id_evento}.xml")
                download.save_as(caminho_salvar)
                print(f"   └─ Salvo com sucesso: {caminho_salvar}")

            except Exception:
                # Fallback caso a página carregue e precise de clique no botão de download
                try:
                    with page.expect_download(timeout=10000) as download_info:
                        page.locator("text=Download XML").or_(page.locator("a:has-text('XML')")).click()
                    
                    download = download_info.value
                    caminho_salvar = os.path.join(PASTA_DOWNLOADS, f"{id_evento}.xml")
                    download.save_as(caminho_salvar)
                    print(f"   └─ Salvo via clique: {caminho_salvar}")
                except Exception as e_inner:
                    print(f"   └─ Falha ao baixar o XML do ID {id_evento}: {e_inner}")

        print("\nProcesso concluído!")
        browser.close()

if __name__ == "__main__":
    main()