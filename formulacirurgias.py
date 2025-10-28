import pandas as pd
import os

# 1️⃣ Caminho da pasta onde estão os arquivos Excel
pasta = "/Users/kristoferramina/phyton test"  # ⬅️ ajuste esse caminho conforme sua pasta
saida = os.path.join(pasta, "Resumo_Cirurgias.xlsx")

# 2️⃣ Listar todos os arquivos Excel
arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]

if not arquivos:
    print("⚠️ Nenhum arquivo .xlsx encontrado na pasta.")
else:
    print(f"📂 {len(arquivos)} arquivos encontrados:\n" + "\n".join(arquivos))

# 3️⃣ DataFrame que armazenará todos os resultados
resumo_total = pd.DataFrame()

# 4️⃣ Processar cada arquivo
for arquivo in arquivos:
    caminho = os.path.join(pasta, arquivo)
    print(f"\n📄 Processando: {arquivo}")

    try:
        # Ler as planilhas
        servicos = pd.read_excel(caminho, sheet_name="Guia de Serviços", header=None)
        extrato = pd.read_excel(caminho, sheet_name="Extrato", header=None)

        # Filtrar onde coluna J (índice 9) é diferente de "CON"
        filtro = servicos[servicos[9] != "CON"]

        # Pegar valores únicos da coluna B (índice 1)
        b_unicos = filtro[1].dropna().unique()

        # Criar DataFrame
        cirurgias = pd.DataFrame({"B": b_unicos})

        # Dicionários equivalentes ao XVERWEIS
        mapa_servicos = dict(zip(servicos[1], servicos[6]))   # B → G
        mapa_extrato = dict(zip(extrato[5], extrato[19]))     # F → T

        # Adicionar colunas A e C
        cirurgias["A"] = cirurgias["B"].map(mapa_servicos)
        cirurgias["C"] = cirurgias["B"].map(mapa_extrato)

        # Formatar valores da coluna C como moeda
        cirurgias["C"] = cirurgias["C"].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if pd.notnull(x)
            else ""
        )

        # Adicionar coluna com o nome do arquivo de origem
        cirurgias["Arquivo"] = arquivo

        # Adicionar ao resumo total
        resumo_total = pd.concat([resumo_total, cirurgias], ignore_index=True)

        print(f"✅ {arquivo} processado com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao processar {arquivo}: {e}")

# 5️⃣ Salvar o resumo consolidado em um novo Excel
if not resumo_total.empty:
    resumo_total.to_excel(saida, index=False)
    print(f"\n🏁 Resumo completo salvo com sucesso em:\n{saida}")
else:
    print("⚠️ Nenhum dado foi processado — verifique se os arquivos têm as planilhas corretas.")
