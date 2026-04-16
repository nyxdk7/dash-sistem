from openpyxl import load_workbook


def extrair_medicao(arquivo):
    wb = load_workbook(arquivo, data_only=True)

    # tenta pegar a aba principal
    aba = wb.active

    linha_inicio = None

    # 🔍 encontrar linha onde começa a tabela
    for i, row in enumerate(aba.iter_rows(values_only=True), start=1):
        valores = [str(c).strip().upper() if c else "" for c in row]

        if "CÓDIGO" in valores and "ITEM" in valores:
            linha_inicio = i
            break

    if linha_inicio is None:
        raise ValueError("Não foi possível encontrar a tabela de medição.")

    # 📌 pegar cabeçalho
    header = [str(c).strip() if c else "" for c in aba[linha_inicio]]

    itens = []

    # 📊 ler linhas abaixo do cabeçalho
    for row in aba.iter_rows(min_row=linha_inicio + 1, values_only=True):
        linha = list(row)

        # parar quando chegar em total
        if any("TOTAL" in str(c).upper() for c in linha if c):
            break

        codigo = linha[0]

        # ignora linhas vazias
        if not codigo:
            continue

        item = {
            "codigo": codigo,
            "item": linha[1],
            "unidade": linha[2],
            "quantidade": linha[3],
            "preco": linha[4],
            "financeiro": linha[5]
        }

        itens.append(item)

    return itens