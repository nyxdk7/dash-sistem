import openpyxl


def limpar_texto(valor):
    if valor is None:
        return ''
    return str(valor).strip()


def converter_numero(valor):
    if valor is None or valor == '':
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    # remove espaços
    texto = texto.replace(' ', '')

    # caso venha no padrão brasileiro: 1.234,56
    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')

    try:
        return float(texto)
    except Exception:
        return 0.0


def normalizar_texto(valor):
    return limpar_texto(valor).lower()


def extrair_cabecalho_padrao(ws):
    cabecalho = {
        'obra': '',
        'empresa': '',
        'contrato': '',
        'boletim': '',
        'medicao': ''
    }

    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 15), values_only=True):
        for cell in row:
            texto = normalizar_texto(cell)

            if not texto:
                continue

            if 'obra' in texto and not cabecalho['obra']:
                cabecalho['obra'] = limpar_texto(cell)
            elif 'empresa' in texto and not cabecalho['empresa']:
                cabecalho['empresa'] = limpar_texto(cell)
            elif 'contrato' in texto and not cabecalho['contrato']:
                cabecalho['contrato'] = limpar_texto(cell)
            elif 'boletim' in texto and not cabecalho['boletim']:
                cabecalho['boletim'] = limpar_texto(cell)
            elif 'medi' in texto and not cabecalho['medicao']:
                cabecalho['medicao'] = limpar_texto(cell)

    return cabecalho


def detectar_tipo_aba(ws):
    titulo = ws.title.strip().lower()

    if 'ilum' in titulo:
        return 'iluminacao'
    if 'eros' in titulo:
        return 'erosao'
    if 'cbuq' in titulo:
        return 'cbuq'
    if 'contrato' in titulo:
        return 'contrato'

    # fallback pelo conteúdo
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True):
        linha = ' '.join([normalizar_texto(c) for c in row if c is not None])

        if 'ilum' in linha:
            return 'iluminacao'
        if 'eros' in linha:
            return 'erosao'
        if 'cbuq' in linha:
            return 'cbuq'
        if 'contrato' in linha:
            return 'contrato'

    return 'desconhecido'


def localizar_linha_cabecalho_generico(ws):
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        valores = [normalizar_texto(v) for v in row if v is not None]
        linha = ' '.join(valores)

        tem_codigo = 'codigo' in linha or 'código' in linha
        tem_descricao = 'descr' in linha
        tem_unidade = 'unid' in linha
        tem_quantidade = 'quant' in linha

        if tem_codigo and tem_descricao and (tem_unidade or tem_quantidade):
            return i

    return None


def mapear_colunas_por_cabecalho(ws, linha_cabecalho):
    headers = []
    for cell in ws[linha_cabecalho]:
        headers.append(normalizar_texto(cell.value))

    def encontrar_coluna(*termos, obrigatorios=True):
        for idx, h in enumerate(headers):
            if not h:
                continue

            if obrigatorios:
                if all(t in h for t in termos):
                    return idx
            else:
                if any(t in h for t in termos):
                    return idx
        return None

    mapa = {
        'aba': encontrar_coluna('aba'),
        'tipo_aba': encontrar_coluna('tipo'),
        'codigo': encontrar_coluna('codigo') if encontrar_coluna('codigo') is not None else encontrar_coluna('código'),
        'descricao': encontrar_coluna('descr'),
        'unidade': encontrar_coluna('unid'),
        'quantidade': encontrar_coluna('quant'),
        'observacao': encontrar_coluna('observ'),
        'marca': encontrar_coluna('marca'),
        'preco_unitario': None,
        'preco_total': None,
    }

    # tenta achar preço unitário
    for idx, h in enumerate(headers):
        if not h:
            continue
        if ('preco' in h or 'preço' in h or 'valor' in h) and ('unit' in h):
            mapa['preco_unitario'] = idx
            break

    # tenta achar preço total
    for idx, h in enumerate(headers):
        if not h:
            continue
        if ('preco' in h or 'preço' in h or 'valor' in h) and ('total' in h):
            mapa['preco_total'] = idx
            break

    return mapa


def extrair_aba_medicao_consolidada(ws):
    cabecalho = extrair_cabecalho_padrao(ws)
    itens = []

    linha_cabecalho = localizar_linha_cabecalho_generico(ws)
    if not linha_cabecalho:
        return cabecalho, itens

    mapa = mapear_colunas_por_cabecalho(ws, linha_cabecalho)

    for row in ws.iter_rows(min_row=linha_cabecalho + 1, max_row=ws.max_row, values_only=True):
        if not row:
            continue

        if all(v is None or limpar_texto(v) == '' for v in row):
            continue

        codigo = row[mapa['codigo']] if mapa['codigo'] is not None and mapa['codigo'] < len(row) else None
        descricao = row[mapa['descricao']] if mapa['descricao'] is not None and mapa['descricao'] < len(row) else None

        if limpar_texto(codigo) == '' and limpar_texto(descricao) == '':
            continue

        quantidade = row[mapa['quantidade']] if mapa['quantidade'] is not None and mapa['quantidade'] < len(row) else 0
        preco_unitario = row[mapa['preco_unitario']] if mapa['preco_unitario'] is not None and mapa['preco_unitario'] < len(row) else 0
        preco_total = row[mapa['preco_total']] if mapa['preco_total'] is not None and mapa['preco_total'] < len(row) else 0

        # se não tiver preço total na planilha, calcula
        quantidade_num = converter_numero(quantidade)
        preco_unitario_num = converter_numero(preco_unitario)
        preco_total_num = converter_numero(preco_total)

        if preco_total_num == 0.0 and quantidade_num and preco_unitario_num:
            preco_total_num = quantidade_num * preco_unitario_num

        item = {
            'aba': limpar_texto(row[mapa['aba']]) if mapa['aba'] is not None and mapa['aba'] < len(row) else ws.title,
            'tipo_aba': limpar_texto(row[mapa['tipo_aba']]) if mapa['tipo_aba'] is not None and mapa['tipo_aba'] < len(row) else detectar_tipo_aba(ws),
            'codigo': limpar_texto(codigo),
            'descricao': limpar_texto(descricao),
            'unidade': limpar_texto(row[mapa['unidade']]) if mapa['unidade'] is not None and mapa['unidade'] < len(row) else '',
            'quantidade': quantidade_num,
            'observacao': limpar_texto(row[mapa['observacao']]) if mapa['observacao'] is not None and mapa['observacao'] < len(row) else '',
            'marca': limpar_texto(row[mapa['marca']]) if mapa['marca'] is not None and mapa['marca'] < len(row) else '',
            'preco_unitario': preco_unitario_num,
            'preco_total': preco_total_num,
        }

        itens.append(item)

    return cabecalho, itens


def extrair_aba_contrato(ws):
    # agora contrato usa a mesma lógica genérica baseada em cabeçalho
    return extrair_aba_medicao_consolidada(ws)


def extrair_aba_iluminacao(ws):
    return extrair_aba_medicao_consolidada(ws)


def extrair_aba_erosao(ws):
    return extrair_aba_medicao_consolidada(ws)


def extrair_aba_cbuq(ws):
    return extrair_aba_medicao_consolidada(ws)


def extrair_medicao(arquivo):
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    cabecalho_geral = None

    nomes_prioritarios = [
        'medição consolidada',
        'medicao consolidada',
        'medição resumo',
        'medicao resumo',
        'resumo'
    ]

    todos_itens = []

    # 1. tenta achar primeiro a aba principal
    for ws in wb.worksheets:
        titulo = ws.title.strip().lower()

        if any(nome in titulo for nome in nomes_prioritarios):
            cabecalho, itens = extrair_aba_medicao_consolidada(ws)

            if cabecalho:
                cabecalho_geral = cabecalho

            if itens:
                return cabecalho_geral or {}, itens

    # 2. se não achar, cai no comportamento antigo
    for ws in wb.worksheets:
        tipo = detectar_tipo_aba(ws)

        if tipo == 'contrato':
            cabecalho, itens = extrair_aba_contrato(ws)
        elif tipo == 'iluminacao':
            cabecalho, itens = extrair_aba_iluminacao(ws)
        elif tipo == 'erosao':
            cabecalho, itens = extrair_aba_erosao(ws)
        elif tipo == 'cbuq':
            cabecalho, itens = extrair_aba_cbuq(ws)
        else:
            continue

        if not cabecalho_geral and cabecalho:
            cabecalho_geral = cabecalho

        todos_itens.extend(itens)

    return cabecalho_geral or {}, todos_itens