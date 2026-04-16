from flask import Flask, render_template, request, redirect, session, url_for, flash
from functools import wraps
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import urllib.parse
from datetime import datetime
from zoneinfo import ZoneInfo
from flask import render_template, request, flash
from werkzeug.utils import secure_filename
from importador_medicao import extrair_medicao
from flask import redirect

from services.importador_medicao import extrair_medicao

app = Flask(__name__)
app.secret_key = 'segredo123'

senha = urllib.parse.quote_plus("1A2b3c4d.")

app.config['SQLALCHEMY_DATABASE_URI'] = f'postgresql://joao:{senha}@34.39.230.118:5432/diario_obra'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB

db = SQLAlchemy(app)


def agora_br():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


def arquivo_xlsx_valido(nome_arquivo):
    return bool(nome_arquivo) and nome_arquivo.lower().endswith('.xlsx')


def to_float(valor, padrao=0.0):
    if valor is None or valor == '':
        return padrao

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    texto = texto.replace('R$', '').replace(' ', '')

    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')

    try:
        return float(texto)
    except ValueError:
        return padrao
    
def arquivo_xlsx_valido(nome_arquivo):
    return nome_arquivo.lower().endswith(".xlsx")


class Obra(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    lote = db.Column(db.String(50), nullable=True)

    usuarios = db.relationship('Usuario', backref='obra', lazy=True)


class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    senha = db.Column(db.String(255), nullable=False)
    nivel = db.Column(db.String(20), nullable=False, default='engenheiro')
    obra_id = db.Column(db.Integer, db.ForeignKey('obra.id'), nullable=True)


class DiarioObra(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    descricao = db.Column(db.Text, nullable=False)
    data_registro = db.Column(db.DateTime, default=agora_br, nullable=False)

    clima = db.Column(db.String(50), nullable=True)
    efetivo = db.Column(db.Text, nullable=True)
    ocorrencias = db.Column(db.Text, nullable=True)

    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    obra_id = db.Column(db.Integer, db.ForeignKey('obra.id'), nullable=False)

    usuario = db.relationship('Usuario')
    obra = db.relationship('Obra')


class Medicao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome_arquivo = db.Column(db.String(255), nullable=False)
    data_importacao = db.Column(db.DateTime, default=agora_br, nullable=False)

    obra_id = db.Column(db.Integer, db.ForeignKey('obra.id'), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

    processo = db.Column(db.String(255), nullable=True)
    contrato = db.Column(db.String(255), nullable=True)
    contratada = db.Column(db.String(255), nullable=True)
    rodovia = db.Column(db.String(255), nullable=True)
    trecho = db.Column(db.String(255), nullable=True)
    subtrecho = db.Column(db.String(255), nullable=True)
    segmento = db.Column(db.String(255), nullable=True)
    periodo = db.Column(db.String(255), nullable=True)
    periodo_acumulado = db.Column(db.String(255), nullable=True)
    medicao = db.Column(db.String(255), nullable=True)

    obra = db.relationship('Obra')
    usuario = db.relationship('Usuario')
    itens = db.relationship(
        'ItemMedicao',
        backref='medicao_rel',
        lazy=True,
        cascade='all, delete-orphan'
    )


class ItemMedicao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    medicao_id = db.Column(db.Integer, db.ForeignKey('medicao.id'), nullable=False)

    aba = db.Column(db.String(255), nullable=True)
    tipo_aba = db.Column(db.String(50), nullable=True)

    codigo = db.Column(db.String(50), nullable=True)
    item = db.Column(db.Text, nullable=True)
    unidade = db.Column(db.String(50), nullable=True)
    quantidade = db.Column(db.Float, nullable=True)

    preco = db.Column(db.Float, nullable=True)
    financeiro = db.Column(db.Float, nullable=True)

    marca = db.Column(db.String(255), nullable=True)
    observacao = db.Column(db.Text, nullable=True)


with app.app_context():
    db.create_all()


def login_required(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('home'))
        return f(*args, **kwargs)
    return wrap


def nivel_required(nivel_permitido):
    def decorator(f):
        @wraps(f)
        def wrap(*args, **kwargs):
            if 'nivel' not in session or session['nivel'] != nivel_permitido:
                return "Acesso negado", 403
            return f(*args, **kwargs)
        return wrap
    return decorator


def ler_planilha_bruta(arquivo):
    import openpyxl

    wb = openpyxl.load_workbook(arquivo, data_only=True)
    abas = []

    for ws in wb.worksheets:
        dados = []

        for row in ws.iter_rows(values_only=True):
            linha = []
            for cell in row:
                linha.append(str(cell) if cell is not None else '')
            dados.append(linha)

        abas.append({
            "nome": ws.title,
            "linhas": dados
        })

    return abas


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()

    user = Usuario.query.filter_by(username=username).first()

    if user and check_password_hash(user.senha, password):
        session['user'] = user.username
        session['user_id'] = user.id
        session['nivel'] = user.nivel
        session['obra_id'] = user.obra_id

        if user.nivel == 'admin':
            return redirect(url_for('dashboard'))
        if user.nivel == 'engenheiro':
            return redirect(url_for('minha_obra'))

        session.clear()
        return "Nível de usuário inválido", 400

    return "Login inválido", 401


@app.route('/dashboard')
@login_required
def dashboard():
    if session.get('nivel') != 'admin':
        return redirect(url_for('minha_obra'))

    obras = Obra.query.all()
    registros = DiarioObra.query.order_by(DiarioObra.data_registro.desc()).limit(10).all()

    total_medicoes = Medicao.query.count()
    total_itens_medicao = ItemMedicao.query.count()
    soma_total = db.session.query(db.func.sum(ItemMedicao.financeiro)).scalar() or 0

    return render_template(
        'dashboard.html',
        admin=True,
        obras=obras,
        registros=registros,
        total_medicoes=total_medicoes,
        total_itens_medicao=total_itens_medicao,
        soma_total=soma_total
    )


@app.route('/minha-obra', methods=['GET', 'POST'])
@login_required
def minha_obra():
    if session.get('nivel') == 'admin':
        return redirect(url_for('dashboard'))

    obra_id = session.get('obra_id')
    if not obra_id:
        return "Você não está vinculado a nenhuma obra", 400

    obra = db.session.get(Obra, obra_id)
    if not obra:
        return "Obra não encontrada", 404

    if request.method == 'POST':
        descricao = request.form.get('descricao', '').strip()
        clima = request.form.get('clima', '').strip()
        quantidades = request.form.getlist('quantidade[]')
        funcoes = request.form.getlist('funcao[]')
        ocorrencias = request.form.get('ocorrencias', '').strip()

        efetivo_lista = []

        for quantidade, funcao in zip(quantidades, funcoes):
            quantidade = quantidade.strip()
            funcao = funcao.strip()

            if quantidade and funcao:
                efetivo_lista.append(f"{quantidade} {funcao}")

        efetivo = ", ".join(efetivo_lista) if efetivo_lista else None

        if not descricao:
            return "Preencha a descrição", 400

        novo_registro = DiarioObra(
            descricao=descricao,
            clima=clima if clima else None,
            efetivo=efetivo,
            ocorrencias=ocorrencias if ocorrencias else None,
            usuario_id=session.get('user_id'),
            obra_id=obra_id
        )

        db.session.add(novo_registro)
        db.session.commit()

        return redirect(url_for('minha_obra'))

    registros = DiarioObra.query.filter_by(obra_id=obra_id).order_by(DiarioObra.data_registro.desc()).all()

    return render_template('minha_obra.html', obra=obra, registros=registros)


@app.route('/contratos')
@login_required
def contratos():
    return render_template('contratos.html')


@app.route('/logout')
@login_required
def logout():
    session.clear()
    return redirect(url_for('home'))


@app.route('/cadastro', methods=['GET', 'POST'])
def cadastro():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        if not username or not password:
            return "Preencha todos os campos", 400

        usuario_existente = Usuario.query.filter_by(username=username).first()
        if usuario_existente:
            return "Usuário já existe", 400

        novo_usuario = Usuario(
            username=username,
            senha=generate_password_hash(password),
            nivel='engenheiro'
        )

        db.session.add(novo_usuario)
        db.session.commit()

        return redirect(url_for('home'))

    return render_template('cadastro.html')


@app.route('/admin')
@login_required
@nivel_required('admin')
def admin():
    usuarios = Usuario.query.all()
    obras = Obra.query.all()
    return render_template('admin.html', usuarios=usuarios, obras=obras)


@app.route('/obra/nova', methods=['GET', 'POST'])
@login_required
@nivel_required('admin')
def nova_obra():
    if request.method == 'POST':
        nome = request.form.get('nome', '').strip()
        lote = request.form.get('lote', '').strip()

        if not nome:
            return "Informe o nome da obra", 400

        obra = Obra(nome=nome, lote=lote if lote else None)
        db.session.add(obra)
        db.session.commit()

        return redirect(url_for('admin'))

    return render_template('nova_obra.html')


@app.route('/admin/vincular-obra/<int:user_id>', methods=['POST'])
@login_required
@nivel_required('admin')
def vincular_obra(user_id):
    usuario = Usuario.query.get_or_404(user_id)
    obra_id = request.form.get('obra_id')

    if obra_id == '':
        usuario.obra_id = None
    else:
        usuario.obra_id = int(obra_id)

    db.session.commit()
    return redirect(url_for('admin'))


@app.route('/admin/alterar-nivel/<int:user_id>', methods=['POST'])
@login_required
@nivel_required('admin')
def alterar_nivel(user_id):
    usuario = Usuario.query.get_or_404(user_id)
    novo_nivel = request.form.get('nivel')

    if novo_nivel not in ['admin', 'engenheiro']:
        return "Nível inválido", 400

    usuario.nivel = novo_nivel
    db.session.commit()

    return redirect(url_for('admin'))


@app.route('/admin/excluir-usuario/<int:user_id>', methods=['POST'])
@login_required
@nivel_required('admin')
def excluir_usuario(user_id):
    usuario = Usuario.query.get_or_404(user_id)

    if usuario.username == session.get('user'):
        return "Você não pode excluir seu próprio usuário", 400

    db.session.delete(usuario)
    db.session.commit()

    return redirect(url_for('admin'))


@app.route('/admin/excluir-obra/<int:obra_id>', methods=['POST'])
@login_required
@nivel_required('admin')
def excluir_obra(obra_id):
    obra = Obra.query.get_or_404(obra_id)

    for usuario in obra.usuarios:
        usuario.obra_id = None

    db.session.delete(obra)
    db.session.commit()

    return redirect(url_for('admin'))


@app.route('/excluir-registro/<int:id>', methods=['POST'])
@login_required
def excluir_registro(id):
    registro = DiarioObra.query.get_or_404(id)

    if session.get('nivel') != 'admin' and registro.usuario_id != session.get('user_id'):
        return "Acesso negado", 403

    db.session.delete(registro)
    db.session.commit()

    if session.get('nivel') == 'admin':
        return redirect(url_for('dashboard'))
    return redirect(url_for('minha_obra'))


@app.route('/editar-registro/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_registro(id):
    registro = DiarioObra.query.get_or_404(id)

    if session.get('nivel') != 'admin' and registro.usuario_id != session.get('user_id'):
        return "Acesso negado", 403

    if request.method == 'POST':
        descricao = request.form.get('descricao', '').strip()
        clima = request.form.get('clima', '').strip()
        quantidades = request.form.getlist('quantidade[]')
        funcoes = request.form.getlist('funcao[]')
        ocorrencias = request.form.get('ocorrencias', '').strip()

        efetivo_lista = []

        for quantidade, funcao in zip(quantidades, funcoes):
            quantidade = quantidade.strip()
            funcao = funcao.strip()

            if quantidade and funcao:
                efetivo_lista.append(f"{quantidade} {funcao}")

        efetivo = ", ".join(efetivo_lista) if efetivo_lista else None

        if not descricao:
            return "Preencha a descrição", 400

        registro.descricao = descricao
        registro.clima = clima if clima else None
        registro.efetivo = efetivo
        registro.ocorrencias = ocorrencias if ocorrencias else None

        db.session.commit()

        if session.get('nivel') == 'admin':
            return redirect(url_for('dashboard'))
        return redirect(url_for('minha_obra'))

    efetivo_formatado = []

    if registro.efetivo:
        itens = registro.efetivo.split(',')

        for item in itens:
            item = item.strip()
            partes = item.split(' ', 1)

            if len(partes) == 2:
                quantidade, funcao = partes
                efetivo_formatado.append({
                    'quantidade': quantidade,
                    'funcao': funcao
                })

    return render_template(
        'editar_registro.html',
        registro=registro,
        efetivo_formatado=efetivo_formatado
    )

@app.route('/importacoes')
def importacoes_redirect():
    return redirect('/medicao-consolidada')

@app.route('/salvar-medicao', methods=['POST'])
@login_required
def salvar_medicao():
    arquivo = request.files.get('arquivo')

    if not arquivo or not arquivo.filename:
        flash('Nenhum arquivo enviado.', 'danger')
        return redirect(url_for('importacoes'))

    nome_arquivo = secure_filename(arquivo.filename)

    if not arquivo_xlsx_valido(nome_arquivo):
        flash('Apenas arquivos .xlsx são suportados.', 'danger')
        return redirect(url_for('importacoes'))

    try:
        cabecalho, itens = extrair_medicao(arquivo)
        cabecalho = cabecalho or {}

        if not itens:
            flash('Nenhum item válido foi encontrado na planilha.', 'warning')
            return redirect(url_for('importacoes'))

        obra_id = request.form.get('obra_id')

        if not obra_id and session.get('nivel') == 'engenheiro':
            obra_id = session.get('obra_id')

        if not obra_id:
            flash('Selecione uma obra.', 'warning')
            return redirect(url_for('importacoes'))

        obra_id = int(obra_id)

        nova_medicao = Medicao(
            nome_arquivo=nome_arquivo,
            obra_id=obra_id,
            usuario_id=session.get('user_id'),
            processo=cabecalho.get('processo'),
            contrato=cabecalho.get('contrato'),
            contratada=cabecalho.get('contratada'),
            rodovia=cabecalho.get('rodovia'),
            trecho=cabecalho.get('trecho'),
            subtrecho=cabecalho.get('sub_trecho'),
            segmento=cabecalho.get('segmento'),
            periodo=cabecalho.get('periodo'),
            periodo_acumulado=cabecalho.get('periodo_acumulado'),
            medicao=cabecalho.get('medicao')
        )

        db.session.add(nova_medicao)
        db.session.flush()

        novos_itens = []

        for item in itens:
            novo_item = ItemMedicao(
                medicao_id=nova_medicao.id,
                aba=str(item.get('aba')).strip() if item.get('aba') is not None else None,
                tipo_aba=str(item.get('tipo_aba')).strip() if item.get('tipo_aba') is not None else None,
                codigo=str(item.get('codigo')).strip() if item.get('codigo') is not None else None,
                item=str(item.get('descricao')).strip() if item.get('descricao') is not None else None,
                unidade=str(item.get('unidade')).strip() if item.get('unidade') is not None else None,
                quantidade=to_float(item.get('quantidade'), 0),
                preco=to_float(item.get('preco_unitario'), 0),
                financeiro=to_float(item.get('preco_total'), 0),
                marca=str(item.get('marca')).strip() if item.get('marca') is not None else None,
                observacao=str(item.get('observacao')).strip() if item.get('observacao') is not None else None
            )
            novos_itens.append(novo_item)

        db.session.bulk_save_objects(novos_itens)
        db.session.commit()

        flash(f'Medição salva com sucesso. {len(novos_itens)} itens gravados.', 'success')
        return redirect(url_for('medicoes'))

    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao salvar medição: {str(e)}', 'danger')
        return redirect(url_for('importacoes'))


@app.route('/medicoes')
@login_required
def medicoes():
    query = Medicao.query

    if session.get('nivel') != 'admin':
        obra_id = session.get('obra_id')
        query = query.filter_by(obra_id=obra_id)

    medicoes = query.order_by(Medicao.id.desc()).all()
    ids_medicoes = [m.id for m in medicoes]
    total_medicoes = len(medicoes)

    if ids_medicoes:
        total_itens = ItemMedicao.query.filter(ItemMedicao.medicao_id.in_(ids_medicoes)).count()
        soma_total = (
            db.session.query(db.func.sum(ItemMedicao.financeiro))
            .filter(ItemMedicao.medicao_id.in_(ids_medicoes))
            .scalar() or 0
        )
    else:
        total_itens = 0
        soma_total = 0

    return render_template(
        'medicoes.html',
        medicoes=medicoes,
        total_medicoes=total_medicoes,
        total_itens=total_itens,
        soma_total=soma_total
    )


@app.route('/medicao/<int:id>')
@login_required
def ver_medicao(id):
    medicao = Medicao.query.get_or_404(id)

    if session.get('nivel') != 'admin' and medicao.obra_id != session.get('obra_id'):
        return "Acesso negado", 403

    itens = ItemMedicao.query.filter_by(medicao_id=id).all()

    return render_template(
        'ver_medicao.html',
        medicao=medicao,
        itens=itens
    )

from flask import redirect, request, render_template, flash
from werkzeug.utils import secure_filename

@app.route('/medicao-consolidada', methods=['GET', 'POST'])
@login_required
def medicao_consolidada():
    cabecalho = None
    itens = []
    resumo = None
    grupos_dashboard = []
    top_itens_atual = []
    nome_arquivo = None

    if request.method == 'POST':

        # 🔴 garante que o botão foi clicado
        if 'enviar' not in request.form:
            flash('Erro no envio do formulário.', 'danger')
            return redirect(request.url)

        # 🔴 garante que o arquivo chegou
        if 'arquivo' not in request.files:
            flash('Arquivo não chegou no servidor.', 'danger')
            return redirect(request.url)

        arquivo = request.files['arquivo']

        # 🔴 garante que não veio vazio
        if arquivo.filename == '':
            flash('Nenhum arquivo selecionado.', 'warning')
            return redirect(request.url)

        nome_arquivo = secure_filename(arquivo.filename)

        if not arquivo_xlsx_valido(nome_arquivo):
            flash('Apenas arquivos .xlsx são suportados.', 'danger')
        else:
            try:
                cabecalho, itens = extrair_medicao_consolidada(arquivo)

                resumo = {
                    'total_itens': len(itens),
                    'total_financeiro_acumulado_anterior': sum(float(item.get('financeiro_acumulado_anterior', 0) or 0) for item in itens),
                    'total_financeiro_liquido_atual': sum(float(item.get('financeiro_liquido_atual', 0) or 0) for item in itens),
                    'total_financeiro_acumulado_atual': sum(float(item.get('financeiro_acumulado_atual', 0) or 0) for item in itens),
                    'total_saldo_financeiro': sum(float(item.get('saldo_financeiro', 0) or 0) for item in itens),
                    'pi_mais_reajuste': 0
                }

                grupos = {}
                for item in itens:
                    grupo = item.get('grupo') or 'SEM_GRUPO'
                    if grupo not in grupos:
                        grupos[grupo] = {
                            'grupo': grupo,
                            'quantidade_itens': 0,
                            'contrato_financeiro': 0,
                            'financeiro_liquido_atual': 0,
                            'financeiro_acumulado_atual': 0,
                            'saldo_financeiro': 0
                        }

                    grupos[grupo]['quantidade_itens'] += 1
                    grupos[grupo]['contrato_financeiro'] += float(item.get('contrato_financeiro', 0) or 0)
                    grupos[grupo]['financeiro_liquido_atual'] += float(item.get('financeiro_liquido_atual', 0) or 0)
                    grupos[grupo]['financeiro_acumulado_atual'] += float(item.get('financeiro_acumulado_atual', 0) or 0)
                    grupos[grupo]['saldo_financeiro'] += float(item.get('saldo_financeiro', 0) or 0)

                grupos_dashboard = sorted(
                    grupos.values(),
                    key=lambda g: g['financeiro_liquido_atual'],
                    reverse=True
                )

                top_itens_atual = sorted(
                    itens,
                    key=lambda i: float(i.get('financeiro_liquido_atual', 0) or 0),
                    reverse=True
                )[:10]

                flash('Aba MEDIÇÃO CONSOLIDADA carregada com sucesso.', 'success')

            except Exception as e:
                flash(f'Erro ao ler a aba MEDIÇÃO CONSOLIDADA: {e}', 'danger')

    return render_template(
        'medicao_consolidada.html',
        cabecalho=cabecalho,
        itens=itens,
        resumo=resumo,
        grupos_dashboard=grupos_dashboard,
        top_itens_atual=top_itens_atual,
        nome_arquivo=nome_arquivo
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)