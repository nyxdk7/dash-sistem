from flask import Flask, render_template, request, redirect, session, url_for
from functools import wraps
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import urllib.parse
from datetime import datetime
from zoneinfo import ZoneInfo
from services.importador_medicao import extrair_medicao
from zoneinfo import ZoneInfo

app = Flask(__name__)
app.secret_key = 'segredo123'

senha = urllib.parse.quote_plus("1A2b3c4d.")

app.config['SQLALCHEMY_DATABASE_URI'] = f'postgresql://joao:{senha}@34.39.230.118:5432/diario_obra'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)


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

    data_registro = db.Column(
        db.DateTime,
        default=lambda: datetime.now(ZoneInfo("America/Rio_Branco")),
        nullable=False
    )

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
    data_importacao = db.Column(db.DateTime, default=lambda: datetime.now(ZoneInfo("America/Rio_Branco")), nullable=False)

    obra_id = db.Column(db.Integer, db.ForeignKey('obra.id'), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)

    processo = db.Column(db.String(255), nullable=True)
    rodovia = db.Column(db.String(255), nullable=True)
    trecho = db.Column(db.String(255), nullable=True)
    subtrecho = db.Column(db.String(255), nullable=True)
    periodo = db.Column(db.String(255), nullable=True)
    medicao = db.Column(db.String(255), nullable=True)

    obra = db.relationship('Obra')
    usuario = db.relationship('Usuario')
    itens = db.relationship('ItemMedicao', backref='medicao_rel', lazy=True, cascade='all, delete-orphan')


class ItemMedicao(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    medicao_id = db.Column(db.Integer, db.ForeignKey('medicao.id'), nullable=False)

    codigo = db.Column(db.String(50), nullable=True)
    item = db.Column(db.Text, nullable=True)
    unidade = db.Column(db.String(50), nullable=True)
    quantidade = db.Column(db.Float, nullable=True)
    preco = db.Column(db.Float, nullable=True)
    financeiro = db.Column(db.Float, nullable=True)


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
                return "Acesso negado"
            return f(*args, **kwargs)
        return wrap
    return decorator


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
        elif user.nivel == 'engenheiro':
            return redirect(url_for('minha_obra'))
        else:
            session.clear()
            return "Nível de usuário inválido"

    return "Login inválido"


@app.route('/dashboard')
@login_required
def dashboard():
    if session.get('nivel') != 'admin':
        return redirect(url_for('minha_obra'))

    obras = Obra.query.all()
    registros = DiarioObra.query.order_by(DiarioObra.data_registro.desc()).limit(10).all()

    return render_template('dashboard.html', admin=True, obras=obras, registros=registros)


@app.route('/minha-obra', methods=['GET', 'POST'])
@login_required
def minha_obra():
    if session.get('nivel') == 'admin':
        return redirect(url_for('dashboard'))

    obra_id = session.get('obra_id')
    if not obra_id:
        return "Você não está vinculado a nenhuma obra"

    obra = db.session.get(Obra, obra_id)
    if not obra:
        return "Obra não encontrada"

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
            return "Preencha a descrição"

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
            return "Preencha todos os campos"

        usuario_existente = Usuario.query.filter_by(username=username).first()
        if usuario_existente:
            return "Usuário já existe"

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
            return "Informe o nome da obra"

        obra = Obra(nome=nome, lote=lote)
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
        return "Você não pode excluir seu próprio usuário"

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
        return "Acesso negado"

    db.session.delete(registro)
    db.session.commit()

    return redirect(url_for('minha_obra'))

@app.route('/editar-registro/<int:id>', methods=['GET', 'POST'])
@login_required
def editar_registro(id):
    registro = DiarioObra.query.get_or_404(id)

    # Permissão:
    # o admin pode editar qualquer registro 
    # engenheiro só pode editar o próprio registro
    if session.get('nivel') != 'admin' and registro.usuario_id != session.get('user_id'):
        return "Acesso negado"

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
            return "Preencha a descrição"

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

import pandas as pd

from services.importador_medicao import extrair_medicao

@app.route('/importacoes', methods=['GET', 'POST'])
@login_required
def importacoes():
    dados = None
    nome_arquivo = None
    cabecalho = None
    total_itens = 0
    obras = Obra.query.all()
    abas = []  # <-- CORREÇÃO AQUI

    if request.method == 'POST':
        arquivo = request.files.get('arquivo')

        if not arquivo or not arquivo.filename:
            flash('Selecione um arquivo .xlsx para importar.', 'warning')
        else:
            nome_arquivo = secure_filename(arquivo.filename)

            if not arquivo_xlsx_valido(nome_arquivo):
                flash('Apenas arquivos .xlsx são suportados.', 'danger')
            else:
                try:
                    cabecalho, itens = extrair_medicao(arquivo)
                    cabecalho = cabecalho or {}
                    dados = itens
                    total_itens = len(itens)

                    # 👇 cria lista de abas baseada nos itens
                    abas = list(set([item.get('aba') for item in itens if item.get('aba')]))

                    if total_itens == 0:
                        flash('Nenhum item válido encontrado na planilha.', 'warning')
                    else:
                        flash(
                            f'Planilha "{nome_arquivo}" lida com sucesso. {total_itens} itens encontrados.',
                            'success'
                        )

                except Exception as e:
                    flash(f'Erro ao processar planilha: {str(e)}', 'danger')

    return render_template(
        'importacoes.html',
        abas=abas,
        nome_arquivo=nome_arquivo,
        obras=obras
    )

@app.route('/salvar-medicao', methods=['POST'])
@login_required
def salvar_medicao():
    arquivo = request.files.get('arquivo')

    if not arquivo or not arquivo.filename:
        return "Nenhum arquivo enviado"

    if not arquivo.filename.endswith('.xlsx'):
        return "Apenas arquivos .xlsx são suportados"

    try:
        cabecalho, itens = extrair_medicao(arquivo)

        obra_id = request.form.get('obra_id')

        # se for admin sem obra vinculada, bloqueia por enquanto
        if not obra_id:
           return "Selecione uma obra"

        nova_medicao = Medicao(
            nome_arquivo=arquivo.filename,
            obra_id=obra_id,
            usuario_id=session.get('user_id'),
            processo=cabecalho.get('processo'),
            rodovia=cabecalho.get('rodovia'),
            trecho=cabecalho.get('trecho'),
            subtrecho=cabecalho.get('subtrecho'),
            periodo=cabecalho.get('periodo'),
            medicao=cabecalho.get('medicao')
        )

        db.session.add(nova_medicao)
        db.session.flush()

        for item in itens:
            novo_item = ItemMedicao(
                medicao_id=nova_medicao.id,
                codigo=str(item.get('codigo')) if item.get('codigo') is not None else None,
                item=str(item.get('item')) if item.get('item') is not None else None,
                unidade=str(item.get('unidade')) if item.get('unidade') is not None else None,
                quantidade=item.get('quantidade') if item.get('quantidade') is not None else 0,
                preco=item.get('preco') if item.get('preco') is not None else 0,
                financeiro=item.get('financeiro') if item.get('financeiro') is not None else 0
            )
            db.session.add(novo_item)

        db.session.commit()
        return redirect(url_for('importacoes'))

    except Exception as e:
        db.session.rollback()
        return f"Erro ao salvar medição: {str(e)}"

@app.route('/medicoes')
@login_required
def medicoes():
    medicoes = Medicao.query.order_by(Medicao.id.desc()).all()

    total_medicoes = len(medicoes)
    total_itens = ItemMedicao.query.count()

    soma_total = db.session.query(db.func.sum(ItemMedicao.financeiro)).scalar() or 0

    return render_template(
        'medicoes.html',
        medicoes=medicoes,
        total_medicoes=total_medicoes,
        total_itens=total_itens,
        soma_total=soma_total
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)