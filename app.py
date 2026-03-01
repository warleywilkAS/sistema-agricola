import pandas as pd
from io import BytesIO
from flask import send_file
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from models import db, FormularioSoja, Pulverizacao
import json

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SECRET_KEY'] = 'chave-secreta-para-formulario-agricola-2026'
db.init_app(app)

# Listas de alvos
INSETOS_ALVO = [
    "Lagarta da soja (Anticarsia gemmatalis)",
    "Lagarta das vagens (Spodoptera spp.)",
    "Lagarta falsa medideira (Chrysodeixis includens)",
    "Lagartas do grupo Heliothinae",
    "Percevejo barriga verde (Dichelops spp.)",
    "Percevejo marrom (Euschistus heros)",
    "Percevejo verde (Nezara viridula)",
    "Percevejo verde pequeno (Piezodorus guildinii)",
    "Broca dos ponteiros (Crocidosema aporema)",
    "Mosca Branca",
    "Outros insetos praga",
    "Tamanduá da soja (Sternechus subsignatus)",
    "Tripes",
    "Vaquinhas (Diabrotica/ Cerotoma/ Colapsis)"
]

DOENCAS_ALVO = [
    "Antracnose (Colletotrichum truncatum)",
    "Cancro da haste (Diaporthe spp.)",
    "Ferrugem asiática (Phakopsora pachyrhizi)",
    "Mancha alvo (Corynespora cassicola)",
    "Mancha de cercospora (Cercospora kikuchii)",
    "Mancha olho-de-rã (Cercospora sojina)",
    "Mancha parda (Septoria glycines)",
    "Mela ou requeima (Rhizoctonia solani)",
    "Mofo branco (Sclerotinia sclerotiorum)",
    "Mildio (Peronospora manshurica)",
    "Oídio (Microsphaera diffusa)",
    "Outras Doenças Fungicas"
]

PLANTAS_DANINHAS = [
    "Buva (Conyza spp.)",
    "Capim-amargoso (Digitaria insularis)",
    "Caruru (Amaranthus spp.)",
    "Capim-pé-de-galinha (Eleusine indica)",
    "Leiteiro (Euphorbia heterophylla)",
    "Picão-preto (Bidens pilosa)",
    "Trapoeraba (Commelina spp.)",
    "Outras Plantas Daninhas"
]

ACAROS = [
    "Ácaro-rajado (Tetranychus urticae)",
    "Ácaro-verde (Mononychellus planki)",
    "Ácaro-branco (Polyphagotarsonemus latus)",
    "Ácaros-vermelhos (Tetranychus spp.)",
    "Outros ácaros"
]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/form', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        try:
            # Criar novo formulário
            formulario = FormularioSoja()
            
            # Identificação
            formulario.numero_produtor = request.form.get('numero_produtor')
            formulario.regiao = request.form.get('regiao')
            formulario.municipio = request.form.get('municipio')
            formulario.meso_idr = request.form.get('meso_idr')
            formulario.area_soja = float(request.form.get('area_soja') or 0)
            formulario.produtividade_media = float(request.form.get('produtividade_media') or 0)
            formulario.cultivar = request.form.get('cultivar')
            formulario.bt = request.form.get('bt')
            formulario.data_plantio = request.form.get('data_plantio')
            formulario.data_emergencia = request.form.get('data_emergencia')
            formulario.houve_adversidade = request.form.get('houve_adversidade')
            formulario.qual_adversidade = request.form.get('qual_adversidade')
            formulario.nome_coletor = request.form.get('nome_coletor')
            formulario.unidade_municipal = request.form.get('unidade_municipal')
            
            # Conhecimento MIP e MID
            formulario.conhecimento_mid = request.form.get('conhecimento_mid')
            formulario.utiliza_mid = request.form.get('utiliza_mid')
            formulario.conhecimento_mip = request.form.get('conhecimento_mip')
            formulario.utiliza_mip = request.form.get('utiliza_mip')
            
            # Controle Plantas Invasoras
            formulario.herbicida_dessecacao_alvo = request.form.get('herbicida_dessecacao_alvo')
            formulario.herbicida_dessecacao_aplicacoes = int(request.form.get('herbicida_dessecacao_aplicacoes') or 0)
            formulario.herbicida_pre_alvo = request.form.get('herbicida_pre_alvo')
            formulario.herbicida_pre_aplicacoes = int(request.form.get('herbicida_pre_aplicacoes') or 0)
            formulario.herbicida_pos_alvo = request.form.get('herbicida_pos_alvo')
            formulario.herbicida_pos_aplicacoes = int(request.form.get('herbicida_pos_aplicacoes') or 0)
            
            # Outras informações
            formulario.tratamento_sementes = request.form.get('tratamento_sementes')
            formulario.sal_mistura = request.form.get('sal_mistura')
            formulario.controle_biologico = request.form.get('controle_biologico')
            
            # FBN
            formulario.inoculacao_sementes = request.form.get('inoculacao_sementes')
            formulario.forma_inoculacao = request.form.get('forma_inoculacao')
            formulario.coinoculacao = request.form.get('coinoculacao')
            formulario.co_mo = request.form.get('co_mo')
            formulario.co_mo_aplicacao = request.form.get('co_mo_aplicacao')
            
            db.session.add(formulario)
            db.session.flush()  # Para obter o ID
            
            # Salvar pulverizações
            # Pré-plantio
            if request.form.get('data_pre_plantio'):
                pulv = Pulverizacao(
                    formulario_id=formulario.id,
                    tipo='pre_plantio',
                    data=request.form.get('data_pre_plantio'),
                    classe_produto='Inseticida',
                    alvo=request.form.get('alvo_pre_plantio')
                )
                db.session.add(pulv)
            
            # Pulverizações pós-emergência (até 7)
            for i in range(1, 8):
                data = request.form.get(f'data_pos_{i}')
                if data:
                    classe = request.form.get(f'classe_pos_{i}')
                    alvo = request.form.get(f'alvo_pos_{i}')
                    if classe and alvo:
                        pulv = Pulverizacao(
                            formulario_id=formulario.id,
                            tipo=f'pos_{i}',
                            data=data,
                            classe_produto=classe,
                            alvo=alvo
                        )
                        db.session.add(pulv)
            
            db.session.commit()
            flash('Formulário salvo com sucesso!', 'success')
            return redirect(url_for('view_records'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao salvar: {str(e)}', 'danger')
            return redirect(url_for('form'))
    
    return render_template('form.html', 
                      insetos=INSETOS_ALVO, 
                      doencas=DOENCAS_ALVO,
                      plantas=PLANTAS_DANINHAS,
                      acaros=ACAROS)

@app.route('/records')
def view_records():
    registros = FormularioSoja.query.order_by(FormularioSoja.data_criacao.desc()).all()
    return render_template('view_records.html', registros=registros)

@app.route('/export/excel')
def export_excel():
    # Buscar todos os registros
    registros = FormularioSoja.query.order_by(FormularioSoja.data_criacao.desc()).all()
    
    # Preparar dados para o Excel
    dados = []
    for r in registros:
        # Contar pulverizações
        pulverizacoes_count = len(r.pulverizacoes)
        alvos = ', '.join([f"{p.classe_produto}: {p.alvo}" for p in r.pulverizacoes[:3]])
        
        dados.append({
            'ID': r.id,
            'Data Criação': r.data_criacao.strftime('%d/%m/%Y %H:%M'),
            'Produtor': r.numero_produtor,
            'Município': r.municipio,
            'Área (ha)': r.area_soja,
            'Produtividade': r.produtividade_media,
            'Cultivar': r.cultivar,
            'BT': r.bt,
            'Data Plantio': r.data_plantio,
            'Adversidade': r.qual_adversidade,
            'N° Pulverizações': pulverizacoes_count,
            'Principais Alvos': alvos,
            'Coletor': r.nome_coletor
        })
    
    # Criar DataFrame
    df = pd.DataFrame(dados)
    
    # Criar arquivo Excel em memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Registros Soja', index=False)
    
    output.seek(0)
    
    return send_file(
        output,
        download_name='registros_soja.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/record/<int:id>')
def view_record(id):
    registro = FormularioSoja.query.get_or_404(id)
    return render_template('view_record.html', registro=registro, insetos=INSETOS_ALVO, doencas=DOENCAS_ALVO)

@app.route('/delete/<int:id>', methods=['POST'])
def delete_record(id):
    registro = FormularioSoja.query.get_or_404(id)
    db.session.delete(registro)
    db.session.commit()
    flash('Registro excluído com sucesso!', 'success')
    return redirect(url_for('view_records'))

@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_record(id):
    registro = FormularioSoja.query.get_or_404(id)
    
    if request.method == 'POST':
        try:
            # Atualizar campos (mesma lógica do POST do form)
            registro.numero_produtor = request.form.get('numero_produtor')
            registro.regiao = request.form.get('regiao')
            registro.municipio = request.form.get('municipio')
            registro.meso_idr = request.form.get('meso_idr')
            registro.area_soja = float(request.form.get('area_soja') or 0)
            registro.produtividade_media = float(request.form.get('produtividade_media') or 0)
            registro.cultivar = request.form.get('cultivar')
            registro.bt = request.form.get('bt')
            registro.data_plantio = request.form.get('data_plantio')
            registro.data_emergencia = request.form.get('data_emergencia')
            registro.houve_adversidade = request.form.get('houve_adversidade')
            registro.qual_adversidade = request.form.get('qual_adversidade')
            registro.nome_coletor = request.form.get('nome_coletor')
            registro.unidade_municipal = request.form.get('unidade_municipal')
            
            registro.conhecimento_mid = request.form.get('conhecimento_mid')
            registro.utiliza_mid = request.form.get('utiliza_mid')
            registro.conhecimento_mip = request.form.get('conhecimento_mip')
            registro.utiliza_mip = request.form.get('utiliza_mip')
            
            registro.herbicida_dessecacao_alvo = request.form.get('herbicida_dessecacao_alvo')
            registro.herbicida_dessecacao_aplicacoes = int(request.form.get('herbicida_dessecacao_aplicacoes') or 0)
            registro.herbicida_pre_alvo = request.form.get('herbicida_pre_alvo')
            registro.herbicida_pre_aplicacoes = int(request.form.get('herbicida_pre_aplicacoes') or 0)
            registro.herbicida_pos_alvo = request.form.get('herbicida_pos_alvo')
            registro.herbicida_pos_aplicacoes = int(request.form.get('herbicida_pos_aplicacoes') or 0)
            
            registro.tratamento_sementes = request.form.get('tratamento_sementes')
            registro.sal_mistura = request.form.get('sal_mistura')
            registro.controle_biologico = request.form.get('controle_biologico')
            
            registro.inoculacao_sementes = request.form.get('inoculacao_sementes')
            registro.forma_inoculacao = request.form.get('forma_inoculacao')
            registro.coinoculacao = request.form.get('coinoculacao')
            registro.co_mo = request.form.get('co_mo')
            registro.co_mo_aplicacao = request.form.get('co_mo_aplicacao')
            
            # Remover pulverizações antigas
            Pulverizacao.query.filter_by(formulario_id=registro.id).delete()
            
            # Adicionar novas pulverizações
            if request.form.get('data_pre_plantio'):
                pulv = Pulverizacao(
                    formulario_id=registro.id,
                    tipo='pre_plantio',
                    data=request.form.get('data_pre_plantio'),
                    classe_produto='Inseticida',
                    alvo=request.form.get('alvo_pre_plantio')
                )
                db.session.add(pulv)
            
            for i in range(1, 8):
                data = request.form.get(f'data_pos_{i}')
                if data:
                    classes = request.form.getlist(f'classe_pos_{i}')
if classes:
    classe_str = ', '.join(classes)  # Converte lista para string "Inseticida, Fungicida"
else:
    classe_str = ''
                    alvo = request.form.get(f'alvo_pos_{i}')
                    if data and classe_str:
    pulv = Pulverizacao(
        formulario_id=formulario.id,
        tipo=f'pos_{i}',
        data=data,
        classe_produto=classe_str,
        alvo=alvo
    )
                        db.session.add(pulv)
            
            db.session.commit()
            flash('Registro atualizado com sucesso!', 'success')
            return redirect(url_for('view_record', id=registro.id))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao atualizar: {str(e)}', 'danger')
    
    return render_template('edit_form.html', 
                      registro=registro, 
                      insetos=INSETOS_ALVO, 
                      doencas=DOENCAS_ALVO,
                      plantas=PLANTAS_DANINHAS,
                      acaros=ACAROS)
    
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
