import pandas as pd
from io import BytesIO
from flask import send_file
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from models import db, FormularioSoja, Pulverizacao
import json
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import os
from export_excel import gerar_excel, orm_para_dict

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

def gerar_opcoes_safra():
    """Gera uma lista de safras no formato AAAA/AAAA, cobrindo alguns anos
    antes e depois do ano atual."""
    ano_atual = datetime.now().year
    return [f"{a}/{a + 1}" for a in range(ano_atual - 2, ano_atual + 6)]

@app.route('/escolher_safra')
def escolher_safra():
    return render_template('escolher_safra.html', safras=gerar_opcoes_safra())

@app.route('/form', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        try:
            # Criar novo formulário
            formulario = FormularioSoja()
            formulario.safra = request.form.get('safra')
            
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
            _advs = request.form.getlist('qual_adversidade')
            formulario.qual_adversidade = ', '.join(_advs) if _advs else None
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
            formulario.herbicida_pos_ns_alvo = request.form.get('herbicida_pos_ns_alvo')
            formulario.herbicida_pos_ns_aplicacoes = int(request.form.get('herbicida_pos_ns_aplicacoes') or 0)
            
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
            # Pré-plantio com múltiplas classes
            if request.form.get('data_pre_plantio'):  # <--- AGORA COM INDENTAÇÃO CORRETA!
                classes_pre = request.form.getlist('classe_pre_plantio')
                if classes_pre:
                    classe_pre_str = ', '.join(classes_pre)
                else:
                    classe_pre_str = ''
                
                alvos_pre = request.form.getlist('alvo_pre_plantio')
                alvo_pre = ', '.join(alvos_pre) if alvos_pre else ''
                
                if classe_pre_str and alvo_pre:
                    pulv = Pulverizacao(
                        formulario_id=formulario.id,
                        tipo='pre_plantio',
                        data=request.form.get('data_pre_plantio'),
                        classe_produto=classe_pre_str,
                        alvo=alvo_pre
                    )
                    db.session.add(pulv)
            
            # Pulverizações pós-emergência (até 7)
            for i in range(1, 8):
                data = request.form.get(f'data_pos_{i}')
                if data:
                    classes = request.form.getlist(f'classe_pos_{i}')
                    if classes:
                        classe_str = ', '.join(classes)
                    else:
                        classe_str = ''
                    
                    alvos = request.form.getlist(f'alvo_pos_{i}')
                    alvo = ', '.join(alvos) if alvos else ''
                    
                    if classe_str and alvo:
                        pulv = Pulverizacao(
                            formulario_id=formulario.id,
                            tipo=f'pos_{i}',
                            data=data,
                            classe_produto=classe_str,
                            alvo=alvo
                        )
                        db.session.add(pulv)
            
            db.session.commit()
            flash('Formulário salvo com sucesso!', 'success')
            return redirect(url_for('view_records'))
            
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao salvar: {str(e)}', 'danger')
            return redirect(url_for('form', safra=request.form.get('safra')))
    
    safra = request.args.get('safra')
    if not safra:
        return redirect(url_for('escolher_safra'))
    
    return render_template('form.html', 
                      safra=safra,
                      insetos=INSETOS_ALVO, 
                      doencas=DOENCAS_ALVO,
                      plantas=PLANTAS_DANINHAS,
                      acaros=ACAROS)

@app.route('/record/<int:id>')
def view_record(id):
    registro = FormularioSoja.query.get_or_404(id)
    return render_template('view_record.html', 
                          registro=registro, 
                          insetos=INSETOS_ALVO, 
                          doencas=DOENCAS_ALVO,
                          plantas=PLANTAS_DANINHAS,
                          acaros=ACAROS)

@app.route('/records')
def view_records():
    safra_filtro = request.args.get('safra')
    query = FormularioSoja.query.order_by(FormularioSoja.data_criacao.desc())
    if safra_filtro:
        query = query.filter_by(safra=safra_filtro)
    registros = query.all()
    safras_existentes = [
        s[0] for s in db.session.query(FormularioSoja.safra)
                                 .distinct()
                                 .order_by(FormularioSoja.safra.desc())
                                 .all()
        if s[0]
    ]
    return render_template('view_records.html',
                            registros=registros,
                            safras=safras_existentes,
                            safra_selecionada=safra_filtro)

@app.route("/exportar_excel")
def exportar_excel():
    safra = request.args.get('safra')
    query = FormularioSoja.query.order_by(FormularioSoja.id)
    if safra:
        query = query.filter_by(safra=safra)
    todos = query.all()
    registros = [orm_para_dict(r) for r in todos]
    filepath = os.path.join("/tmp", "MesoIDR_Export.xlsx")
    gerar_excel(registros, filepath)
    nome_arquivo = f"MesoIDR_Exportacao_{safra.replace('/', '-')}.xlsx" if safra else "MesoIDR_Exportacao_TodasSafras.xlsx"
    return send_file(filepath,
                     as_attachment=True,
                     download_name=nome_arquivo,
                     mimetype="application/vnd.openxmlformats-officedocument"
                              ".spreadsheetml.sheet")
    
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
            registro.safra = request.form.get('safra')
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
            _advs = request.form.getlist('qual_adversidade')
            registro.qual_adversidade = ', '.join(_advs) if _advs else None
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
            registro.herbicida_pos_ns_alvo = request.form.get('herbicida_pos_ns_alvo')
            registro.herbicida_pos_ns_aplicacoes = int(request.form.get('herbicida_pos_ns_aplicacoes') or 0)
            
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

            # Pré-plantio com múltiplas classes
            if request.form.get('data_pre_plantio'):  # <--- PRECISA DE 12 ESPAÇOS NO INÍCIO!
                  classes_pre = request.form.getlist('classe_pre_plantio')
                  if classes_pre:
                      classe_pre_str = ', '.join(classes_pre)
                  else:
                      classe_pre_str = ''
                  
                  alvos_pre = request.form.getlist('alvo_pre_plantio')
                  alvo_pre = ', '.join(alvos_pre) if alvos_pre else ''
                  
                  if classe_pre_str and alvo_pre:
                      pulv = Pulverizacao(
                          formulario_id=registro.id,
                          tipo='pre_plantio',
                          data=request.form.get('data_pre_plantio'),
                          classe_produto=classe_pre_str,
                          alvo=alvo_pre
                      )
                      db.session.add(pulv)
            
            for i in range(1, 8):
                data = request.form.get(f'data_pos_{i}')
                if data:
                    classes = request.form.getlist(f'classe_pos_{i}')
                    if classes:
                        classe_str = ', '.join(classes)
                    else:
                        classe_str = ''
                    
                    alvos = request.form.getlist(f'alvo_pos_{i}')
                    alvo = ', '.join(alvos) if alvos else ''
                    
                    if data and classe_str:
                        pulv = Pulverizacao(
                            formulario_id=registro.id,
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
                      safras=gerar_opcoes_safra(),
                      insetos=INSETOS_ALVO, 
                      doencas=DOENCAS_ALVO,
                      plantas=PLANTAS_DANINHAS,
                      acaros=ACAROS)
    
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
