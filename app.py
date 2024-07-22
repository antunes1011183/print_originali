from flask import Flask, request, jsonify, redirect, url_for, render_template
from flask_sqlalchemy import SQLAlchemy
from werkzeug.utils import secure_filename
import os
import csv
import chardet
import socket
import logging
from flask_paginate import Pagination, get_page_args

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///expedicao_novo.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
db = SQLAlchemy(app)

def generate_zpl(data, quantity=1):
    it_codigo = data['it_codigo']
    it_codigo_with_newlines = it_codigo[:6] + ' ' + it_codigo[6:]
    
    # Extrair o último número da string `nr_linha`
    nr_linha_split = data['nr_linha'].split('-')
    nr_linha_final = nr_linha_split[-1] if len(nr_linha_split) > 1 else data['nr_linha']
    
    zpl = f"""
    ^XA
    ~TA000
    ~JSN
    ^LT0
    ^MNW
    ^MTT
    ^PON
    ^PMN
    ^LH0,0
    ^JMA
    ^PR5,5
    ~SD8
    ^JUS
    ^LRN
    ^CI27
    ^PA0,1,1,0
    ^XZ
    ^XA
    ^MMT
    ^PW831
    ^LL575
    ^LS0
    ^FO13,16^GB803,545,4,,0^FS
    ^FO15,429^GB799,0,3^FS
    ^FO15,330^GB799,0,3^FS
    ^FO15,202^GB799,0,3^FS
    ^FO362,20^GB0,234,3^FS
    ^FO250,205^GB0,227,3^FS
    ^FO452,432^GB0,129,3^FS
    ^FO250,294^GB562,0,3^FS
    ^FO253,253^GB562,0,3^FS
    ^FO455,471^GB357,0,3^FS
    ^FO16,400^GB234,0,3^FS
    ^FO16,367^GB234,0,3^FS
    ^FO134,205^GB0,126,3^FS
    ^FT726,443^A0I,21,25^FH\\^CI28^FDMADE IN BRAZIL^FS^CI27
    ^FT444,536^A0I,21,23^FH\\^CI28^FDITALINEA INDUSTRIA DE MOVEIS LTDA^FS^CI27
    ^FT444,506^A@I,20,20,TT0003M_^FH\\^CI28^FDAlameda Todeschini, 370 - Bairro Verona, ^FS^CI27
    ^FT444,484^A@I,20,20,TT0003M_^FH\\^CI28^FDCEP: 95700-834 Bento Gonçalves/RS - Brasil,^FS^CI27
    ^FT444,465^A0I,17,20^FH\\^CI28^FDCNPJ: 02.017.451/0001-67^FS^CI27
    ^FT444,445^A0I,16,20^FH\\^CI28^FDFone: (54) 3455-5100^FS^CI27
    ^FT130,446^A0I,20,23^FH\\^CI28^FDORIGINALI^FS^CI27
    ^FT805,211^A0I,49,48^FH\\^CI28^FDPED:^FS^CI27
    ^FT359,204^A@I,38,38,TT0003M_^FH\\^CI28^FDL:^FS^CI27
    ^FT247,403^A@I,25,25,TT0003M_^FH\\^CI28^FDOC:^FS^CI27
    ^FT245,371^A@I,25,25,TT0003M_^FH\\^CI28^FDVol:^FS^CI27
    ^FT245,335^A@I,25,25,TT0003M_^FH\\^CI28^FDPC:^FS^CI27
    ^FO520,480^GFA,673,2048,32,:Z64:eJy1lE2O2zAMhSknBgbupovqAl0NNJfwEbQw72Nk1WMYWQnKJftIywrVqnGLQRlYsSl/evyRdaUrEc30tKt4qiPfc6KQ15BvOekAz6aOJPOeORLz7BbieZRBPX4eWfkp5zzknIZMeZ2SDPBQ2Ka8yvy4uGWcx7jzWEo94DEoD3LCKuBD4WWtsGFQPjq+khMGqPLwCM+V/0GXbPTTBdGDXwsvQovRj43+LmT1t1/05UV+6nvlY41/Uv5m9OEJqfIoGo1s9KWMRj8kQgFvRj9sVp+1CTZ/8URX678Oj/wA/1i/lPqv9LFdSv48k+No9eEx+pAZ0HCjL8JGX68mf6Im/yRYk39q8sevqb/82fqnIbX1F34zvCv8kb/wNX+oCY90G/1n/0Wd8Lqr/bf6g0YvK+SMPQtePeFe9F30kj2uhSP2PKvH81J4cOgWLnw9ABGEeN7ykT95af8iFddH3j1Vn4Z7uSTcd+jrU1gP3pqEsd99o5J/Y9p/vQtU+t8YYijUVzr0rWHjFeqdevqeDN/RD+vBv9HU0cfBI5tPbax3jaGSu32/b735v7f0u35jnfytTTWSvnmOL+fzfn7+0ZhfByCn6AtzJzzOhE4D/oE/1+9sgJY/0T/l/6f+Z+tP/Mn++/7+r3a2/3rfX2P7/v8JBuD+zA==:4916
    ^FPH,2^FT802,391^A@I,39,34,TT0003M_^FH\\^CI28^FD{data['nome_emit']}^FS^CI27
    ^FPH,3^FT805,345^A@I,39,40,TT0003M_^FH\\^CI28^FD{data['cidade']}^FS^CI27
    ^FT313,345^A@I,39,38,TT0003M_^FH\\^CI28^FD{data['estado']}^FS^CI27
    ^FT109,238^A0I,85,119^FH\\^CI28^FD{data['cod_box']}^FS^CI27
    ^FT339,89^A0I,93,119^FH\\^CI28^FD{data['cod_box']}^FS^CI27
    ^FT702,210^A0I,54,63^FH\\^CI28^FD{data['nr_pedido']}^FS^CI27
    ^FT802,153^A@I,37,40,TT0003M_^FH\\^CI28^FD{data['nome_transp']}^FS^CI27
    ^FT258,77^A0I,50,68^FH\\^CI28^FB300,2,0,C,0^FD{it_codigo_with_newlines}^FS^CI27
    ^FT240,276^A0I,50,63^FH\\^CI28^FD{data['nr_carga']}^FS^CI27
    ^FT319,212^A@I,40,38,TT0003M_^FH\\^CI28^FD{nr_linha_final}^FS^CI27
    ^FT238,213^A@I,39,36,TT0003M_^FH\\^CI28^FD{data['volume']}^FS^CI27
    ^FPH,3^FT802,308^A0I,20,18^FH\\^CI28^FD{data['desc_item_1']}^FS^CI27
    ^FT802,270^A0I,20,20^FH\\^CI28^FD{data['desc_item_2']}^FS^CI27
    ^FT198,408^A@I,25,25,TT0003M_^FH\\^CI28^FD{data['nr_pedrep']}^FS^CI27
    ^FT198,377^A@I,27,27,TT0003M_^FH\\^CI28^FD{data['nr_vol_ped']}^FS^CI27
    ^FT198,340^A@I,27,27,TT0003M_^FH\\^CI28^FD{data['nr_pedcli']}^FS^CI27
    ^BY2,3,83^FT779,52^BCI,,Y,N
    ^FH\\^FD>;{data['cod_barra']}^FS
    ^LRY^FO137,255^GB113,0,76^FS
    ^LRY^FO369,143^GB441,0,57^FS
    ^LRY^FO22,26^GB228,0,177^FS
    ^LRN
    ^PQ{quantity},0,1,Y
    ^XZ
    """
    return zpl

class ExpedicaoNovo(db.Model):
    __tablename__ = 'expedicao_novo'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    cod_barra = db.Column(db.String(50))
    nr_pedrep = db.Column(db.String(50))
    nr_vol_ped = db.Column(db.String(50))
    nome_transp = db.Column(db.String(100))
    nr_linha = db.Column(db.String(100))  
    nr_pedido = db.Column(db.String(50))
    nr_pedcli = db.Column(db.String(50))
    nr_carga = db.Column(db.String(50))
    it_codigo = db.Column(db.String(50))
    nr_volume = db.Column(db.String(50))
    nr_sequencia = db.Column(db.String(50))
    cod_box = db.Column(db.String(50))
    desc_item_1 = db.Column(db.String(255))
    desc_item_2 = db.Column(db.String(255))
    nome_emit = db.Column(db.String(100))
    cidade = db.Column(db.String(100))
    estado = db.Column(db.String(2))
    dt_embarque = db.Column(db.String(50))
    medida = db.Column(db.String(50))
    volume = db.Column(db.String(50))
    item_impresso = db.Column(db.String(3), default='não')  

with app.app_context():
    db.create_all()

def detectar_codificacao(file_path):
    with open(file_path, 'rb') as f:
        resultado = chardet.detect(f.read())
        return resultado['encoding']

def importar_csv_para_banco(file_path):
    encoding = detectar_codificacao(file_path)
    try:
        with open(file_path, mode='r', encoding=encoding) as file:
            csv_reader = csv.DictReader(file, delimiter=';')
            
            headers = csv_reader.fieldnames
            print(f"Cabeçalhos encontrados: {headers}")

            expected_headers = ['cod-barra', 'nr-pedrep', 'nr-vol-ped', 'nome-transp', 'nr-linha', 'nr-pedido', 'nr-pedcli', 'nr-carga', 'it-codigo', 'nr-volume', 'nr-sequencia', 'cod-box', 'desc-item-1', 'desc-item-2', 'nome-emit', 'cidade', 'estado', 'dt-embarque', 'medida', 'volume']
            
            if not all(header in headers for header in expected_headers):
                raise KeyError("Cabeçalhos do CSV não correspondem aos esperados.")
            
            for linha in csv_reader:
                filename = os.path.splitext(os.path.basename(file_path))[0]
                
                expedicao = ExpedicaoNovo(
                    cod_barra=linha['cod-barra'],
                    nr_pedrep=linha['nr-pedrep'],
                    nr_vol_ped=linha['nr-vol-ped'],
                    nome_transp=linha['nome-transp'],
                    nr_linha=filename,
                    nr_pedido=linha['nr-pedido'],
                    nr_pedcli=linha['nr-pedcli'],
                    nr_carga=linha['nr-carga'],
                    it_codigo=linha['it-codigo'],
                    nr_volume=linha['nr-volume'],
                    nr_sequencia=linha['nr-sequencia'],
                    cod_box=linha['cod-box'],
                    desc_item_1=linha['desc-item-1'],
                    desc_item_2=linha['desc-item-2'],
                    nome_emit=linha['nome-emit'],
                    cidade=linha['cidade'],
                    estado=linha['estado'],
                    dt_embarque=linha['dt-embarque'],
                    medida=linha['medida'],
                    volume=linha['volume'],
                    item_impresso='não'
                )
                db.session.add(expedicao)
            db.session.commit()
        print(f"Arquivo {file_path} importado com sucesso usando a codificação {encoding}.")
    except UnicodeDecodeError as e:
        print(f"Erro de decodificação com a codificação {encoding}: {e}")
    except KeyError as e:
        print(f"Erro de chave: {e}")

@app.route('/')
def index():
    return render_template('index.html')
@app.route('/upload', methods=['GET', 'POST'])
def upload_arquivos():
    if request.method == 'POST':
        if 'arquivos' not in request.files:
            return jsonify({"message": "Nenhum arquivo enviado"}), 400

        arquivos = request.files.getlist('arquivos')
        if not arquivos:
            return jsonify({"message": "Nenhum arquivo selecionado"}), 400

        for arquivo in arquivos:
            if arquivo.filename == '':
                continue

            filename = secure_filename(arquivo.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            arquivo.save(file_path)

            with app.app_context():
                importar_csv_para_banco(file_path)

        return redirect(url_for('index'))

    return render_template('upload.html')

@app.route('/visualizar')
def visualizar():
    page, per_page, offset = get_page_args(page_parameter='page', per_page_parameter='per_page')
    per_page = 10
    
    filters = {
        'cod_barra': request.args.get('cod_barra', ''),
        'cod_box': request.args.get('cod_box', ''),
        'desc_item_1': request.args.get('desc_item_1', ''),
        'desc_item_2': request.args.get('desc_item_2', ''),
        'dt_embarque': request.args.get('dt_embarque', ''),
        'it_codigo': request.args.get('it_codigo', ''),
        'medida': request.args.get('medida', ''),
        'nome_emit': request.args.get('nome_emit', ''),
        'nome_transp': request.args.get('nome_transp', ''),
        'nr_linha': request.args.get('nr_linha', ''),
        'nr_pedido': request.args.get('nr_pedido', ''),
        'nr_sequencia': request.args.get('nr_sequencia', ''),
        'nr_volume': request.args.get('nr_volume', ''),
        'volume': request.args.get('volume', ''),
        'item_impresso': request.args.get('item_impresso', '')
    }

    query = ExpedicaoNovo.query
    if filters['cod_barra']:
        query = query.filter(ExpedicaoNovo.cod_barra.like(f"%{filters['cod_barra']}%"))
    if filters['cod_box']:
        query = query.filter(ExpedicaoNovo.cod_box.like(f"%{filters['cod_box']}%"))
    if filters['desc_item_1']:
        query = query.filter(ExpedicaoNovo.desc_item_1.like(f"%{filters['desc_item_1']}%"))
    if filters['desc_item_2']:
        query = query.filter(ExpedicaoNovo.desc_item_2.like(f"%{filters['desc_item_2']}%"))
    if filters['dt_embarque']:
        query = query.filter(ExpedicaoNovo.dt_embarque.like(f"%{filters['dt_embarque']}%"))
    if filters['it_codigo']:
        query = query.filter(ExpedicaoNovo.it_codigo.like(f"%{filters['it_codigo']}%"))
    if filters['medida']:
        query = query.filter(ExpedicaoNovo.medida.like(f"%{filters['medida']}%"))
    if filters['nome_emit']:
        query = query.filter(ExpedicaoNovo.nome_emit.like(f"%{filters['nome_emit']}%"))
    if filters['nome_transp']:
        query = query.filter(ExpedicaoNovo.nome_transp.like(f"%{filters['nome_transp']}%"))
    if filters['nr_linha']:
        query = query.filter(ExpedicaoNovo.nr_linha.like(f"%{filters['nr_linha']}%"))
    if filters['nr_pedido']:
        query = query.filter(ExpedicaoNovo.nr_pedido.like(f"%{filters['nr_pedido']}%"))
    if filters['nr_sequencia']:
        query = query.filter(ExpedicaoNovo.nr_sequencia.like(f"%{filters['nr_sequencia']}%"))
    if filters['nr_volume']:
        query = query.filter(ExpedicaoNovo.nr_volume.like(f"%{filters['nr_volume']}%"))
    if filters['volume']:
        query = query.filter(ExpedicaoNovo.volume.like(f"%{filters['volume']}%"))
    if filters['item_impresso']:
        query = query.filter(ExpedicaoNovo.item_impresso == filters['item_impresso'])
    
    total = query.count()
    itens = query.offset(offset).limit(per_page).all()

    pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap5')
    
    return render_template('visualizar.html', itens=itens, page=page, per_page=per_page, pagination=pagination, filters=filters)

@app.route('/consulta', methods=['GET'])
def consulta():
    quantity = request.args.get('quantity', 1, type=int)
    cod_barra = request.args.get('cod_barra')
    
    if not cod_barra:
        return jsonify({"message": "Código de barra não fornecido"}), 400

    item = ExpedicaoNovo.query.filter_by(cod_barra=cod_barra).first()
    
    if not item:
        return jsonify({"message": "Item não encontrado"}), 404

    data = {
        'cod_barra': item.cod_barra,
        'nr_pedrep': item.nr_pedrep,
        'nr_vol_ped': item.nr_vol_ped,
        'nome_transp': item.nome_transp,
        'nr_linha': item.nr_linha,
        'nr_pedido': item.nr_pedido,
        'nr_pedcli': item.nr_pedcli,
        'nr_carga': item.nr_carga,
        'it_codigo': item.it_codigo,
        'nr_volume': item.nr_volume,
        'nr_sequencia': item.nr_sequencia,
        'cod_box': item.cod_box,
        'desc_item_1': item.desc_item_1,
        'desc_item_2': item.desc_item_2,
        'nome_emit': item.nome_emit,
        'cidade': item.cidade,
        'estado': item.estado,
        'dt_embarque': item.dt_embarque,
        'medida': item.medida,
        'volume': item.volume
    }

    zpl = generate_zpl(data, quantity)

    print_zpl_to_printer(zpl)

    item.item_impresso = 'sim'
    db.session.commit()
    
    return jsonify({"message": "Item impresso com sucesso"}), 200

def print_zpl_to_printer(zpl):
    printer_ip = '192.168.1.3'
    port = 9100
    try:
        logging.debug(f"Conectando à impressora {printer_ip} na porta {port}")
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((printer_ip, port))
            s.sendall(zpl.encode('utf-8'))
            logging.debug("ZPL enviado com sucesso")
    except socket.timeout:
        logging.error(f"Tempo de conexão esgotado ao tentar conectar-se à impressora {printer_ip}")
    except ConnectionRefusedError:
        logging.error(f"Conexão recusada ao tentar conectar-se à impressora {printer_ip}")
    except Exception as e:
        logging.error(f"Erro ao conectar-se à impressora: {e}")

if __name__ == "__main__":
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(host='0.0.0.0', debug=True)