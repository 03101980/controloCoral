import sqlite3 as lite
from sqlite3 import Error
import datetime

def eliminarMembros(id_membros):
    try:
        with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
            cur = con.cursor()
            query = "DELETE FROM membros WHERE id_membros = ?"
            # Executar a instrução SQL DELETE
            cur.execute(query, (id_membros,))
            # Confirmar as alterações no banco de dados        
            con.commit()
        return 1
    except Error as e:
        return e
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()

def conection():
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    cur = con.cursor()    
    #si ce le premier demarage, il inser le passe par defaut
    try:
        cur.execute("SELECT senha_login FROM conection")
        #il va compter le nombre de donneé dans la banc
        rows = cur.fetchall()
        for i in rows:
            i = i
        return i[0]
        
    except Exception as e:
        print(e)
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()

def NumeroEntier(numero, signe):
    numero2 = numero
    numero2 = str(numero2).replace(signe,"")#tirrer le separateur entre les chiffres
    return numero2 # renvoi un chiffre entier
   
def inserir_membro(nome, sexo, paroquia, morada, contacto, voz, foto, action, id_membros, data, estado):
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    cur = con.cursor()    
    #si ce le premier demarage, il inser le passe par defaut
    try:
        if action == 0:
            with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
                cur = con.cursor()
                cur.execute("INSERT INTO membros (nome_membros, sexo_membros, paroquia_membros, morada_membros, contacto_membros, voz_membros, foto_membros, data_ingresso_membros, estado_membros) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (nome, sexo, paroquia, morada, contacto, voz, foto, data, estado))
                con.commit()
                id_recent = cur.lastrowid      
                #inserir_membrosCota(id_recent)
            with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
                cur = con.cursor() 
                um, dois = executar_pagamento(cur, id_recent, 0, 8, 2023, data)
                um = um + dois
        else:
            with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
                cur = con.cursor()
                cur.execute("UPDATE membros SET nome_membros = ?, sexo_membros = ?, paroquia_membros = ?, morada_membros = ?, contacto_membros = ?, voz_membros = ?, foto_membros = ?, estado_membros = ? WHERE id_membros = ?", 
                            (nome, sexo, paroquia, morada, contacto, voz, foto, estado, id_membros))
                con.commit()
        return 1
    except Error as e:
        return e
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()

def activation_membro(id_membros, estado):
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    cur = con.cursor()    
    #si ce le premier demarage, il inser le passe par defaut
    try:
        with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
            cur = con.cursor()
            cur.execute("UPDATE membros SET estado_membros = ? WHERE id_membros = ?", (estado, id_membros))
            con.commit()
        return 1
    except Error as e:
        return e
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()

def inserir_membrosCota(id_membros):
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    cur = con.cursor()
    data_pagamento = datetime.date.today()
    mes = 9
    ano = 2023
    valor = 0
    try:
        cur.execute("INSERT INTO cotas (id_membro, valor, data_pagamento, mes_pagamento, ano_pagamento) VALUES (?, ?, ?, ?, ?)", (id_membros, valor, data_pagamento, mes, ano))
        con.commit()
        return 1
    except lite.Error as e:
        return e
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()         
def carregar_dados_do_banco(letra_procurada, estado):
    try:
        if estado == "Tudo":  
            consulta_sql = "SELECT nome_membros, paroquia_membros, morada_membros, contacto_membros, voz_membros, id_membros FROM membros WHERE nome_membros LIKE ? ORDER BY nome_membros"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{letra_procurada}%",))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        else:
            if estado == "Reciclagem":
                estado = "desactivado"
            consulta_sql = "SELECT nome_membros, paroquia_membros, morada_membros, contacto_membros, voz_membros, id_membros FROM membros WHERE nome_membros LIKE ? AND estado_membros = ? ORDER BY nome_membros"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{letra_procurada}%", estado))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        return dados_membros
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()
def carregar_dados_excel():
    try:   
        consulta_sql = "SELECT nome_membros, paroquia_membros, morada_membros, contacto_membros, voz_membros, data_ingresso_membros FROM membros"
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(consulta_sql,)  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        return dados_membros
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()
def preencher_comboBox():
    try:
        consulta_sql = "SELECT id_membros, nome_membros FROM membros ORDER BY nome_membros"
    
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(consulta_sql,)  # Adicionando % antes e depois para procurar nomes que contêm a letra
        nomes = cur.fetchall()
        return nomes
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
    finally:
        # Fechar a conexão com o banco de dados
        if con:
            con.close()
def carregar_cotas_nome(filtro, estado):
    try:
        if estado == "Tudo":  
            consulta_sql = "SELECT id_membros, nome_membros pas FROM membros WHERE nome_membros LIKE ? ORDER BY nome_membros"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{filtro}%",))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        else:
            if estado == "Reciclagem":
                estado = "desactivado"
            consulta_sql = "SELECT id_membros, nome_membros FROM membros WHERE nome_membros LIKE ? AND estado_membros = ? ORDER BY nome_membros"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{filtro}%", estado))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        return e
def carregar_cotas_valor(id_membro, mesinferior, messuperior, ano):
    try:   
        query = "SELECT valor FROM cotas WHERE id_membro = ? AND mes_pagamento > ? AND mes_pagamento < ? AND ano_pagamento = ?" 
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(query, (id_membro,mesinferior, messuperior, ano))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        return e
def mes_contribuido(id_doMembros):
    try:   
        query = "SELECT mes_pagamento, ano_pagamento, valor, id_pagamento FROM cotas WHERE id_membro = ? ORDER BY id_pagamento  DESC LIMIT 1" 
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(query, (id_doMembros,))
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        return e
def validar_cota(id_membros, total_a_pagar, valor_cota):
    total_a_pagar = int(total_a_pagar)
    valor_cota = int(valor_cota)
    data_pagamento = datetime.date.today()  
    try:
        pagamentoCota = inserir_caixa('Pagamento Cota', total_a_pagar, 0, 'entrada')  
        # Carregar o último pagamento
        dados = mes_contribuido(id_membros)
        data_pagamento = datetime.date.today()
        # Iniciar as variáveis
        reste = 0  # O resto depois do pagamento
        pagarDivida = 0
        contage = 0
        ultimo_pagamento = 0
        # Iterar sobre os dados do último pagamento
        for dado in dados:
            mes_pagamento = dado[0]
            ano_pagamento = dado[1]
            ultimo_pagamento = dado[2]
            id_pagamento = dado[3]
        # Verificar se houve um pagamento incompleto
        if 0 <= ultimo_pagamento < valor_cota:
            # Completar o último pagamento
            pagarDivida, totalpagar = completar_ultimo_pagamento(ultimo_pagamento, total_a_pagar, valor_cota)
            # Executar o pagamento
            with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
                cur = con.cursor()
                cur.execute("UPDATE cotas SET valor = ?, data_pagamento = ?, mes_pagamento = ?, ano_pagamento = ? WHERE id_pagamento = ?", 
                            (pagarDivida, data_pagamento, mes_pagamento, ano_pagamento, id_pagamento))
                con.commit()
        else:
            totalpagar = total_a_pagar
        # Somente novo pagamento
        contage, reste = calcular_contage_reste(totalpagar, valor_cota)
        with lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES) as con:
            cur = con.cursor()
            while contage > 0:
                # Executar os pagamentos
                #pagarDivida, total_a_pagar, mes_pagamento, ano_pagamento = executar_pagamento(cur, id_membros, valor_a_pagar, mes_pagamento, ano_pagamento, data_pagamento)
                mes_pagamento, ano_pagamento = executar_pagamento(cur, id_membros, valor_cota, mes_pagamento, ano_pagamento, data_pagamento)             
                contage -= 1
            if reste > 0:
                # Inserir também o resto, caso exista
                #pagarDivida, _, mes_pagamento, ano_pagamento = executar_pagamento(cur, id_membros, valor_a_pagar, mes_pagamento, ano_pagamento, data_pagamento)
                mes_pagamento, ano_pagamento =executar_pagamento(cur, id_membros, reste, mes_pagamento, ano_pagamento, data_pagamento)
        return 1, pagamentoCota
    except Error as e:
        return e
def completar_ultimo_pagamento(ultimo_pagamento, total_a_pagar, valorCota):
    # Completar o último pagamento    
    divida = valorCota - ultimo_pagamento # para saber o que reste para pagar
    if total_a_pagar >= divida: # o nosso valor é > a divida       
        total_a_pagar -= divida # retirar somente a divida
        pagarDivida = valorCota # valor normal do pagamento
    else:
        pagarDivida = total_a_pagar + ultimo_pagamento 
        total_a_pagar = 0
    return pagarDivida, total_a_pagar

def calcular_contage_reste(total_a_pagar, valor_limite):
    # Calcular o número de meses a serem pagos e o resto
    contage = total_a_pagar // valor_limite
    reste = total_a_pagar % valor_limite
    return contage, reste
def executar_pagamento(cur, id_membros, valor_a_pagar, mes_pagamento, ano_pagamento, data_pagamento):
    if mes_pagamento == 12:
        ano_pagamento += 1
        mes_pagamento = 1
    else:
        mes_pagamento += 1
    # Executar um pagamento
    cur.execute("INSERT INTO cotas (id_membro, valor, mes_pagamento, ano_pagamento, data_pagamento) VALUES (?, ?, ?, ?, ?)", 
                (id_membros, valor_a_pagar, mes_pagamento, ano_pagamento, data_pagamento))
    #pagarDivida = valor_a_pagar
    return mes_pagamento, ano_pagamento    
 
def inserir_caixa(motivo, entrada, saida, movim):
    data_pagamento = datetime.date.today()
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    try:
        with con:
            cur = con.cursor()        
            saldo = saldo_anterior()
            entrada = int(entrada)
            saida = int(saida)
            
            novo_saldo = saldo + entrada - saida
            if movim == "saida":
                if saldo or ((saldo / 4) * 3) > saida:
                    # Inserir entrada no caixa
                    cur.execute("INSERT INTO caixa (data, motivo, entrada, saida, saldo) VALUES (?, ?, ?, ?, ?)", (data_pagamento, motivo, entrada, saida, novo_saldo))
                    con.commit()
                    return 1
                else:
                    return 0
            else:
                # Inserir entrada no caixa
                cur.execute("INSERT INTO caixa (data, motivo, entrada, saida, saldo) VALUES (?, ?, ?, ?, ?)", (data_pagamento, motivo, entrada, saida, novo_saldo))
                con.commit()
                return 1
    except Error as e:
        return e
    finally:
        con.close()

def saldo_anterior():
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    try:
        with con:
            cur = con.cursor()
            # Obter o último saldo do caixa
            cur.execute("SELECT saldo FROM caixa ORDER BY id_caixa DESC LIMIT 1;")
            result = cur.fetchone()            
            saldo = 0
            if result:
                saldo = int(result[0])
        return saldo
    except Error as e:
        return []
    finally:
        con.close()

def ActualizarDadoDoGrupo(action, nome, cota, passe):
    con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
    try:
        with con:
            cur = con.cursor()
            if action == 1:
                cur.execute("UPDATE dadoGrupo SET nome_grupo = ?, pass_grupo = ?, valor_cota = ?", (nome, passe, cota))
            else:
                cur.execute("INSERT INTO dadoGrupo (nome_grupo, pass_grupo, valor_cota) VALUES (?, ?, ?)", (nome, passe, cota))
            con.commit()
        return 1
    except Error as e:
        return e
    finally:
        con.close()
def carregar_todos_dados(id_procurado):
    try:  
        consulta_sql = "SELECT * FROM membros WHERE id_membros = {};".format(id_procurado)
    
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(consulta_sql)  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
def caregarDadoGrupo():
    try:  
        consulta_sql = "SELECT * FROM dadoGrupo"
        con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
        cur = con.cursor()
        cur.execute(consulta_sql)  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
def carregar_dadosCaixa(motivo, data_inicial, data_final):
    try:
        if motivo == 'Outros':
            motivo = 'Pagamento Cota'
            consulta_sql = f"SELECT data, motivo, entrada, saida, saldo FROM caixa WHERE motivo != ? AND data BETWEEN '{data_inicial}' AND '{data_final}'"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"{motivo}",))  # Adicionando % antes e depois para procurar nomes que contêm a letra 
        else:
            if motivo == 'Tudo':
                motivo = ''
            consulta_sql = f"SELECT data, motivo, entrada, saida, saldo FROM caixa WHERE motivo LIKE ? AND data BETWEEN '{data_inicial}' AND '{data_final}'"   
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{motivo}%",))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        print("Erro ao carregar dados do banco:", e)
        return []
def carregar_dadosCaixaExcel(motivo, data_inicial, data_final):
    try:
        if motivo == 'Outros':
            motivo = 'Pagamento Cota'
            consulta_sql = f"SELECT data, motivo, entrada, saida FROM caixa WHERE motivo != ? AND data BETWEEN '{data_inicial}' AND '{data_final}'"
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"{motivo}",))  # Adicionando % antes e depois para procurar nomes que contêm a letra 
        else:
            if motivo == 'Tudo':
                motivo = ''

            consulta_sql = f"SELECT data, motivo, entrada, saida FROM caixa WHERE motivo LIKE ? AND data BETWEEN '{data_inicial}' AND '{data_final}'"   
            con = lite.connect("bd_ccdaDundo.db", detect_types=lite.PARSE_DECLTYPES | lite.PARSE_COLNAMES)
            cur = con.cursor()
            cur.execute(consulta_sql, (f"%{motivo}%",))  # Adicionando % antes e depois para procurar nomes que contêm a letra
        dados_membros = cur.fetchall()
        con.close()
        return dados_membros
    except Error as e:
        return []

