import psycopg2

def conectar_banco_dados():
    try:
        conn = psycopg2.connect(
            host="179.124.193.130",
            port=5432,
            user="userwisemanager",
            password="pwdwisemanager",
            database="wisemanager" 
        )
        print("Conexão bem-sucedida ao banco de dados PostgreSQL.")
        return conn
    except (psycopg2.Error) as error:
        print("Erro ao se conectar ao banco de dados PostgreSQL:", error)


def consultar_dados_cadastro():
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        consulta = """
            SELECT 
                sg.id idSolicitacaogasto
                , REPLACE(REPLACE(REPLACE(cl.cpfcnpj, '.', ''), '/', ''), '-', '') AS cnpj
                , ep.empresa
                , ep.codMatriz
                , ep.chavesistema codEmpresa
                , sg.formapagamento
                , rc.descricao descricaoContab
                , rc.codigo
                , nf.naturezafinanceira
                , numeroparcelas
                , autorizarpa 
                , sg.numeroOs
                , cl.terceiro
                , cl.estado
                , sg.usarateiocentrocusto 
                , sg.valor as valorSg
                , sg.idrateiocc 
                , sg.historico
                , ep.id
            FROM 
                solicitacaoGasto sg
                , cliente cl
                , empresa ep
                , rateiocc rc
                , naturezafinanceira nf
                , anexosolicitacaogasto a
            WHERE 
                sg.aprovado = 'S'
                AND sg.enviaefetivacao = 'S'
                AND sg.efetivado = 'N'
                AND cl.id = sg.idFornecedor
                and sg.id = a.idsolicitacaogasto
                AND a.desconsiderAranexo = 'N'
                AND (a.tipo = 'DANFE' OR a.tipo = 'DANFE-ADTO')
                AND ep.id = sg.idEmpresa
                -- AND ep.idsegmento = 2
                AND sg.formapagamento <> 'Caixa Usado'
--                AND sg.formapagamento  = 'B'
                AND rc.id = sg.idcontabilizacaopadrao 
                AND nf.id = sg.idNaturezaFinanceira
                AND autorizaRpa = 'S'
                AND (a.notacaptadarpa = 'N' or a.notacaptadarpa is null)
            ORDER BY
                ep.empresa
                , sg.id
        """
        cursor.execute(consulta)
        resultados = cursor.fetchall()
        cursor.close()
        conn.close()
        return resultados

def consultar_nota_fiscal(id_solicitacao):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        consulta = """ 
            SELECT 
                TO_CHAR(valor, 'FM999999999D99') as valor
                , replace(TO_CHAR(date(dataemissaonota), 'DD-MM-YYYY'), '-', '') as dataemissaonota
                , numerodocto
                , serie
                , replace(TO_CHAR(date(datavencimento), 'DD-MM-YYYY'), '-','') as datavencimento
                , tipodoctonota
                , valorretinss as inss 
                , valorretir as ir 
                , valorretpiscofinscsl as piscofinscsl
                , valorretiss as iss 
            FROM 
                anexoSolicitacaoGasto
            WHERE 
                idSolicitacaoGasto = %s
                AND desconsiderAranexo = 'N'
                AND (tipo = 'DANFE' OR tipo = 'DANFE-ADTO')
                AND (notacaptadarpa = 'N' or notacaptadarpa is null)
            """
        cursor.execute(consulta, (id_solicitacao,))
        resultados = cursor.fetchall()
        cursor.close()
        conn.close()

        return resultados
    
def consultar_boleto(id_solicitacao):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        consulta = """ 
            SELECT 
                replace(TO_CHAR(date(datavencimento), 'DD-MM-YYYY'), '-','') as datavencimento
            FROM 
                anexoSolicitacaoGasto
            WHERE 
                idSolicitacaoGasto = %s
                AND desconsiderAranexo = 'N'
                AND tipo = 'Boleto'
                AND (notacaptadarpa = 'N' or notacaptadarpa is null)
            """
        cursor.execute(consulta, (id_solicitacao,))
        resultados = cursor.fetchall()
        cursor.close()
        conn.close()

        return resultados

def atualizar_anexosolicitacaogasto(numerodocto):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        update = """
            UPDATE 
                anexosolicitacaogasto
            SET 
                notacaptadarpa = 'S'
            WHERE 
                numerodocto = %s
                AND desconsiderAranexo = 'N'
                AND (tipo = 'DANFE' OR tipo = 'DANFE-ADTO')
        """
        cursor.execute(update, (numerodocto,))
        conn.commit()
        cursor.close()
        conn.close()
        print("Atualização realizada com sucesso.")

def consultar_rateio(id_solicitacao):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        query = """
            SELECT
                ccs.chavesistema centroCusto
                , sum(cr.valor) Valor
            FROM 
                centroscustorateadosolicitacaogasto cr,centrocustoshadow ccs
            WHERE
                cr.idsolicitacaogasto = %s
                AND ccs.id = cr.idcentrocustoshadow
            GROUP BY 
                ccs.chavesistema
        """
        cursor.execute(query, (id_solicitacao,))
        resultados = cursor.fetchall()
        conn.commit()
        cursor.close()
        conn.close()
        return resultados

def consultar_rateio_aut(id_rateiocc):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        query = """
        select
            cc.ChaveSistema
            , rc.percentual
        FROM 
            centrocustoshadow cc, rateiocentrocusto rc
        WHERE 
            rc.idrateiocc = %s
            AND cc.id = rc.idcentrocustoshadow
        """
        cursor.execute(query, (id_rateiocc,))
        resultados = cursor.fetchall()
        conn.commit()
        cursor.close()
        conn.close()
        return resultados    
    

def numero_parcelas(id_solicitacao, numerodocto):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        query = """
        select  
            replace(TO_CHAR(date(datavencimento), 'DD-MM-YYYY'), '-','') as datavencimento 
        from 
            anexosolicitacaogasto a 
        where 
            idsolicitacaogasto = %s
            and doccobranca = 'S' 
            and numerodocto = %s
            and desconsideraranexo = 'N'  
        order by 
        	replace(TO_CHAR(date(datavencimento), 'DD-MM-YYYY'), '-','')
        """
        cursor.execute(query, (id_solicitacao, numerodocto))
        resultados = cursor.fetchall()
        conn.commit()
        cursor.close()
        conn.close()
        return resultados

def autoriza_rpa_para_n(id_solicitacao):
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        update = """
            UPDATE 
                solicitacaoGasto
            SET 
                autorizaRpa = 'N'
            WHERE 
                solicitacaoGasto.id = %s
        """
        cursor.execute(update, (id_solicitacao,))
        conn.commit()
        cursor.close()
        conn.close()
        print("Atualização realizada com sucesso.")


def consultar_chat_id():
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        query = """
            select 
                idchattelegram 
            from 
                usuario 
            where 
                login = 'rpa'
        """
        cursor.execute(query)
        resultados = cursor.fetchall()
        # print("Resultados da consulta:", resultados)
        conn.commit()
        cursor.close()
        conn.close()
        return resultados
    
def consultar_token_bot():
    conn = conectar_banco_dados()
    if conn is not None:
        cursor = conn.cursor()
        query = """
            select 
                chaveavisotelegram 
            from 
                sistema 
            where 
                chaveavisotelegram is not null
        """
        cursor.execute(query)
        resultados = cursor.fetchall()
        # print("Resultados da consulta:", resultados)
        conn.commit()
        cursor.close()
        conn.close()
        return resultados
