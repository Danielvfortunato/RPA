import psycopg2

def conectar_banco_dados():
    try:
        conn = psycopg2.connect(
            host="179.124.193.185",
            port=5432,
            user="userwisemanager",
            password="pwdwisemanager",
            database="wisemanager" 
        )
        print("Conexão bem-sucedida ao banco de dados PostgreSQL.")
        
        return conn
    
    except (psycopg2.Error) as error:
        print("Erro ao se conectar ao banco de dados PostgreSQL:", error)

def consultar_id_solicitacao():
    conn = conectar_banco_dados()
    
    if conn is not None:
        cursor = conn.cursor()
        consulta = """
            SELECT 
                id 
            FROM 
                solicitacaogasto 
            WHERE 
                aprovado = 'S' 
                AND enviaefetivacao = 'S' 
                AND efetivado = 'N' 
                AND autorizaRpa = 'S'
            ORDER BY 
                id DESC
            """
        cursor.execute(consulta)
        resultados = cursor.fetchall()
        # for row in resultados:
        #     print("ID:", row[0])
        cursor.close()
        conn.close()
        
        return resultados

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
            FROM 
                solicitacaoGasto sg
                , cliente cl
                , empresa ep
                , rateiocc rc
                , naturezafinanceira nf
            WHERE 
                sg.aprovado = 'S'
                AND sg.enviaefetivacao = 'S'
                AND sg.id = 135868
                AND sg.efetivado = 'N'
                AND cl.id = sg.idFornecedor
                AND ep.id = sg.idEmpresa
                AND ep.idsegmento = 2
                AND sg.formapagamento <> 'Caixa Usado'
                AND sg.formapagamento  = 'B'
                AND rc.id = sg.idrateiocc
                AND nf.id = sg.idNaturezaFinanceira
                -- and empresa = 'GERACAO - LAGES'
                -- AND autorizaRpa = 'S'
                AND (select count(*) from anexosolicitacaogasto where idsolicitacaogasto = sg.id and dataemissaonota is null) > 0
            order by 
                ep.chavesistema
                LIMIT 100
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
                TO_CHAR(valor, 'FM999999999D99') as valor,
                TO_CHAR(date(dataemissaonota), 'DD-MM-YYYY') as dataemissaonota,
                numerodocto,
                serie,
                datavencimento,
                tipodoctonota
            FROM 
                anexoSolicitacaoGasto
            WHERE 
                idSolicitacaoGasto = %s
                AND desconsiderAranexo = 'N'
                AND (tipo = 'DANFE' OR tipo = 'DANFE-ADTO')
                Limit 1
            """
        cursor.execute(consulta, (id_solicitacao,))
        resultados = cursor.fetchall()
        cursor.close()
        conn.close()

        return resultados

    

# retorno = consultar_dados_cadastro()
# idSolicitacaogasto, cnpj, empresa, codEmpresa, formapagamento, descricaoContab, codigo, naturezafinanceira = retorno

# print("idSolicitacaogasto:", idSolicitacaogasto)
# print("cnpj:", cnpj)
# print("empresa:", empresa)
# print("codEmpresa:", codEmpresa)
# print("formapagamento:", formapagamento)
# print("descricaoContab:", descricaoContab)
# print("codigo:", codigo)
# print("naturezafinanceira:", naturezafinanceira)

