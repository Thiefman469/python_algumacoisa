import os
import oracledb
import pandas as pd

try:
    conn = oracledb.connect(user= "rm572510", password= "181207", dsn='oracle.fiap.com.br:1521/ORCL')
    # Cursor: onde serão aplicadas as instruções SQL
    inst_cadastro = conn.cursor()
    inst_consulta = conn.cursor()
    inst_alteração = conn.cursor()
    inst_exclusão = conn.cursor()
except Exception as e:
    # informa um eventual erro
    print("Erro: ", e)
else:
    # flag para conexao ok
    conexao = True # ok

while conexao:
    os.system("cls")
    print("""
    ------ CRUD - PETSHOP ======
    0 - SAIR
    1 - Cadastrar pet
    2 - Listar pets
    3 - Editar pets
    4 - Excluir pets
    5 - Excluir todos os Pets
    """)
    escolha = int(input("Escolha: "))

    os.system("cls")

    match escolha:
        case 0 :
            break
        case 1: # Cadastrar pets
            try:
                print("------ CADASTRAR PETS ------")
                tipo = input("Tipo.....: ")
                nome = input("Nome.....: ")
                idade = int(input("Idade.....: "))

                # montar a instrucao SQL de cadastro
                cadastro = f"""
                INSERT INTO petshop
                (tipo_pet, nome_pet, idade)
                VALUES ('{tipo}', '{nome}', {idade})
                """

                # executar a gravação na tabela
                inst_exclusão.execute(cadastro)
                conn.commit()
            except Exception as e:
                print("Erro: ", e)
            else:
                print("\n### DADOS GRAVADOS COM SUCESSO ###")
        
        case 2: # listar os pets
            print("----- LISTAR PETS -----")
            lista_dados = []

            # montar a instrução de consulta
            inst_consulta.execute("SELECT * FROM petshop")

            data = inst_consulta.fetchall() # joga todos os dados no objeto data

            if data: # há dados no objeto data
                for dt in data:
                    lista_dados.append(dt)
            else: # Não retornou registros
                input("Tabela vazia!")
                continue

            # ordenar a lista
            lista_dados.sort()

            # Criando um Dataframe (gera uma planilha com os dados)
            dados_df = pd.DataFrame(
                lista_dados,
                columns=['Id', 'Tipo', 'Nome', 'Idade'],
            )

            print(dados_df.to_string(
                index=False,
                col_space={
                    "Id": 5,
                    "Tipo": 15,
                    "Nome": 25,
                    "Idade": 8
                }
            ))

            input("\Dados listados!")

        case 3: # Editar os pets
            try:
                print("------ EDITANDO DADOS DO PET ------")
                consulta = """
                SELECT id, nome_pet
                FROM petshop
                ORDER BY ID
                """

                inst_consulta.execute(consulta)

                data = inst_consulta.fetchall() # retorna os dados da consulta

                # se a tabela estiver vazia
                if not data:
                    input("A tabela está vaiza")
                    continue

                # transformar os dados em um dataframe
                dados_df = pd.DataFrame(
                    data,
                    columns=['Id', 'Nome']
                )

                print(dados_df.to_string(index=False, col_space=10))

                lista_dados = []

                pet_id = int(input("Id: "))

                consulta = f"""
                SELECT *
                FROM petshop
                WHERE id = {pet_id}
                """

                inst_consulta.execute(consulta)
                data = inst_consulta.fetchall() # retorna os dados da consulta

                if not data:
                    input("Id inexistente")
                    continue

                # preencher a lista
                for dt in data:
                    lista_dados.append(dt)

                novo_tipo = input("Novo tipo: ")
                novo_nome = input("Novo Nome: ")
                nova_idade = input("Nova idade: ")

                # Construção da instrução de edição (bars)
                alteracao = f"""
                UPDATE petshop
                SET tipo_pet='{novo_tipo}', nome_pet='{novo_nome}', idade={nova_idade}
                WHERE id = {pet_id}
                """
                inst_alteração.execute(alteracao)
                conn.commit()
            except Exception as e:
                print("Erro: ", e)
            else:
                print("\n##### DADOS ATUALIZADOS #####")

        case 4: # excluir os pets
            print("----- EXCLUIR PET -----")
            lista_dados = []

            pet_id = int(input("Id: "))

            consulta = f"SELECT * FOM petshop WHERE id = {pet_id}"
            inst_consulta.execute(consulta)
            data = inst_consulta.fetchall()

            for dt in data:
                lista_dados.append(dt)

            if len(lista_dados) == 0:
                print("Não há um pet cadastrado com este ID")
            else:
                # criacao da instrucao de exclusao
                exclusao = f"DELETE FROM  petshop WHERE id = {pet_id}"
                inst_exclusão.execute(exclusao)
                conn.commit
                input("##### PET EXCLUIDO #####")
        case 5: # excluir todos os pets
            print("!!!!! EXCLUIR TODOS OS PETS !!!!!")
            confirma = input("Confirma a exclusão? [S]IM OU [N]ÃO")
            if confirma.upper() == "S":
                exclusao = "DELETE FROM petshop"
                inst_exclusão.execute(exclusao)
                conn.commit()

                data_reset_ids = """ALTER TABLE petshop MODIFY(ID GENERATED AS IDENTITY (START WITH 1))"""
                inst_exclusão.execute(data_reset_ids)
                conn.commit()

            else:
                print("Exclusão cancelada pelo usuário!")
