import fdb
from datetime import date
from datetime import datetime
from unicodedata import normalize

def remove_acento(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

def dia_da_semana():
    #verifica qual eh o dia da semana
    # grava arquivo ULTIMADATA.txt
    indice_da_semana = date.weekday(date.today())
    global dia_semana
    #dia_da_semana = DIAS[indice_da_semana]
    
    if indice_da_semana == 0: #segunda feira
        with open('c:\\kochpalm\\ULTIMADATA.txt', 'w') as file:
            file.write(str(datetime.today())+'\n')
            file.write('Segunda Feira\n')
            file.write('checado\n')
            file.write('Gerado automaticamente pelo Script')
        #dia_semana = "1000000"
        return "1000000"
    elif indice_da_semana == 1: #terca feira
        with open('c:\\kochpalm\\ULTIMADATA.txt', 'w') as file:
            file.write(str(datetime.today())+'\n')
            file.write('Terca Feira\n')
            file.write('checado\n')
            file.write('Gerado automaticamente pelo Script')
        #dia_semana = "0100000"
        return "0100000"
    elif indice_da_semana == 2: #quarta feira
        with open('c:\\kochpalm\\ULTIMADATA.txt', 'w') as file:
            file.write(str(datetime.today())+'\n')
            file.write('Quarta Feira\n')
            file.write('checado\n')
            file.write('Gerado automaticamente pelo Script')
        #dia_semana = "0010000"
        return "0010000"
    elif indice_da_semana == 3: #quinta feira
        with open('c:\\kochpalm\\ULTIMADATA.txt', 'w') as file:
            file.write(str(datetime.today())+'\n')
            file.write('Quinta Feira\n')
            file.write('checado\n')
            file.write('Gerado automaticamente pelo Script')
        #return dia_semana
        return "0001000"
    else:
        #sexta feira
        with open('c:\\kochpalm\\ULTIMADATA.txt', 'w') as file:
            file.write(str(datetime.today())+'\n')
            file.write('Sexta Feira\n')
            file.write('checado\n')
            file.write('Gerado automaticamente pelo Script')
        #return dia_semana
        return "0000100"

def gera_produtos():
    try:
        print('Gerando arquivo PRODUTOS.TXT')
        lista_produtos = []
        con = fdb.connect(
            host='localhost', database='C:\\KOCH\\SD\\SDDB.FDB',
            user='SYSDBA', password='masterkey'
          )
        cur = con.cursor()
        cur.execute('select a.codite, a.descricao, a.qtdemb, a.unidade, a.codgru, COALESCE(a.preco1, 0.00), COALESCE(a.icms4, 0), COALESCE(a.restricao,0), a.obs, COALESCE(a.preco4,0), a.premin, a.quantidade1, a.validade, a.promocao, b.descricao from CADASTRO a, grupo b WHERE (a.ATIVO=1) and (a.codgru = b.codgru) AND COALESCE(a.QUANTIDADE1,0)>0 ORDER BY a.DESCRICAO')
        #0 - codite
        #1 - descricao
        #2 - qtdemb
        #3 - unidade
        #4 - codgru
        #5 - preco1
        #6 - icms4
        #7 - restricao
        #8 - obs
        #9 - preco4
        #10 - premin
        #11 - quantidade1
        #12 - validade
        #13 - promocao
        #14 - descricao tabela grupo
        produtos = cur.fetchall()
        print(produtos)

        for i in produtos:
            try:
                if (i[1] != '') and (i[5] > 0.00):
                    mquantidade = i[11]
                    if i[6] > 0.00:
                        mqtdcx = i[5] / i[6]
                        mqtdcx = float(mqtdcx).__round__(2)
                    else:
                        mqtdcx = 0
                    lista_produtos.append(str(i[0]) + ';')
                    if str(i[12]) != "None":
                        dataFormatada = i[12].strftime('%d/%m/%Y')
                        lista_produtos.append(str(i[1]+'('+dataFormatada+')'+';'))
                    else:
                        lista_produtos.append(str(i[1])+';')

                    lista_produtos.append(str(i[2]) + ';')
                    lista_produtos.append(str(i[3]) + ';')
                    lista_produtos.append(str(mquantidade) + ';')
                    if i[7] == '1': #quer dizer que esta marcado campo vencido
                        lista_produtos.append('Produto Vencido;')
                    elif str(i[8]) == 'None':
                        lista_produtos.append(' ;')
                    else:
                        lista_produtos.append(str(i[8]) + ';')

                    lista_produtos.append(str(i[5]).replace('.',',') + ';')  # tirei o filtro da promocao
                    lista_produtos.append(' ;') #nao sei onde vai isso no palm mas estava no layout do delphi
                    lista_produtos.append(str(mqtdcx).replace('.',',')+';')
                    lista_produtos.append(str(i[9]).replace('.',',') + ';')
                    lista_produtos.append(str(i[5]).replace('.',',') + ';')  # tirado preco de custo e colocado preco de venda para aparecer margem zerada
                    if i[14] != 'None':
                        lista_produtos.append('Forn.: ' + str(i[14]) + '/ CX Fech.: ' + str(i[13]) + ';')
                    else:
                        lista_produtos.append(';')
                    if i[10] > 0.00:
                        lista_produtos.append(str(i[10]).replace('.',',')+';'+'\n')
                    else:
                        lista_produtos.append('0,00;'+'\n')
            except EOFError:
                break

        with open('c:\\kochpalm\\PRODUTOS.txt', 'w') as f:
            f.writelines(lista_produtos)
        print('arquivo PRODUTOS.TXT salvo...')
        cur.close()
        con.close()
    except:
        print('Erro na geração do arquivo PRODUTOS.TXT')

def lista_vendedores():
    con = fdb.connect(
        host='localhost', database='C:\\KOCH\\SD\\SDDB.FDB',
        user='SYSDBA', password='masterkey'
    )
    ativos = []
    ativos.append('S')
    cur = con.cursor()
    cur.execute('select codven, semroteiro from vendedores where ativo=(?) order by codven',ativos,)
    vendedores = cur.fetchall()
    return vendedores

def gera_clientes():
    #Gera a base de clientes de cada vendedor com as cobrancas por dia da semana
#    try:
    print('Gerando arquivo CLIENTES.TXT')
    visita=[]
    clientes=[]
    vendedores = lista_vendedores()
    print(vendedores)
    dia_de_hj = dia_da_semana()

    for c in vendedores:
        con = fdb.connect(
            host='localhost', database='C:\\KOCH\\SD\\SDDB.FDB',
            user='SYSDBA', password='masterkey'
          )
        print('Gerando arquivo clientes do vendedor: '+str(c[0]))
        visita.append(c[0]) #coloco o codigo do vendedor na lista do dia para pegar os clientes do dia do vendedores especifico
        visita.append('S') #PALM = S
        visita.append('S') #ATIVO = S
        visita.append(dia_de_hj) #dia de hoje da semana
        cur = con.cursor()

        if str(c[1]) == 'S':
            cur.execute('select codcli, razao_social, nome_fantasia, endereco, numero, estado, cidade, bairro, complemento, cep, cgc, inscricao_estadual, fone1, fone2, contato, dias, tipocob from clientes WHERE  vendedor = (?) and palm=(?) and ativo=(?) ORDER BY CODCLI',visita, )
            print('Todos os clientes')
        else:
            visita.append('1111100') #coloquei para aparecer o nome dos vendedores nos palms sempre
            #cur.execute('select a.codven, b.codcli, b.razao_social, b.nome_fantasia, b.endereco, b.numero, b.estado, b.cidade, b.bairro, b.complemento, b.cep, b.cgc, b.inscricao_estadual, b.fone1, b.fone2, b.contato, b.dias, b.tipocob from vendedores a, clientes b WHERE a.ATIVO="S" and b.palm="S" and b.ativo="S" and (SubString(visita from 1 for 1)=''1'') ORDER BY CODCLI',visita,)
            #cur.execute('select b.codcli, b.razao_social, b.nome_fantasia, b.endereco, b.numero, b.estado, b.cidade, b.bairro, b.complemento, b.cep, b.cgc, b.inscricao_estadual, b.fone1, b.fone2, b.contato, b.dias, b.tipocob from vendedores a, clientes b WHERE a.ATIVO="S" and b.palm="S" and b.ativo="S" and b.visita =(?) ORDER BY CODCLI',visita, )
            cur.execute('select codcli, razao_social, nome_fantasia, endereco, numero, estado, cidade, bairro, complemento, cep, cgc, inscricao_estadual, fone1, fone2, contato, dias, tipocob from clientes WHERE  vendedor = (?) and palm=(?) and ativo=(?) and (visita=(?) or visita=(?)) ORDER BY CODCLI',visita, )
        visita=[]
        clientes = cur.fetchall()
        #0 - codcli
        #1 - razao_social
        #2 - nome_fantasia
        #3 - endereco
        #4 - numero
        #5 - estado
        #6 - cidade
        #7 - bairro
        #8 - complemento
        #9 - cep
        #10 - cgc
        #11 - inscricao_estadual
        #12 - fone1
        #13 - fone2
        #14 - contato
        #15 - dias
        #16 - tipocob

        arquivo_clientes = []
        mtotalcli = 0
        for i in clientes:
            arquivo_clientes.append(str(i[0])+';')
            #arquivo_clientes.append(i[1]+';')
            inc = str(i[1])
            inc = remove_acento(inc)
            arquivo_clientes.append(str(inc)+';')
            if str(i[2]) == 'None' or str(i[2]) == '' :
                arquivo_clientes.append(' ;')
            else:
                #arquivo_clientes.append(i[2]+';')
                inc = str(i[2])
                inc = remove_acento(inc)        
                arquivo_clientes.append(str(inc)+';')
            if str(i[3]) == 'None' or str(i[3]) == '':
                arquivo_clientes.append(' ,') #jogo virgula pq o endereço eh formado de rua + numero
            else:
                arquivo_clientes.append(str(i[3])+',')
                
            if str(i[4]) == 'None' or str(i[4]) == '':
                arquivo_clientes.append(';')
            else:
                arquivo_clientes.append(str(i[4])+';')
                
            if str(i[5]) != 'None':
                arquivo_clientes.append(str(i[5])+';')
            else:
                arquivo_clientes.append(' ;')
            if str(i[6]) == 'None' or str(i[6]) == '': 
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(i[6]+';')
            if str(i[7]) == 'None' or str(i[7]) == '':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[7])+';')
            if str(i[8]) == 'None' or str(i[8]) ==' ' or str(i[8]) == '':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[8])+';')
            if str(i[9]) == 'None' or str(i[9]) == '':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[9])+';')
            #arquivo_clientes.append(str(i[10])+';')
            if str(i[10]) == '' or str(i[10]) == 'None':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[10])+';')

            if str(i[11]) == '' or str(i[11]) == 'None':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[11])+';')
            if str(i[12]) =='' or str(i[12]) == 'None':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[12]).replace('.',' ')+';')
            if str(i[13]) =='' or str(i[13]) == 'None':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[13]).replace('.',' ')+';')
            if str(i[14]) =='' or str(i[14]) == 'None':
                arquivo_clientes.append(' ;')
            else:
                arquivo_clientes.append(str(i[14])+';')
            arquivo_clientes.append(str(c[0])+';')
            if str(i[15]) != 'None':
                arquivo_clientes.append(str(i[15])+';')
            else:
                arquivo_clientes.append('0;')
            if str(i[16]) != 'None' : #se tem prazo
                arquivo_clientes.append(str(i[16])+';')
            elif str(i[16]) == '': #se prazo em branco, tem que jogar espaço mas não entendo pq nao cai no caso de cima
                arquivo_clientes.append(' ;')# tem que jogar vazio pra nao dar erro = 06/09/2019
            else:
                arquivo_clientes.append('A;')
            arquivo_clientes.append(str(mtotalcli)+';'+'\n')# nao sei pra que serve

            #myString = ','.join(map(str, myList)) ***** COLOCA VIRGULA SEPARANDO OS ITENS DA LISTA
            if str(c[0]).__len__() < 2:
                x = '0'+str(c[0])
                with open(f'c:\\kochpalm\\{x}\\envia\\CLIENTES.txt', 'w') as f:
                    f.writelines(arquivo_clientes)
            else:
                with open(f'c:\\kochpalm\\{c[0]}\\envia\\CLIENTES.txt', 'w') as f:
                    f.writelines(arquivo_clientes)
        print('arquivo CLIENTES.TXT gerado com sucesso...')
        cur.close()
        con.close()

        #chama a geracao do MOVCLI.txt
        gera_movcli(clientes, str(c[0]))

    # except:
    #     print('Erro na geração dos arquivos CLINTES_py.TXT')

def gera_movcli(clientes, codven):
    codcli = []
    lista_movcli = []
    for c in clientes:
        codcli.append(str(c[0])) # codigo do cliente que veio do select da listagem dos clientes do dia
        con = fdb.connect(
            host='localhost', database='C:\\KOCH\\SD\\SDDB.FDB',
            user='SYSDBA', password='masterkey'
          )
        cur = con.cursor()
        cur.execute('select chave, codcli, datavenc, valor, tipo, parcela from MOVCLI WHERE CODCLI=(?) AND COALESCE(VALORPAG,0.00) < COALESCE(VALOR,0.00) ORDER BY DATAVENC',codcli,)
        movcli = cur.fetchall()
        for z in movcli:
            #0 - chave
            #1 - codcli
            #2 - datavenc
            #3 - valor
            #4 - tipo
            #5 - parcela
            lista_movcli.append(str(z[0])+';')
            lista_movcli.append(str(z[1])+';')
            #lista_movcli.append(str(z[2])+';')
            dataFormatada = z[2].strftime('%d/%m/%Y')
            lista_movcli.append(dataFormatada+';')
            lista_movcli.append(str(z[3]).replace('.',',')+';')
            if str(z[4]) == '' or str(z[4]) == 'None':
                lista_movcli.append(' ;')
            else:
                lista_movcli.append(str(z[4])+';')
            lista_movcli.append(str(z[5])+';'+'\n')
            if str(codven).__len__() < 2:
                x = '0'+str(codven)
                with open(f'c:\\kochpalm\\{x}\\envia\\MOVCLI.txt', 'w') as f:
                    f.writelines(lista_movcli)
            else:
                with open(f'c:\\kochpalm\\{codven}\\envia\\MOVCLI.txt', 'w') as f:
                    f.writelines(lista_movcli)
        cur.close()
        con.close()
        codcli = []
    print('arquivo MOVCLI.TXT gerado com sucesso...')

def gera_mensagem():
    #Gera o arquivo mensagem.txt
    dia_hj = date.weekday(date.today())
    if dia_hj == 0: #segunda feira
        with open('c:\\kochpalm\\mensagem.txt', 'w') as file:
            file.write(str(datetime.today())+' Hoje e Segunda Feira - Gerado automaticamente pelo Script; ; ; ;')

    elif dia_hj == 1: #terca feira
        with open('c:\\kochpalm\\mensagem.txt', 'w') as file:
            file.write(str(datetime.today())+' Hoje e Terça Feira - Gerado automaticamente pelo Script; ; ; ;')

    elif dia_hj == 2: #quarta feira
        with open('c:\\kochpalm\\mensagem.txt', 'w') as file:
            file.write(str(datetime.today())+' Hoje e Quarta Feira - Gerado automaticamente pelo Script; ; ; ;')

    elif dia_hj == 3: #quinta feira
        with open('c:\\kochpalm\\mensagem.txt', 'w') as file:
            file.write(str(datetime.today())+' Hoje e Quinta Feira - Gerado automaticamente pelo Script; ; ; ;')

    elif dia_hj == 4: #sexta feira
        with open('c:\\kochpalm\\mensagem.txt', 'w') as file:
            file.write(str(datetime.today())+' Hoje e Sexta Feira - Gerado automaticamente pelo Script; ; ; ;')
    print("arquivo mensagem.txt gerado")
    
def main():
    print('*' * 39)
    print('* Gerador Automático Dados Vendedores *')
    print('*' * 39)
    #conectar()
    #gera arquivo produtos.txt
    gera_produtos()
    #gera arquivo clientes.txt de cada vendedor
    gera_clientes()
    gera_mensagem()


if __name__ == "__main__":
    main()
