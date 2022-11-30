import numpy as np
import pandas as pd

# variáveis utilizadas

arq_consumo = 'consumo.xls'

tabela_consumo = ['CONSUMO POR UF', 'CONSUMO CATIVO POR UF', 'CONSUMO RESIDENCIAL POR UF', 
          'CONSUMO INDUSTRIAL POR UF', 'CONSUMO OUTROS POR UF']

tabela_consumidores = ['CONSUMIDORES RESIDENCIAIS POR F', 'CONSUMIDORES_INDUSTRIAIS_POR_F', 
                       'CONSUMIDORES_COMERCIAIS_POR_F', 'CONSUMIDORES_OUTROS_POR_F']

lst_estado = ['Rondônia',
              'Acre',
              'Amazonas', 
              'Roraima', 
              'Pará', 
              'Amapá', 
              'Tocantins', 
              'Maranhão', 
              'Piauí', 
              'Ceará', 
              'Rio Grande do Norte',
              'Paraíba',
              'Pernambuco',
              'Alagoas',
              'Sergipe',
              'Bahia', 
              'Minas Gerais', 
              'Espírito Santo',
              'Rio de Janeiro',
              'São Paulo',
              'Paraná',
              'Santa Catarina',
              'Rio Grande do Sul',
              'Mato Grosso do Sul',
              'Mato Grosso', 
              'Goiás',
              'Distrito Federal']


tp_consumo = {
    "total"  : 0,
    "cativo" : 1,
    "residencial" : 2,
    "industrial" : 3,
    "outros" : 4
}

nbr_consumo = {
    "residencial" : 0,
    "comercial" : 1,
    "industrial" : 2,
    "outros" : 3
}

#######################

#######################
# Funções
#######################

def to_letter (dado):
    return chr(dado + 96)

def ano_em_range_excel (ano, ano0=2004, nbr=26, nbr_meses=12, debug=True):
    inicio = (ano-ano0)*12+1 # começando em B
    inicioc = to_letter(inicio+1) if (inicio/nbr) < 1 else to_letter(int(inicio/nbr)) + to_letter(inicio%nbr + 1)
    fim=inicio+nbr_meses-1
    fimc = to_letter(fim+1) if (fim/nbr) < 1 else to_letter(int(fim/nbr)) + to_letter(fim%nbr + 1)
    if debug:
        print ('ano_em_range_excel = ',inicioc+":"+fimc)
    return inicioc+":"+fimc

def ler_ano_consumo_total_por_estado (idx_tabela_consumo, ano, debug=False):
    tmp=ano_em_range_excel(ano, debug=debug)
    if debug:
        print ('tmp =',tmp)
    try:
        df = pd.read_excel(arq_consumo,
                       sheet_name=tabela_consumo[idx_tabela_consumo], 
                       header=[5], # 6a linha
                       #index_col=0, # index está na 1a coluna
                       usecols="A,"+tmp, # pega o ano certo
                       nrows=28, # pega todos os estados
                       names=['Estado', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', "Jul", 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                      )
    except Exception as e:
        print (e)
    df['Ano'] = ano
    return df    

def consumo_total_por_estado (idx_tabela_consumo, ano=np.arange(2004,2021), debug=False):
    if isinstance(ano, int):
        df = ler_ano_consumo_total_por_estado (idx_tabela_consumo, ano, debug=debug)
        return df
    else:
        try:
            lst = list()
            if debug:
                print (type(ano))
            
            if isinstance(ano, np.ndarray):
                if debug:
                    print ('é ndarray')
                it = np.nditer(ano)
            else:
                if debug:
                    print ('não é ndarray')
                it=iter(ano)
                
            for anoi in it:
                if debug:
                    print ('anoi =',anoi)
                tmpdf = ler_ano_consumo_total_por_estado (idx_tabela_consumo, anoi, debug=debug) 
                if debug:
                    print ('deu', anoi)
                lst.append(tmpdf)
                
            df = pd.concat(lst)
            return df                
        except:
            print ('parametro \'ano\' precisa ser inteiro, lista de inteiros ou numpy.ndarray')
            return None
        

def ler_ano_consumo_total_por_estado_long (idx_tabela_consumo, ano, debug=False):
    tmp=ano_em_range_excel(ano, debug=debug)
    if debug:
        print ('tmp =',tmp)
    try:
        df = pd.read_excel(arq_consumo,
                       sheet_name=tabela_consumo[idx_tabela_consumo], 
                       header=[5], # 6a linha
                       #index_col=0, # index está na 1a coluna
                       usecols="A,"+tmp, # pega o ano certo
                       nrows=28, # pega todos os estados
                       names=['Estado', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', "Jul", 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                      )
    except Exception as e:
        print (e)
    df['Ano'] = ano
    return df            
        
def consumo_total_por_estado_long (idx_tabela_consumo, ano=np.arange(2004,2021), debug=False):
    if isinstance(ano, int):
        df = ler_ano_consumo_total_por_estado_long (idx_tabela_consumo, ano, debug=debug)
        return df
    else:
        try:
            lst = list()
            if debug:
                print (type(ano))
            
            if isinstance(ano, np.ndarray):
                if debug:
                    print ('é ndarray')
                it = np.nditer(ano)
            else:
                if debug:
                    print ('não é ndarray')
                it=iter(ano)
                
            for anoi in it:
                if debug:
                    print ('anoi =',anoi)
                tmpdf = ler_ano_consumo_total_por_estado_long (idx_tabela_consumo, anoi, debug=debug) 
                if debug:
                    print ('deu', anoi)
                lst.append(tmpdf)
                
            df = pd.concat(lst)
            return df                
        except:
            print ('parametro \'ano\' precisa ser inteiro, lista de inteiros ou numpy.ndarray')
            return None
        
        
def ler_ano_consumidores_por_estado (idx_tabela_consumidores, ano, debug=False):
    tmp=ano_em_range_excel(ano, debug=debug)
    df = pd.read_excel(arq_consumo, 
                  sheet_name=tabela_consumidores[idx_tabela_consumidores],
                  header=[5], 
                  #index_col=0,
                  usecols="A,"+tmp,
                  nrows=28,
                  names=['Estado', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', "Jul", 'Ago', 'Set', 'Out', 'Nov', 'Dez']
                 )
    df['Ano'] = ano
    return df        
        
def consumidores_por_estado (idx_tabela_consumidores, ano=np.arange(2004,2021)):
    if isinstance(ano, int):
        df = ler_ano_consumidores_por_estado (idx_tabela_consumidores, ano)
        return df
    else:
        try:
            lst = list()
            it=iter(ano)
            for anoi in it:
                tmpdf = ler_ano_consumidores_por_estado (idx_tabela_consumidores, anoi) 
                lst.append(tmpdf)
            df = pd.concat(lst)
            return df                
        except:
            print ('parametro \'ano\' precisa ser inteiro ou sequência de inteiros')
            return None

        
def pega_estado_idx (matricula):
    mat = pd.read_csv('uga', sep=",", names=['matricula'])
    lst = mat.index[mat['matricula']==matricula].to_list()
    if (len(lst) == 0):
        print ("O número de matrícula não é válido")
        return None
    elif (len(lst) > 1):
        print ("Procura encontrou mais de uma matrícula válida. Fale com o professor caso o problema persista")
        return None
    else:
        return lst[0]
