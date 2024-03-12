import datetime

def obter_meses():
    hoje = datetime.date.today()
    
    # Nomes dos meses em português
    meses = [
        'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
        'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
    ]
    
    # Retornar o nome do mês atual em português
    MESONE = meses[hoje.month - 1]
    
    # Retornar o nome do mês passado em português
    primeiro_dia_mes_passado = (hoje.replace(day=1) - datetime.timedelta(days=1))
    MESTWO = meses[primeiro_dia_mes_passado.month - 1]
    
    # Retornar o nome do mês retrasado em português
    primeiro_dia_mes_retrasado = (primeiro_dia_mes_passado.replace(day=1) - datetime.timedelta(days=1))
    MESTREE = meses[primeiro_dia_mes_retrasado.month - 1]
    
    # Retornar o nome do mês anterior ao retrasado em português
    primeiro_dia_mes_ant_ret = (primeiro_dia_mes_retrasado.replace(day=1) - datetime.timedelta(days=1))
    MESFOUR = meses[primeiro_dia_mes_ant_ret.month - 1]
    
    return MESONE, MESTWO, MESTREE, MESFOUR

MESONE, MESTWO, MESTREE, MESFOUR = obter_meses()