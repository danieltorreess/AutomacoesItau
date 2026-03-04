from datetime import datetime, timedelta


def datas_para_busca():

    hoje = datetime.now().date()

    if hoje.weekday() == 0:  # segunda

        return [
            hoje - timedelta(days=2),
            hoje - timedelta(days=1),
            hoje
        ]

    return [hoje]