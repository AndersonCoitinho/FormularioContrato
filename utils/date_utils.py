from datetime import datetime
import locale

def format_data_extenso(data_str):
    # Converter a data em um objeto datetime
    data = datetime.strptime(data_str, '%Y-%m-%d')

    # Definir a localidade para o idioma desejado (por exemplo, 'pt_BR' para Português do Brasil)
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except locale.Error:
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR')
        except locale.Error:
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
            except locale.Error:
                locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')    

    # Mapeando meses em inglês para português
    meses_em_portugues = {
        'January': 'Janeiro',
        'February': 'Fevereiro',
        'March': 'Março',
        'April': 'Abril',
        'May': 'Maio',
        'June': 'Junho',
        'July': 'Julho',
        'August': 'Agosto',
        'September': 'Setembro',
        'October': 'Outubro',
        'November': 'Novembro',
        'December': 'Dezembro'
    }    

    # Formatar a data por extenso
    data_extenso = data.strftime('%d de %B de %Y')  # %d: dia, %B: mês por extenso, %Y: ano
    for mes_ingles, mes_portugues in meses_em_portugues.items():
        data_extenso = data_extenso.replace(mes_ingles, mes_portugues)

    return data_extenso