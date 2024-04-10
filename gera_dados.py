import csv
import random as rd
import xlsxwriter
from faker import Faker
from datetime import datetime, timedelta

fake = Faker('pt_BR')
n_deals = 7
n_tickets = 22
n_line_item = 22
# Função para gerar data de nascimento entre 18 e 65 anos atrás


def gera_data_assinatura_contrato():
    return fake.date_between('-30d', 'today').strftime('%d/%m/%Y')


def gera_nome_cliente():
    return fake.bs()


def gera_cnpj():
    return fake.cnpj()


def gera_tipo_negocio(tipo_negocio):
    i = rd.randint(0, len(tipo_negocio)-1)
    tipo_negocio_ret = tipo_negocio[i]
    return tipo_negocio_ret


def gera_fidelidade():
    return rd.randint(6, 12)


def gera_golive_plano():
    return fake.date_between('-30d', 'today').strftime('%d/%m/%Y')


def gera_status_mod(status_mod):
    i = rd.randint(0, len(status_mod)-1)
    status_mod_ret = status_mod[i]
    return status_mod_ret


def gera_ticket_id():
    ticket_id = rd.randint(0, n_tickets)
    return ticket_id


def gera_deal_id():
    ticket_id = rd.randint(1, n_deals)
    return ticket_id


def gera_valor_plano(valor_mod):
    i = rd.randint(0, len(valor_mod)-1)
    valor_mod_ret = valor_mod[i]
    return valor_mod_ret


def gera_nome_valor_modulo():
    info_mod = [{'nome': 'Boas Vindas', 'valor': 499},
                {'nome': 'Boleto', 'valor': 699},
                {'nome': 'Cobrança', 'valor': 999},
                {'nome': 'Falha Técnica', 'valor': 1299},
                {'nome': 'Vendedor Online', 'valor': 1499}]
    i = rd.randint(0, 4)
    return info_mod[i]['nome'], info_mod[i]['valor']


def transform_csv(data, filename):
    # Inserir dados em csv
    keys = data[0].keys()

    with open(filename, 'w', newline='') as output_file:
        dict_writer = csv.DictWriter(output_file, keys)
        dict_writer.writeheader()
        dict_writer.writerows(data)


def transform_excel(data, filename, k):
    # Inserir dados em excel
    columns = list(data[0].keys())
    rows = [list(result.values()) for result in data]

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(filename)

    date_format = workbook.add_format({'num_format': 'dd/mm/yy'})

    for row in range(len(rows)):
        for column in range(len(columns)):
            if row == 0:
                worksheet.write(row, column, columns[column])
            else:
                if column == k:
                    worksheet.write(
                        row, column, rows[row-1][column], date_format)
                else:
                    worksheet.write(row, column, rows[row-1][column])
    workbook.close()


deals = []
tipo_negocio = ['newbusiness', 'existingbusiness', 'Cross-Sell', 'Reversão']
for i in range(1, n_deals):
    if i == (n_deals-2):
        i = 1000
    deal_closedwon = {
        'deal_id': i,
        'prop.-closedate': gera_data_assinatura_contrato(),
        'prop.-dealname': gera_nome_cliente(),
        'prop.-cnpj': gera_cnpj(),
        'prop.-dealtype': gera_tipo_negocio(tipo_negocio),
        'prop.-fidelidade': gera_fidelidade()
    }
    deals.append(deal_closedwon)

ticket_deal = []
status_mod = ['Ativo', 'Inativo']
valor_mod = [499, 699, 999, 1299, 1499]
for i in range(1, n_tickets):
    ticket = {'ticket_id': i,
              'prop.-modulo_boleto_e_cobranca': gera_status_mod(status_mod),
              'prop.-modulo_vendedor_online': gera_status_mod(status_mod),
              'prop.-modulo_vendedor_online_e_boas_vindas': gera_status_mod(status_mod),
              'prop.-modulo_cobranca': gera_status_mod(status_mod),
              'prop.-modulo_falha_tecnica': gera_status_mod(status_mod),
              'prop.-data_entrega_modulo___boas_vindas': gera_golive_plano(),
              'prop.-data_go_live_modulo___boleto': gera_golive_plano(),
              'prop.-data_entrega_modulo___cobranca': gera_golive_plano(),
              'prop.-data_entrega_modulo___falha_tecnica': gera_golive_plano(),
              'prop.-data_entrega_modulo___vendedor_online': gera_golive_plano(),
              # 'prop.-amount': gera_valor_plano(valor_mod),
              'deal_id': gera_deal_id()
              }
    ticket_deal.append(ticket)

line_items = []
for i in range(1, n_line_item):
    nome_mod, valor_mod = gera_nome_valor_modulo()
    line_item = {'id': i,
                 'prop.-name': nome_mod,
                 'prop.-amount': valor_mod,
                 'deal_id': gera_deal_id()
                 }
    line_items.append(line_item)

data = [deals, ticket_deal, line_items]
file_names = ['deals', 'ticket_deal', 'line_items']
k_i = [1, 5, 5]
i = 0
for i in range(len(data)):
    #file_name = file_names[i]+'.csv'
    #transform_csv(data[i], file_name)

    file_name = file_names[i]+'.xlsx'
    transform_excel(data[i], file_name, k=k_i[i])
    i += 1
