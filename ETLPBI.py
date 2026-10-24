import pandas as pd
import openpyxl as xl
import numpy as np
import matplotlib as pl
import warnings
from openpyxl import load_workbook
from datetime import datetime, timedelta   

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# tratamento base fase de transposição

classificão_dict = {
    'Itens do check list': 'Classificação',
    'Data da Monitoria': 'Dados monitoria',
    'Data do Atendimento': 'Dados monitoria',
    'Matrícula do Consultor Monitorado': 'Dados monitoria',
    'Nome do Avaliador': 'Dados monitoria',
    'Origem do Atendimento': 'Dados monitoria',
    'Canal de Atendimento': 'Dados monitoria',
    'Título': 'Dados monitoria',
    'Condução - Seguiu script de abertura?': 'Processo',
    'Condução - Se apresentou de forma educada e profissional?': 'Humano',
    'Condução - Manteve comunicação assertiva, correta e formal?': 'Humano',
    'Condução - Personalizou o atendimento?': 'Humano',
    'Condução - Conduziu o atendimento sem abandono?': 'Humano',
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 'Humano',
    'Condução - Seguiu script de encerramento?': 'Processo',
    'Consulta Agendamento - Atendimento citou processo de agendamento de consulta?': 'Ínicio de bloco',
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 'Processo',
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 'Processo',
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 'Processo',
    'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?': 'Ínicio de bloco',
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Processo',
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Processo',
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 'Processo',
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Processo',
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 'Processo',
    'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Processo',
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Processo',
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 'Processo',
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Processo',
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 'Processo',
    'Consulta Cancelamento - Atendimento citou processo de cancelamento de consulta?': 'Ínicio de bloco',
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Consulta Cancelamento - Ofertou nova data?': 'Processo',
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 'Processo',
    'Exame Agendamento - Atendimento citou processo de agendamento de exame?': 'Ínicio de bloco',
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 'Processo',
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 'Processo',
    'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 'Processo',
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 'Processo',
    'Exame Agendamento - Ofereceu opções para o agendamento?': 'Processo',
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 'Processo',
    'Exame Agendamento - Orientou sobre o preparo de exame?': 'Processo',
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 'Processo',
    'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?': 'Ínicio de bloco',
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 'Processo',
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 'Processo',
    'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Processo',
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 'Processo',
    'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 'Processo',
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 'Processo',
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 'Processo',
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 'Processo',
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 'Processo',
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 'Processo',
    'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Processo',
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 'Processo',
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 'Processo',
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 'Processo',
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 'Processo',
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 'Processo',
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 'Processo',
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 'Processo',
    'Exame Cancelamento - Atendimento citou processo de cancelamento de exame?': 'Ínicio de bloco',
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Processo',
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Processo',
    'Exame Cancelamento - Atendente ofertou nova data?': 'Processo',
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 'Processo',
    'Tabulação do atendimento:': 'Processo',
    'Visão Avaliador - Nota Csat do atendimento (0 a 5):': 'Item CX visão avaliador',
    'Visão Cliente - Atendimento foi bom?': 'Item CX visão avaliador',
    'Visão Processos - Processos atenderam a necessidade do cliente?': 'Item CX visão avaliador',
    'Observações sobre processos não concluídos:': 'Item CX visão avaliador',
    'Classificação (0-5)': 'Item CX visão avaliador',
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 'Processo',
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 'Processo',
    'Condução - Acordo de Pausa': 'Processo',
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 'Humano',
    'Tipo de Item': 'Dados monitoria',
    'Caminho': 'Dados monitoria',
    'Unidades citadas no atendimento': 'Dados monitoria',
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 'Humano'
}

responsabilidade_dict = {
    'Itens do check list': 'Responsabilidade',
    'Data da Monitoria': 'Dados monitoria',
    'Data do Atendimento': 'Dados monitoria',
    'Matrícula do Consultor Monitorado': 'Dados monitoria',
    'Nome do Avaliador': 'Dados monitoria',
    'Origem do Atendimento': 'Dados monitoria',
    'Canal de Atendimento': 'Dados monitoria',
    'Título': 'Dados monitoria',
    'Condução - Seguiu script de abertura?': 'Humano',
    'Condução - Se apresentou de forma educada e profissional?': 'Humano',
    'Condução - Manteve comunicação assertiva, correta e formal?': 'Humano',
    'Condução - Personalizou o atendimento?': 'Humano',
    'Condução - Conduziu o atendimento sem abandono?': 'Humano',
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 'Humano',
    'Condução - Seguiu script de encerramento?': 'Humano',
    'Consulta Agendamento - Atendimento citou processo de agendamento de consulta?': 'Ínicio de bloco',
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 'Regra de negocio',
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 'Humano',
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 'Humano',
    'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?': 'Ínicio de bloco',
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Regra de negocio',
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Humano',
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 'Regra de negocio',
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Regra de negocio',
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 'Humano',
    'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Regra de negocio',
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Humano',
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 'Regra de negocio',
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Regra de negocio',
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 'Humano',
    'Consulta Cancelamento - Atendimento citou processo de cancelamento de consulta?': 'Ínicio de bloco',
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Consulta Cancelamento - Ofertou nova data?': 'Humano',
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 'Humano',
    'Exame Agendamento - Atendimento citou processo de agendamento de exame?': 'Ínicio de bloco',
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 'Regra de negocio',
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 'Humano',
    'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 'Regra de negocio',
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 'Humano',
    'Exame Agendamento - Ofereceu opções para o agendamento?': 'Humano',
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 'Humano',
    'Exame Agendamento - Orientou sobre o preparo de exame?': 'Humano',
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 'Humano',
    'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?': 'Ínicio de bloco',
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 'Humano',
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 'Humano',
    'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Regra de negocio',
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 'Humano',
    'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 'Regra de negocio',
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 'Humano',
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 'Humano',
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 'Humano',
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 'Humano',
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 'Humano',
    'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Regra de negocio',
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 'Humano',
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 'Regra de negocio',
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 'Humano',
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 'Humano',
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 'Humano',
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 'Humano',
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 'Humano',
    'Exame Cancelamento - Atendimento citou processo de cancelamento de exame?': 'Ínicio de bloco',
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Humano',
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Humano',
    'Exame Cancelamento - Atendente ofertou nova data?': 'Humano',
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 'Humano',
    'Tabulação do atendimento:': 'Humano',
    'Visão Avaliador - Nota Csat do atendimento (0 a 5):': 'Item CX visão avaliador',
    'Visão Cliente - Atendimento foi bom?': 'Item CX visão avaliador',
    'Visão Processos - Processos atenderam a necessidade do cliente?': 'Item CX visão avaliador',
    'Observações sobre processos não concluídos:': 'Item CX visão avaliador',
    'Classificação (0-5)': 'Item CX visão avaliador',
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 'Humano',
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 'Humano',
    'Condução - Acordo de Pausa': 'Humano',
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 'Humano',
    'Tipo de Item': 'Dados monitoria',
    'Caminho': 'Dados monitoria',
    'Unidades citadas no atendimento': 'Dados monitoria',
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 'Humano'
}

status_responsabilidade_dict = {
    'Itens do check list': 'Status responsabilidade',
    'Data da Monitoria': 'Dados monitoria',
    'Data do Atendimento': 'Dados monitoria',
    'Matrícula do Consultor Monitorado': 'Dados monitoria',
    'Nome do Avaliador': 'Dados monitoria',
    'Origem do Atendimento': 'Dados monitoria',
    'Canal de Atendimento': 'Dados monitoria',
    'Título': 'Dados monitoria',
    'Condução - Seguiu script de abertura?': 'Gestão de ação',
    'Condução - Se apresentou de forma educada e profissional?': 'Gestão de ação',
    'Condução - Manteve comunicação assertiva, correta e formal?': 'Gestão de ação',
    'Condução - Personalizou o atendimento?': 'Gestão de ação',
    'Condução - Conduziu o atendimento sem abandono?': 'Gestão de ação',
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 'Gestão de ação',
    'Condução - Seguiu script de encerramento?': 'Gestão de ação',
    'Consulta Agendamento - Atendimento citou processo de agendamento de consulta?': 'Ínicio de bloco',
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 'Política empresa/convênio',
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 'Gestão de ação',
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 'Gestão de ação',
    'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?': 'Ínicio de bloco',
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Política empresa/convênio',
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Gestão de ação',
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 'Política empresa/convênio',
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Política empresa/convênio',
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 'Gestão de ação',
    'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?': 'Ínicio de bloco',
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Política empresa/convênio',
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Gestão de ação',
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 'Política empresa/convênio',
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Política empresa/convênio',
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 'Gestão de ação',
    'Consulta Cancelamento - Atendimento citou processo de cancelamento de consulta?': 'Ínicio de bloco',
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Consulta Cancelamento - Ofertou nova data?': 'Gestão de ação',
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 'Gestão de ação',
    'Exame Agendamento - Atendimento citou processo de agendamento de exame?': 'Ínicio de bloco',
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 'Política empresa/convênio',
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 'Gestão de ação',
    'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 'Política empresa/convênio',
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 'Gestão de ação',
    'Exame Agendamento - Ofereceu opções para o agendamento?': 'Gestão de ação',
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 'Gestão de ação',
    'Exame Agendamento - Orientou sobre o preparo de exame?': 'Gestão de ação',
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 'Gestão de ação',
    'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?': 'Ínicio de bloco',
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 'Gestão de ação',
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 'Gestão de ação',
    'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Política empresa/convênio',
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 'Gestão de ação',
  'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 'Política empresa/convênio',
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 'Gestão de ação',
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 'Gestão de ação',
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 'Gestão de ação',
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 'Gestão de ação',
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 'Gestão de ação',
    'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?': 'Ínicio de bloco',
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Política empresa/convênio',
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 'Gestão de ação',
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 'Política empresa/convênio',
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 'Gestão de ação',
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 'Gestão de ação',
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 'Gestão de ação',
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 'Gestão de ação',
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 'Gestão de ação',
    'Exame Cancelamento - Atendimento citou processo de cancelamento de exame?': 'Ínicio de bloco',
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Gestão de ação',
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Gestão de ação',
    'Exame Cancelamento - Atendente ofertou nova data?': 'Gestão de ação',
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 'Gestão de ação',
    'Tabulação do atendimento:': 'Gestão de ação',
    'Visão Avaliador - Nota Csat do atendimento (0 a 5):': 'Item CX visão avaliador',
    'Visão Cliente - Atendimento foi bom?': 'Item CX visão avaliador',
    'Visão Processos - Processos atenderam a necessidade do cliente?': 'Item CX visão avaliador',
    'Observações sobre processos não concluídos:': 'Item CX visão avaliador',
    'Classificação (0-5)': 'Item CX visão avaliador',
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 'Gestão de ação',
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 'Gestão de ação',
    'Condução - Acordo de Pausa': 'Gestão de ação',
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 'Gestão de ação',
    'Tipo de Item': 'Dados monitoria',
    'Caminho': 'Dados monitoria',
    'Unidades citadas no atendimento': 'Dados monitoria',
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 'Gestão de ação'
}

pontuacao_dict = {
    'Condução - Seguiu script de abertura?': 1,
    'Condução - Se apresentou de forma educada e profissional?': 1,
    'Condução - Manteve comunicação assertiva, correta e formal?': 10,
    'Condução - Personalizou o atendimento?': 1,
    'Condução - Conduziu o atendimento sem abandono?': 0.5,
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 10,
    'Condução - Seguiu script de encerramento?': 1,
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 9,
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 0.5,
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 5,
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 7,
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 4,
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 9,
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 7,
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 0.5,
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 7,
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 3,
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 9,
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 7,
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 0.5,
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Consulta Cancelamento - Ofertou nova data?': 0.5,
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 0.5,
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 7,
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 0.5,
   'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 8,
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 0.5,
    'Exame Agendamento - Ofereceu opções para o agendamento?': 3,
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 0.5,
    'Exame Agendamento - Orientou sobre o preparo de exame?': 0.5,
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 0.5,
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 0.5,
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 0.5,
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 7,
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 3,
    'Exame Reagendamento Pré - Paciente concordou com a data e horario agendado?': 8,
    'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 8,
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 0.5,
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 3,
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 0.5,
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 0.5,
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 0.5,
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 7,
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 2,
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 8,
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 0.5,
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 0.5,
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 0.5,
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 3,
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 0.5,
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 0.5,
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 0.5,
    'Exame Cancelamento - Atendente ofertou nova data?': 0.5,
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 3,
    'Tabulação do atendimento:': 5,
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 10,
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 0.5,
    'Condução - Acordo de Pausa': 0.5,
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 5,
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 5
}

status_ncg_dict = {
    'Itens do check list': 'Status NCG',
    'Condução - Seguiu script de abertura?': 'NÃO CONFORMIDADE',
    'Condução - Se apresentou de forma educada e profissional?': 'NÃO CONFORMIDADE',
    'Condução - Manteve comunicação assertiva, correta e formal?': 'NÃO CONFORMIDADE',
    'Condução - Personalizou o atendimento?': 'NÃO CONFORMIDADE',
    'Condução - Conduziu o atendimento sem abandono?': 'NÃO CONFORMIDADE GRAVE',
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 'NÃO CONFORMIDADE',
    'Condução - Seguiu script de encerramento?': 'NÃO CONFORMIDADE',
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 'NÃO CONFORMIDADE',
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 'NÃO CONFORMIDADE',
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 'NÃO CONFORMIDADE',
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Cancelamento - Ofertou nova data?': 'NÃO CONFORMIDADE GRAVE',
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 'NÃO CONFORMIDADE GRAVE',
   'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 'NÃO CONFORMIDADE',
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Ofereceu opções para o agendamento?': 'NÃO CONFORMIDADE',
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Orientou sobre o preparo de exame?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 'NÃO CONFORMIDADE',
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Cancelamento - Atendente ofertou nova data?': 'NÃO CONFORMIDADE GRAVE',
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 'NÃO CONFORMIDADE',
    'Tabulação do atendimento:': 'NÃO CONFORMIDADE',
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 'NÃO CONFORMIDADE',
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 'NÃO CONFORMIDADE GRAVE',
    'Condução - Acordo de Pausa': 'NÃO CONFORMIDADE',
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 'NÃO CONFORMIDADE',
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 'NÃO CONFORMIDADE GRAVE'
}

bloco_dict = {
    'Itens do check list': 'Bloco',
    'Condução - Seguiu script de abertura?': 'Condução',
    'Condução - Se apresentou de forma educada e profissional?': 'Condução',
    'Condução - Manteve comunicação assertiva, correta e formal?': 'Condução',
    'Condução - Personalizou o atendimento?': 'Condução',
    'Condução - Conduziu o atendimento sem abandono?': 'Condução',
    'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?': 'Condução',
    'Condução - Seguiu script de encerramento?': 'Condução',
    'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Agendamento',
    'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Agendamento',
    'Consulta Agendamento - Paciente concordou com a data e horário ofertado?': 'Consulta Agendamento',
    'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?': 'Consulta Agendamento',
    'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?': 'Consulta Agendamento',
    'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Confirmação',
    'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Confirmação',
    'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?': 'Consulta Reagendamento Pós',
    'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?': 'Consulta Reagendamento Pré',
    'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?': 'Consulta Reagendamento Pré',
    'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Cancelamento',
    'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Cancelamento',
    'Consulta Cancelamento - Ofertou nova data?': 'Consulta Cancelamento',
    'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?': 'Consulta Cancelamento',
    'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Exame agendamento',
    'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Exame agendamento',
    'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?': 'Exame agendamento',
    'Exame Agendamento - Verbalizou se o pedido médico tem data?': 'Exame agendamento',
    'Exame Agendamento - Paciente concordou com a data  e horario agendado?': 'Exame agendamento',
    'Exame Agendamento - Respeitou as regras de conflito de exames?': 'Exame agendamento',
    'Exame Agendamento - Ofereceu opções para o agendamento?': 'Exame agendamento',
    'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?': 'Exame agendamento',
    'Exame Agendamento - Orientou sobre o preparo de exame?': 'Exame agendamento',
    'Exame Agendamento - Orientou sobre envio de documentações para autorização?': 'Exame Agendamento',
    'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?': 'Exame Confirmação',
    'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Exame confirmação',
    'Exame Confirmação - Encerrou com informações do agendamendo exame?': 'Exame confirmação',
    'Exame Confirmação - Orientou sobre envio de documentações para autorização?': 'Exame Confirmação',
    'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?': 'Exame Reagendamento Pré',
    'Exame Reagendamento Pré - Ofereceu opções para o agendamento?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Orientou sobre o preparo de exame?': 'Consulta Reagendamento Pré',
    'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?': 'Exame Reagendamento Pré',
    'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?': 'Exame Reagendamento Pós',
    'Exame Reagendamento Pós - Ofereceu opções para o agendamento?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Orientou sobre o preparo de exame?': 'Consulta Reagendamento Pós',
    'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?': 'Exame Reagendamento Pós',
    'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?': 'Exame cancelamento',
    'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?': 'Exame cancelamento',
    'Exame Cancelamento - Atendente ofertou nova data?': 'Exame cancelamento',
    'Exame Cancelamento - Confirmou as informações do cancelamento de exame?': 'Exame cancelamento',
    'Tabulação do atendimento:': 'Processos gerais',
    'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?': 'Processos gerais',
    'Consulta Agendamento - Informou o valor correto para o agendamento realizado': 'Consulta Agendamento',
    'Condução - Acordo de Pausa': 'Condução',
    'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas': 'Processos gerais',
    'Processos gerais - O agendamento foi registrado corretamente no sistema?': 'Processos gerais'
}

class Bases():
    def __init__(self,dados=None,ano=2026, base_particular=None,base_bd_novo=None,base_bd_empilhado=None,grupo_df=None,dime_itens=None,monitoria=None,dime_feedback=None,dime_gerente=None,dime_operador=None,dime_coordenador=None,dime_supervisor=None,dime_ncg_nc=None,dime_conformidade=None,feedback=None,base_tratada=None,base_bd=None,base_tratada1=None,base_final=None,base_meta=None):
        self.base_monitoria = monitoria
        self.feedback = feedback
        self.base_bd = base_bd
        self.base_tratada = base_tratada
        self.base_tratada1 = base_tratada1
        self.base_final = base_final
        self.base_meta = base_meta
        self.dime_ncg_nc = dime_ncg_nc
        self.dime_conformidade = dime_conformidade
        self.dime_supervisor = dime_supervisor
        self.dime_coordenador = dime_coordenador
        self.dime_operador = dime_operador
        self.dime_gerente = dime_gerente
        self.base_bd_novo = base_bd_novo
        self.dime_feedback = dime_feedback
        self.dime_itens = dime_itens
        self.dados = dados
        self.ano = ano
        self.base_bd_empilhado = base_bd_empilhado
        self.grupo_df = grupo_df
        self.base_particular = base_particular
         
    def carregar(self):
        self.base_monitoria = pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\Novo_paper.xlsx")
        self.base_paper_amo = pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\paper_amo.xlsx")
        self.base_particular = pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\particular.xlsx")
        self.feedback=pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\feedback.xlsx")
        self.base_bd=pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\BD_novo_dia_a_dia.xlsx")
        self.base_bd_novo =pd.read_excel(r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\data.xlsx")
        
        self.base_monitoria['Chave de Identificação do Atendimento:'] = self.base_monitoria['Chave de Identificação do Atendimento:'].astype(str).str.replace(r'\s+', '', regex=True)
        self.base_particular['Chave de Identificação do Atendimento:'] = self.base_particular['Chave de Identificação do Atendimento:'].astype(str).str.replace(r'\s+', '', regex=True)
        self.feedback['Chave de Identificação do Atendimento:'] = self.feedback['Chave de Identificação do Atendimento:'].astype(str).str.replace(r'\s+', '', regex=True)
        self.base_paper_amo['Chave de Identificação do Atendimento:'] = self.base_paper_amo['Chave de Identificação do Atendimento:'].astype(str).str.replace(r'\s+', '', regex=True)
       
    def calendario(self):
        data_inicio = datetime(self.ano, 1, 1)
        data_fim = datetime(self.ano, 12, 31)
    
        todas_datas = pd.date_range(start=data_inicio, end=data_fim, freq='D')
    
        self.dados = []
        for data in todas_datas:
            self.dados.append({
                        'data': data.strftime('%d/%m/%Y'),  # Formato dd/mm/aaaa
                        'mes': data.month,  # Mês como número (1-12)
                        'dia': data.day,
                        'ano': data.year
                    })
        calendario_dict = {
                        1: 'janeiro de 2026',
                        2: 'fevereiro de 2026',
                        3: 'março de 2026',
                        4: 'abril de 2026',
                        5: 'maio de 2026',
                        6: 'junho de 2026',
                        7: 'julho de 2026',
                        8: 'agosto de 2026',
                        9: 'setembro de 2026',
                        10: 'outubro de 2026',
                        11: 'novembro de 2026',
                        12: 'dezembro de 2026'
                    }
        self.dados = pd.DataFrame(self.dados)
        self.dados['ANO/MES']=self.dados['mes'].map(calendario_dict)

    def concatenar(self):
        # CORREÇÃO PARA O FUTURE WARNING
        # Filtrar as bases e remover colunas 100% vazias antes do concat para evitar comportamento deprecado
        bases = [self.base_monitoria, self.base_particular, self.base_paper_amo]
        bases_limpas = []
        for df in bases:
            if df is not None and not df.empty:
                bases_limpas.append(df.dropna(axis=1, how='all'))
        
        if bases_limpas:
            self.base_monitoria = pd.concat(bases_limpas, ignore_index=True)

    def tratamento1(self):
        self.base_tratada=pd.melt(self.base_monitoria,
                     id_vars=['Nome do Avaliador:',
                                'Data da Monitoria:',
                                'Data do Atendimento:',
                                'Matrícula do Consultor Monitorado:',
                                'Chave de Identificação do Atendimento:',
                                'Unidades citadas no atendimento',
                                'Origem do Atendimento:',
                                'Canal de Atendimento:',
                                'Consulta Agendamento - Atendimento citou processo de agendamento  de consulta?',
                                'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?',
                                'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?',
                                'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?',
                                'Consulta Cancelamento - Atendimento citou processo de cancelamento  de consulta?',
                                'Exame Agendamento - Atendimento citou processo de agendamento de exame?',
                                'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?',
                                'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?',
                                'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?',
                                'Exame Cancelamento - Atendimento citou processo de cancelamento  de exame?',
                                'Visão Cliente - Atendimento foi bom?',
                                'Visão Processos - Processos atenderam a necessidade do cliente?',
                                'Visão Avaliador - Nota Csat do atendimento (0 a 5):','Modificado',
                                'Criado',
                                'Criado por',
                                'Modificado por',
                                'Link do Item',
                                'Tipo de Item',
                                'Caminho',
                                'Observações sobre processos não concluídos:'],
                    value_vars=['Condução - Seguiu script de abertura?',
                                'Condução - Se apresentou de forma educada e profissional?',
                                'Condução - Manteve comunicação assertiva, correta e formal?',
                                'Condução - Personalizou o atendimento?',
                                'Condução - Conduziu o atendimento sem abandono?',
                                'Condução - Evitou aspectos negativos aos serviços, marca ou unidade?',
                                'Condução - Seguiu script de encerramento?',
                                'Consulta Agendamento - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Consulta Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Consulta Agendamento - Paciente concordou com a data e horário ofertado?',
                                'Consulta Agendamento - Incluiu justificativa de pacientes oncológicos?',
                                'Consulta Agendamento - Ofereceu opções de navegação para o agendamento?',
                                'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?',
                                'Consulta Confirmação - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Consulta Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?',
                                'Consulta Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Consulta Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Consulta Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?',
                                'Consulta Reagendamento Pré - Ofereceu opções de data e horário para o reagendamento ou lista de espera?',
                                'Consulta Reagendamento Pré - Paciente concordou com a data e horário reagendado?',
                                'Consulta Reagendamento Pré - O agendamento foi impactado por conflito com exames ou outras consultas?',
                                'Consulta Reagendamento Pré - Incluiu justificativa de pacientes oncológicos?',
                                'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?',
                                'Consulta Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Consulta Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Consulta Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?',
                                'Consulta Reagendamento Pós - Ofereceu opções de data e horário para o reagendamento ou lista de espera?',
                                'Consulta Reagendamento Pós - Paciente concordou com a data e horario reagendado?',
                                'Consulta Reagendamento Pós - O agendamento foi impactado por conflito com exames ou outras consultas?',
                                'Consulta Reagendamento Pós - Incluiu justificativa de pacientes oncológicos?',
                                'Consulta Cancelamento - Atendimento citou processo de cancelamento  de consulta?',
                                'Consulta Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Consulta Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Consulta Cancelamento - Ofertou nova data?',
                                'Consulta Cancelamento - Encerrou com informações do cancelamento da consulta?',
                                'Exame Agendamento - Atendimento citou processo de agendamento de exame?',
                                'Exame Agendamento - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Exame Agendamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Exame Agendamento - Cobertura do convênio impactou negativamente no agendamento?',
                                'Exame Agendamento - Verbalizou se o pedido médico tem data?',
                                'Exame Agendamento - Paciente concordou com a data  e horario agendado?',
                                'Exame Agendamento - Respeitou as regras de conflito de exames?',
                                'Exame Agendamento - Ofereceu opções para o agendamento?',
                                'Exame Agendamento - Fez checklist de perguntas obrigatórias com o paciente?',
                                'Exame Agendamento - Orientou sobre o preparo de exame?',
                                'Exame Agendamento - Orientou sobre envio de documentações para autorização?',
                                'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?',
                                'Exame Confirmação - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Exame Confirmação - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Exame Confirmação - Encerrou com informações do agendamendo exame?',
                                'Exame Confirmação - Orientou sobre envio de documentações para autorização?',
                                'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?',
                                'Exame Reagendamento Pré - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Exame Reagendamento Pré - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Exame Reagendamento Pré - Cobertura do convênio impactou negativamente no agendamento?',
                                'Exame Reagendamento Pré - Verbalizou se o pedido médico tem data?',
                                'Exame Reagendamento Pré - Paciente concordou com a data  e horario agendado?',
                                'Exame Reagendamento Pré - Respeitou as regras de conflito de exames?',
                                'Exame Reagendamento Pré - Ofereceu opções para o agendamento?',
                                'Exame Reagendamento Pré - Fez checklist de perguntas obrigatórias com o paciente?',
                                'Exame Reagendamento Pré - Orientou sobre o preparo de exame?',
                                'Exame Reagendamento Pré - Orientou sobre envio de documentações para autorização?',
                                'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?',
                                'Exame Reagendamento Pós - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Exame Reagendamento Pós - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Exame Reagendamento Pós - Cobertura do convênio impactou negativamente no agendamento?',
                                'Exame Reagendamento Pós - Verbalizou se o pedido médico tem data?',
                                'Exame Reagendamento Pós - Paciente concordou com a data e horario agendado?',
                                'Exame Reagendamento Pós - Respeitou as regras de conflito de exames?',
                                'Exame Reagendamento Pós - Ofereceu opções para o agendamento?',
                                'Exame Reagendamento Pós - Fez checklist de perguntas obrigatórias com o paciente?',
                                'Exame Reagendamento Pós - Orientou sobre o preparo de exame?',
                                'Exame Reagendamento Pós - Orientou sobre envio de documentações para autorização?',
                                'Exame Cancelamento - Atendimento citou processo de cancelamento  de exame?',
                                'Exame Cancelamento - Confirmou dados obrigatórios do paciente conforme processo?',
                                'Exame Cancelamento - Confirmou dados do paciente antes de informar dados sigilosos (LGPD)?',
                                'Exame Cancelamento - Atendente ofertou nova data?',
                                'Exame Cancelamento - Confirmou as informações do cancelamento de exame?',
                                'Tabulação do atendimento:',
                                'Processos gerais - Registrou informações na observação, pertinentes ao atendimento, corretamente?',
                                'Processos gerais - O agendamento foi registrado corretamente no sistema?',
                                'Consulta Agendamento - Informou o valor correto para o agendamento realizado',
                                'Processos gerais - Respondeu com clareza e corretamente as perguntas do paciente no caso de contato apenas para tirar duvidas',
                                'Condução - Acordo de Pausa'],
                    var_name="Item checklist",
                    value_name="Conformidade")
       
    def tratamento2(self):
        self.base_tratada['Data monitoria calculo feedback']=self.base_tratada['Data da Monitoria:']
        self.base_tratada=self.base_tratada[['Nome do Avaliador:', 
                                             'Data da Monitoria:',
                                             'Data do Atendimento:',
                                            'Matrícula do Consultor Monitorado:',
                                            'Chave de Identificação do Atendimento:',
                                            'Unidades citadas no atendimento', 
                                            'Origem do Atendimento:',
                                            'Canal de Atendimento:',
                                            'Consulta Agendamento - Atendimento citou processo de agendamento  de consulta?',
                                            'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?',
                                            'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?',
                                            'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?',
                                            'Consulta Cancelamento - Atendimento citou processo de cancelamento  de consulta?',
                                            'Exame Agendamento - Atendimento citou processo de agendamento de exame?',
                                            'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?',
                                            'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?',
                                            'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?',
                                            'Exame Cancelamento - Atendimento citou processo de cancelamento  de exame?',
                                            'Visão Cliente - Atendimento foi bom?',
                                            'Visão Processos - Processos atenderam a necessidade do cliente?',
                                            'Observações sobre processos não concluídos:',
                                            'Criado por',
                                            'Tipo de Item',
                                            'Caminho',
                                            "Item checklist",
                                            "Conformidade"]]
        
    def tratamento3(self):
        self.base_tratada['Conformidade']=np.where(self.base_tratada['Conformidade']=="",self.base_tratada['Conformidade'].fillna('Sem preencher',inplace=True),self.base_tratada['Conformidade'])
        self.base_tratada['Status conformidade']=np.where((self.base_tratada['Conformidade']=='C - Conforme')|(self.base_tratada['Conformidade']=='NA - Não Se Aplica')|(self.base_tratada['Conformidade']=='Sem preencher'),'C - Conforme','NC - Não Conforme')
        self.base_tratada['Pontuação recebida']=np.where(self.base_tratada['Conformidade']=='NC - Não Conforme',0,self.base_tratada['Item checklist'].map(pontuacao_dict))
        self.base_tratada['NCG']=self.base_tratada['Item checklist'].map(status_ncg_dict)
        self.base_tratada['Bloco']=self.base_tratada['Item checklist'].map(bloco_dict)
        self.base_tratada['Classificação']=self.base_tratada['Item checklist'].map(classificão_dict)
        self.base_tratada['Responsabilidade']=self.base_tratada['Item checklist'].map(responsabilidade_dict)

        base_zerada = self.base_tratada.query('`Status conformidade` == "NC - Não Conforme" and `NCG` == "NÃO CONFORMIDADE GRAVE"').copy()
        base_zerada['Nota zerada']=0
        base_zerada_nota=base_zerada[['Chave de Identificação do Atendimento:','Nota zerada']]
        nota_zerada=base_zerada_nota.set_index('Chave de Identificação do Atendimento:')['Nota zerada'].to_dict()
        base_zerada_responbilidade=base_zerada[['Chave de Identificação do Atendimento:','Responsabilidade']]
        base_zerada_responbilidade = base_zerada.rename(columns={'Responsabilidade': 'Responsabilidade_zerada'}).copy()
        responsabilidade_zerada=base_zerada_responbilidade.set_index('Chave de Identificação do Atendimento:')['Responsabilidade_zerada'].to_dict()

        self.base_tratada['Nota zerada']=self.base_tratada['Chave de Identificação do Atendimento:'].map(nota_zerada)
        self.base_tratada['Responsabilidade_zerada']=self.base_tratada['Chave de Identificação do Atendimento:'].map(responsabilidade_zerada)
        self.base_tratada['Nota zerada']=pd.to_numeric(self.base_tratada['Nota zerada'], errors='coerce')
        self.base_tratada['Nota atendimento']=np.where(self.base_tratada['Nota zerada']==0,0,self.base_tratada.groupby('Chave de Identificação do Atendimento:')['Pontuação recebida'].transform('sum')/2)

        base_humano=self.base_tratada.query("`Responsabilidade`== 'Humano'").copy()
        base_humano['Nota humano']=base_humano.groupby('Chave de Identificação do Atendimento:')['Pontuação recebida'].transform('sum')
        base_humano=base_humano[['Chave de Identificação do Atendimento:','Nota humano']]
        nota_humano=base_humano.set_index('Chave de Identificação do Atendimento:')['Nota humano'].to_dict()

        self.base_tratada['Nota humano']=np.where(self.base_tratada['Responsabilidade_zerada']=='Humano',0,self.base_tratada['Chave de Identificação do Atendimento:'].map(nota_humano))

        base_regra_negocio=self.base_tratada.query("`Responsabilidade`== 'Regra de negocio'").copy()
        base_regra_negocio['Nota Regra negocio']=base_regra_negocio.groupby('Chave de Identificação do Atendimento:')['Pontuação recebida'].transform('sum')
        base_regra_negocio=base_regra_negocio[['Chave de Identificação do Atendimento:','Nota Regra negocio']]
        nota_regra_negocio=base_regra_negocio.set_index('Chave de Identificação do Atendimento:')['Nota Regra negocio'].to_dict()
        self.base_tratada['Nota Regra negocio']=np.where(self.base_tratada['Responsabilidade_zerada']=='Regra de negocio',0,self.base_tratada['Chave de Identificação do Atendimento:'].map(nota_regra_negocio))

        monitores=['Mauricio Novaes Santos','Denise de Oliveira Basso','Suzane Sales Galveia','Aline Dos Santos Malaquias','Aline dos Santos Malaquias','Enzo Samid Marinho de Carvalho']

        def Avaliador(valor):
            if valor in monitores:
                return 'Monitor'
            else:
                return 'Supervisor'
        self.base_tratada['Tipo de avaliador']=self.base_tratada['Nome do Avaliador:'].apply(Avaliador)
        self.base_tratada['Mes ref monitoria']=self.base_tratada['Data da Monitoria:'].dt.month
        self.base_tratada['Ano ref monitoria']=self.base_tratada['Data da Monitoria:'].dt.year

        meses = {
            1: 'Janeiro',
            2: 'Fevereiro',
            3: 'Março',
            4: 'Abril',
            5: 'Maio',
            6: 'Junho',
            7: 'Julho',
            8: 'Agosto',
            9: 'Setembro',
            10: 'Outubro',
            11: 'Novembro',
            12: 'Dezembro'
        }

        self.base_tratada['Mes ref monitoria']=self.base_tratada['Mes ref monitoria'].map(meses)
        
        self.base_tratada['Matrícula do Consultor Monitorado:']=pd.to_numeric( self.base_tratada['Matrícula do Consultor Monitorado:'],errors='coerce')
        
        meses = {
                    1: 'Janeiro',
                    2: 'Fevereiro',
                    3: 'Março',
                    4: 'Abril',
                    5: 'Maio',
                    6: 'Junho',
                    7: 'Julho',
                    8: 'Agosto',
                    9: 'Setembro',
                    10: 'Outubro',
                    11: 'Novembro',
                    12: 'Dezembro'
                }
        self.base_tratada['Dado ref monitoria']=self.base_tratada['Data da Monitoria:'].dt.month
        self.base_tratada['Mes ref monitoria']=self.base_tratada['Dado ref monitoria'].map(meses)
        self.base_tratada['GESTOR_I_chave_monitoria']= self.base_tratada['Nome do Avaliador:'].str.strip().str.upper()
        self.base_tratada['GESTOR_I_chave_monitoria']= self.base_tratada['Nome do Avaliador:'].str.upper().str.replace(' ','')
        self.base_tratada['GESTOR_I_chave_monitoria']= self.base_tratada['Nome do Avaliador:'].str.upper().str.replace(r'\s+', '', regex=True)
        
        # Correção da Geração de Chave (Padronizando para Chave Mensal para compatibilidade com base_bd_novo)
        self.base_tratada['Data_TEMP'] = pd.to_datetime(self.base_tratada['Data da Monitoria:'], errors='coerce')
        self.base_tratada['Matrícula STR'] = self.base_tratada['Matrícula do Consultor Monitorado:'].astype(str).str.replace(r'\.0$', '', regex=True)
        self.base_tratada['Chave_bd_supervisor'] = self.base_tratada['Data_TEMP'].dt.strftime('%Y-%m') + self.base_tratada['Matrícula STR']
        
        self.base_tratada['Data da Monitoria:'] = self.base_tratada['Data_TEMP'].dt.strftime('%Y-%m-%d')
        
        self.base_tratada=self.base_tratada[['Nome do Avaliador:',
'Data da Monitoria:',
'Data do Atendimento:',
'Matrícula do Consultor Monitorado:',
'Chave de Identificação do Atendimento:',
'Unidades citadas no atendimento',
'Origem do Atendimento:',
'Canal de Atendimento:',
'Item checklist',
'Status conformidade',
'NCG',
'Bloco',
'Classificação',
'Responsabilidade',
'Nota atendimento',
'Nota humano',
'Nota Regra negocio',
'Tipo de avaliador',
'Mes ref monitoria',
'Ano ref monitoria',
'Dado ref monitoria',
'Chave_bd_supervisor',
'Observações sobre processos não concluídos:',
'Criado por',
'Tipo de Item',
'Caminho']]

    def tratamento4(self):
        self.base_tratada1=self.base_monitoria[['Nome do Avaliador:',
'Data da Monitoria:',
'Data do Atendimento:',
'Matrícula do Consultor Monitorado:',
'Chave de Identificação do Atendimento:',
'Unidades citadas no atendimento',
'Origem do Atendimento:',
'Canal de Atendimento:',
'Consulta Agendamento - Atendimento citou processo de agendamento  de consulta?',
'Consulta Confirmação - Atendimento citou processo de confirmação de consulta?',
'Consulta Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré data de atendimento?',
'Consulta Reagendamento Pós - Atendimento teve processo de reagendamento de exame pós data de atendimento?',
'Consulta Cancelamento - Atendimento citou processo de cancelamento  de consulta?',
'Exame Agendamento - Atendimento citou processo de agendamento de exame?',
'Exame Confirmação - Atendimento citou processo de confirmação agendamento de exame?',
'Exame Reagendamento Pré - Atendimento citou processo de reagendamento de exame pré atendimento?',
'Exame Reagendamento Pós - Atendimento citou processo de reagendamento de exame pós atendimento?',
'Exame Cancelamento - Atendimento citou processo de cancelamento  de exame?',
'Visão Cliente - Atendimento foi bom?',
'Visão Processos - Processos atenderam a necessidade do cliente?']]
                  
    def tratamento5(self):
        self.base_bd = self.base_bd.rename(columns={'GESTOR 1': 'GESTOR I', 'GESTOR 2 ': 'GESTOR II','GESTOR 3': 'GESTOR III'})
        textos_excluir = [
                            'Filtros aplicados:',
                            'Parâmetro é Segmento',
                            'DATA é igual a ou está depois de 01/08/2025']
        
        self.base_bd = self.base_bd.loc[~self.base_bd['DATA'].isin(textos_excluir)]
        self.base_bd['DATA'] = pd.to_datetime(self.base_bd['DATA'], errors='coerce')
        
        self.base_bd['Mes ref bd']=self.base_bd['DATA'].dt.month
        self.base_bd['Mes ref ano']=self.base_bd['DATA'].dt.year
        
        # Correção da Geração de Chave BD Antiga (Mensal para unificar com BD Nova)
        self.base_bd['Matrícula STR'] = self.base_bd['MATRICULA'].astype(str).str.replace(r'\.0$', '', regex=True)
        self.base_bd['Chave_bd_supervisor'] = self.base_bd['DATA'].dt.strftime('%Y-%m') + self.base_bd['Matrícula STR']
        
        meses = {
            1: 'Janeiro',
            2: 'Fevereiro',
            3: 'Março',
            4: 'Abril',
            5: 'Maio',
            6: 'Junho',
            7: 'Julho',
            8: 'Agosto',
            9: 'Setembro',
            10: 'Outubro',
            11: 'Novembro',
            12: 'Dezembro'
        }

        self.base_bd['Mes ref bd']=self.base_bd['Mes ref bd'].map(meses)

        self.base_bd["Dia semana"]=self.base_bd['DATA'].dt.day_name()

        dias_semana = {
            'Monday': 'Segunda-feira',
            'Tuesday': 'Terça-feira',
            'Wednesday': 'Quarta-feira',
            'Thursday': 'Quinta-feira',
            'Friday': 'Sexta-feira',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }

        self.base_bd["Dia semana"]=self.base_bd["Dia semana"].map(dias_semana)

        dias_contar = {
        'Segunda-feira':1,
            'Terça-feira':'oi',
        'Quarta-feira':1,
        'Quinta-feira':1,
            'Sexta-feira':1,
        'Sábado':0,
        'Domingo':0
        }

        self.base_bd['contar']=0

        def Contar(valor):
            if valor == 'Sábado':
                return 0
            elif valor == 'Domingo':
                return 0
            else:
                return 1

        self.base_bd['contar']=self.base_bd["Dia semana"].apply(Contar)
        self.base_bd['Dados meta monitoria']=self.base_bd.groupby(['Mes ref ano','Mes ref bd','GESTOR I','NOME'])['contar'].transform('sum')
        
        self.base_bd['Nº MES'] =self.base_bd['DATA'].dt.month
        self.base_bd['DATA']=self.base_bd['DATA'].dt.strftime('%Y-%m-%d') 
        meses_bd_novo = {
                        12025: "janeiro de 2025",
                        22025: "fevereiro de 2025",
                        32025: "março de 2025",
                        42025: "abril de 2025",
                        52025: "maio de 2025",
                        62025: "junho de 2025",
                        72025: "julho de 2025",
                        82025: "agosto de 2025",
                        92025: "setembro de 2025",
                        102025: "outubro de 2025",
                        112025: "novembro de 2025",
                        122025: "dezembro de 2025",
                        12026: "janeiro de 2026",
                        22026: "fevereiro de 2026",
                        32026: "março de 2026",
                        42026: "abril de 2026",
                        52026: "maio de 2026",
                        62026: "junho de 2026",
                        72026: "julho de 2026",
                        82026: "agosto de 2026",
                        92026: "setembro de 2026",
                        102026: "outubro de 2026",
                        112026: "novembro de 2026",
                        122026: "dezembro de 2026"
        }
        
        self.base_bd['Chave_mes']=self.base_bd['Nº MES'].astype('str')+''+self.base_bd['Mes ref ano'].astype('str')
        self.base_bd['Chave_mes'] = pd.to_numeric(self.base_bd['Chave_mes'], errors='coerce').fillna(0).astype(int)
        self.base_bd['ANO/MES']=self.base_bd['Chave_mes'].map(meses_bd_novo)
        self.base_bd=self.base_bd[['DATA', 'MATRICULA', 'NOME', 'CARGO', 'DESCRIÇÃO DO CARGO', 'GESTOR I',
       'GESTOR II', 'GESTOR III', 'NOME TWILIO', 'NOME INGENIUM', 'TELEFONIA', 'INÍCIO INSS',
       'SEGMENTAÇÃO VERTICAL', 'SIGLA', 'REGIONAL', 'TEMPO DE CASA', 'ID SCORB ',
       'LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISC', 'LOGIN HIGH CONNECT',
       'LOGIN INGENIUM', 'LOGIN SMARTNX', 'LOGIN INSTANT', 'LOGIN MEDCLINIC',
       'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN SMART', 'LOGIN TASY',
       'LOGIN TWILIO', 'LOGIN VOIP', 'NOME TASY ', 'Mes ref bd', 'Mes ref ano',
       'Chave_bd_supervisor', 'Dia semana', 'contar', 'Dados meta monitoria',
       'Nº MES', 'Chave_mes', 'ANO/MES']]
    
    def tratamento6(self):
        self.base_bd['LOGIN DYNAMO']=0
        self.base_bd['LOGIN INSTANT SOLUTION']=0
        self.base_bd['LOGIN MV H9J']=0
        self.base_bd['LOGIN SISAGENDA']=0
        self.base_bd['Meta supervisor']=0
        self.base_bd=self.base_bd.rename(columns={'GESTOR I':'SUPERVISOR','GESTOR II':'COORDENADOR','GESTOR III':'GERENTE','LOGIN CISC':'LOGIN CISCO','LOGIN HIGH CONNECT':'LOGIN HIGH CONECT'})
        self.base_bd=self.base_bd[['ANO/MES','DATA','Chave_bd_supervisor', 'GERENTE', 'COORDENADOR', 'SUPERVISOR','MATRICULA','NOME','SIGLA','REGIONAL','LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISCO', 'LOGIN DYNAMO', 'LOGIN HIGH CONECT', 'LOGIN INGENIUM', 'LOGIN INSTANT SOLUTION', 'LOGIN MEDCLINIC', 'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN MV H9J', 'LOGIN SISAGENDA', 'LOGIN SMART', 'LOGIN SMARTNX', 'LOGIN TASY', 'LOGIN TWILIO', 'LOGIN VOIP', 'Meta supervisor']]
        
        def meta_sup(valor):
            if valor >=20:
                return 4
            elif valor >=15:
                return 3
            elif valor >= 10:
                return 2
            elif valor >= 5:
                return 1
            else:
                return 0

        self.base_bd['Contar']=self.base_bd.groupby(['SUPERVISOR','ANO/MES'])['NOME'].transform('count')

        dicionario=self.base_bd[['ANO/MES','DATA']].drop_duplicates('DATA')
        dicionario=pd.DataFrame(dicionario)
        dicionario['Nº dias']=dicionario.groupby('ANO/MES')['DATA'].transform('count')
   
        
        self.base_bd=pd.merge(self.base_bd,
                              dicionario,
                              on='ANO/MES',
                              how='left')
        
        self.base_bd['Nº operador']=self.base_bd['Contar']/self.base_bd['Nº dias']
        self.base_bd['Meta supervisor']=self.base_bd['Nº operador']*4
        
        # Correção da exclusão: Agora ele exclui a chave mensal, preservando um registro do operador pro mês inteiro sem apagar históricos antigos
        self.base_bd=self.base_bd.drop_duplicates('Chave_bd_supervisor')
        
        self.base_bd=self.base_bd[['ANO/MES','Chave_bd_supervisor', 'GERENTE', 'COORDENADOR', 'SUPERVISOR','MATRICULA','NOME','SIGLA','REGIONAL','LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISCO', 'LOGIN DYNAMO', 'LOGIN HIGH CONECT', 'LOGIN INGENIUM', 'LOGIN INSTANT SOLUTION', 'LOGIN MEDCLINIC', 'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN MV H9J', 'LOGIN SISAGENDA', 'LOGIN SMART', 'LOGIN SMARTNX', 'LOGIN TASY', 'LOGIN TWILIO', 'LOGIN VOIP', 'Meta supervisor']]
        
    def tratamento7(self):
            self.base_tratada['Data monitoria calculo feedback']=self.base_tratada['Data da Monitoria:']
            self.base_tratada['Data monitoria calculo feedback']=pd.to_datetime(self.base_tratada['Data monitoria calculo feedback'])
            
            # CORREÇÃO PARA O [Errno 28] No space left on device: Remoção de chaves duplicadas no feedback
            feedback_limpo = self.feedback.drop_duplicates(subset=['Chave de Identificação do Atendimento:'])

            self.base_tratada=pd.merge(self.base_tratada,
                                    feedback_limpo,
                                    on='Chave de Identificação do Atendimento:',
                                    how='left')
            
            self.base_tratada['Tempo aplicaçao']=(self.base_tratada['Data monitoria calculo feedback']-self.base_tratada['Data do Feedback:']).dt.days
            self.base_tratada['Tempo aplicaçao']=self.base_tratada['Tempo aplicaçao']
            
    def tratamento8(self):
        self.base_tratada['Chave_bd_supervisor']=self.base_tratada['Chave_bd_supervisor'].astype(str)
        self.base_tratada['Chave_bd_supervisor'] = self.base_tratada['Chave_bd_supervisor'].astype(str).str.replace('.0', '', regex=False)
        
        # CORREÇÃO: Adicionando o mapeamento das colunas que recebem _x devido ao pd.merge do tratamento7
        self.base_tratada=self.base_tratada.rename(columns={
            'Matrícula do Consultor Monitorado:_x':'Matrícula do Consultor Monitorado:',
            'Criado por_x': 'Criado por',
            'Tipo de Item_x': 'Tipo de Item',
            'Caminho_x': 'Caminho'
        })
    
        self.base_tratada=self.base_tratada[['Nome do Avaliador:', 
                                            'Data da Monitoria:', 
                                            'Data do Atendimento:',
                                            'Matrícula do Consultor Monitorado:',
                                            'Chave de Identificação do Atendimento:',
                                            'Unidades citadas no atendimento', 
                                            'Origem do Atendimento:',
                                            'Canal de Atendimento:', 
                                            'Item checklist', 
                                            'Status conformidade', 
                                            'NCG',
                                            'Bloco', 
                                            'Classificação', 
                                            'Responsabilidade', 
                                            'Nota atendimento',
                                            'Nota humano', 
                                            'Nota Regra negocio', 
                                            'Tipo de avaliador',
                                            'Mes ref monitoria', 
                                            'Ano ref monitoria', 
                                            'Dado ref monitoria',
                                            'Chave_bd_supervisor', 
                                            'Data monitoria calculo feedback',
                                            'Matrícula do Consultor Monitorado:_y', 
                                            'Data do Feedback:',
                                            'Nome do gestor responsável pelo feedback:',
                                            'Parecer sobre pontuações no bloco Condução:',
                                            'Condução - Item(s) Contestado(s):',
                                            'Parecer sobre pontuações no bloco Agendamento de Consulta:',
                                            'Agendamento de Consulta - Item(s) Contestado(s):',
                                            'Parecer sobre pontuações no bloco Agendamento de Exame:',
                                            'Agendamento de Exame - Item(s) Contestado(s):',
                                            'Observações sobre o feedback:', 
                                            'Criado por', 
                                            'Tipo de Item',
                                            'Caminho', 
                                            'Tempo aplicaçao',
                                            'Observações sobre processos não concluídos:']]

    def tratamento9(self):
        self.base_bd_novo['Chave_bd_supervisor']=0
        self.base_bd_novo['Meta supervisor']=0

        self.base_bd_novo=self.base_bd_novo[['ANO/MES','Chave_bd_supervisor', 'GERENTE', 'COORDENADOR', 'SUPERVISOR','MATRICULA','NOME','SIGLA','REGIONAL','LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISCO', 'LOGIN DYNAMO', 'LOGIN HIGH CONECT', 'LOGIN INGENIUM ', 'LOGIN INSTANT SOLUTION', 'LOGIN MEDCLINIC', 'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN MV H9J', 'LOGIN SISAGENDA', 'LOGIN SMART', 'LOGIN SMARTNX', 'LOGIN TASY', 'LOGIN TWILIO', 'LOGIN VOIP', 'Meta supervisor']]
        self.base_bd_novo['ANO/MES'] =self.base_bd_novo ['ANO/MES'].astype(str)
        textos_excluir = ['Filtros aplicados:','situacao é ATIVO','sigla não é ALMAVIVA, NEOBPO ou TEL','Parâmetro é Sigla']
        for texto in textos_excluir:
            self.base_bd_novo = self.base_bd_novo[~self.base_bd_novo['ANO/MES'].str.contains(texto, na=False)]
        
        # Correção da Geração de Chave da Base Nova: Formatando string e criando Chave Mensal igual
        self.base_bd_novo['Data_temp'] = pd.to_datetime(self.base_bd_novo['ANO/MES'], errors='coerce')
        self.base_bd_novo['Matrícula STR'] = self.base_bd_novo['MATRICULA'].astype(str).str.replace(r'\.0$', '', regex=True)
        self.base_bd_novo['Chave_bd_supervisor'] = self.base_bd_novo['Data_temp'].dt.strftime('%Y-%m') + self.base_bd_novo['Matrícula STR']
        
        meses_nomes = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho", 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
        self.base_bd_novo['ANO/MES'] = self.base_bd_novo['Data_temp'].dt.month.map(meses_nomes) + " de " + self.base_bd_novo['Data_temp'].dt.year.astype(str).str.replace(r'\.0$', '', regex=True)
        
        # Removendo duplicatas para previnir explosão cartesiana
        self.base_bd_novo = self.base_bd_novo.drop_duplicates('Chave_bd_supervisor')
        
        self.base_bd_novo['Nº operador']=self.base_bd_novo.groupby(['NOME'])['NOME'].transform('count')
        self.base_bd_novo=self.base_bd_novo.rename(columns={'LOGIN INGENIUM ':'LOGIN INGENIUM'})

        
    def tratamento10(self):
        self.base_bd=self.base_bd.rename(columns={'MATRICULA':'MATRICULA'})
        self.base_bd_novo=self.base_bd_novo.rename(columns={'MATRICULA':'MATRICULA'})
        self.base_bd = self.base_bd.reset_index(drop=True)
        self.base_bd_novo = self.base_bd_novo.reset_index(drop=True)
     
        self.base_bd = self.base_bd[['ANO/MES','Chave_bd_supervisor', 'GERENTE', 'COORDENADOR', 'SUPERVISOR','MATRICULA','NOME','SIGLA','REGIONAL','LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISCO', 'LOGIN DYNAMO', 'LOGIN HIGH CONECT', 'LOGIN INGENIUM', 'LOGIN INSTANT SOLUTION', 'LOGIN MEDCLINIC', 'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN MV H9J', 'LOGIN SISAGENDA', 'LOGIN SMART', 'LOGIN SMARTNX', 'LOGIN TASY', 'LOGIN TWILIO', 'LOGIN VOIP', 'Meta supervisor']]
        self.base_bd_novo = self.base_bd_novo[['ANO/MES','Chave_bd_supervisor', 'GERENTE', 'COORDENADOR', 'SUPERVISOR','MATRICULA','NOME','SIGLA','REGIONAL','LOGIN AVAYA', 'LOGIN BOT', 'LOGIN CISCO', 'LOGIN DYNAMO', 'LOGIN HIGH CONECT', 'LOGIN INGENIUM', 'LOGIN INSTANT SOLUTION', 'LOGIN MEDCLINIC', 'LOGIN MV CROSS', 'LOGIN MV FLEX', 'LOGIN MV H9J', 'LOGIN SISAGENDA', 'LOGIN SMART', 'LOGIN SMARTNX', 'LOGIN TASY', 'LOGIN TWILIO', 'LOGIN VOIP', 'Meta supervisor']]

        self.base_bd_empilhado = pd.concat([self.base_bd,self.base_bd_novo], ignore_index=True)
        
        # CORREÇÃO PARA O [Errno 28] No space left on device: Remoção de chaves duplicadas no empilhado antes do Merge
        self.base_bd_empilhado = self.base_bd_empilhado.drop_duplicates(subset=['Chave_bd_supervisor'])
        
        # Correção do mapping para suportar de forma dinâmica independentemente se for 2025 ou 2026
        meses_para_int = {
            "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
            "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
        }
        
        self.base_bd_empilhado['mes_ref'] = self.base_bd_empilhado['ANO/MES'].str.split(' ').str[0].map(meses_para_int)

    def tratamento11(self):
        # Correção do Typo: Alterado de 'Chave_bd_supervisor_x' para 'Chave_bd_supervisor'
        self.base_tratada=pd.merge(self.base_tratada,
                                  self.base_bd_empilhado,
                                  on='Chave_bd_supervisor',
                                  how='left')
        
        self.base_final=self.base_tratada
        self.base_final=self.base_final.rename(columns={'Dado ref monitoria':'Nº mes ref monitoria'})
        
        
    def tratamento12(self):
        self.dime_ncg_nc=self.base_final[['Chave de Identificação do Atendimento:','NCG']].drop_duplicates('Chave de Identificação do Atendimento:')
        self.dime_conformidade = (self.base_final[self.base_final['Status conformidade'] == 'NC - Não Conforme'][['Chave de Identificação do Atendimento:', 'Item checklist', 'Status conformidade']].drop_duplicates('Chave de Identificação do Atendimento:'))
        self.dime_supervisor=(self.base_bd_novo[['SUPERVISOR']]).drop_duplicates()
        self.dime_feedback=self.base_final[['Chave de Identificação do Atendimento:','Matrícula do Consultor Monitorado:','Tipo de avaliador','Nota humano','Data da Monitoria:','ANO/MES','Tempo aplicaçao','Observações sobre o feedback:']].drop_duplicates('Chave de Identificação do Atendimento:')
        self.dime_itens=self.base_final[['Chave de Identificação do Atendimento:','Item checklist','Status conformidade']].drop_duplicates('Chave de Identificação do Atendimento:')
        self.dime_coordenador =(self.base_bd_novo[['COORDENADOR']]).drop_duplicates()
        self.dime_gerente =(self.base_bd_novo[['GERENTE']]).drop_duplicates()
        return self.dime_ncg_nc,self.dime_conformidade,self.dime_feedback,self,self.dime_itens,self.dime_coordenador,self.dime_gerente,self.dime_supervisor
        
dataset=Bases()
dataset.carregar()
dataset.calendario()
dataset.concatenar()
dataset.tratamento1()
dataset.tratamento2()
dataset.tratamento3()
dataset.tratamento4()
dataset.tratamento5()
dataset.tratamento6()
dataset.tratamento7()
dataset.tratamento8()
dataset.tratamento9()
dataset.tratamento10()
dataset.tratamento11()
dataset.tratamento12()

base_ncg_nc = dataset.dime_ncg_nc
base_conformidade = dataset.dime_conformidade
base_itens = dataset.dime_itens
base_supervisor = dataset.dime_supervisor
base_coordenador = dataset.dime_coordenador
base_gerente = dataset.dime_gerente
base_feedback = dataset.dime_feedback
base_final = dataset.base_final
base_BD = dataset.base_bd_empilhado