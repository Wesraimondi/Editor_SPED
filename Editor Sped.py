# ==============================================================================
# Editor de Arquivos SPED Fiscal / EFD
# Data da última alteração: 2026-09-16
# Descrição das alterações realizadas:
# - Adicionada a função corrigir_0200_ncm_vazio_ou_zerado (Bloco 0 / Auditoria) para localizar automaticamente registros 0200 onde o Campo 8 (NCM) esteja vazio ou com '00', exibir janela interativa detalhada com tabela de produtos, permitindo visualizar e informar o valor correto de NCM individualmente (duplo-clique), em lote ou via assistente passo a passo com consulta integrada à tabela TIPI e atualização direta no arquivo SPED.
# - Atualizada a função listar_campos_c100_25_26 (Análise - Resumo C100 / Documentos Fiscais) para extrair e exibir na tabela e na exportação em XLS as colunas de CST de PIS e CST de COFINS para os registros (C100/C170/C175, A100/A170, C500/C501/C505, D100/D101/D105, F100, F120, F600).
# - Adicionada a função verificar_e_importar_0200_faltantes_h010 para validar itens do inventário (H010 campo 2) contra o cadastro de itens (0200 campo 2), exibir janela interativa com os faltantes e permitir selecionar arquivo TXT/SPED de referência para localizar, extrair e inserir os registros 0200 completos na sequência do Bloco 0, recalculando os totalizadores 0990 e Bloco 9.
# - Ajustado o rateio de PIS/COFINS por CFOP (aplicar_pis_cofins_c100_por_cfop) para priorizar CFOPs geradores de receita/tributados em notas multi-CFOP (ex.: 5124 industrialização vs 5902 retorno, 6124 vs 6902), eliminando distorções de rateio contra o Athena.
# - Atualizada a paginação do PDF no Relatório Comparador Athena x SPED Fiscal para aplicar limite por CFOP individual (até 40 notas/CFOP), garantindo que todos os CFOPs com divergências sejam impressos no PDF sem que um único CFOP com muitas notas (ex: 1352) bloqueie os demais.
# - Corrigido o Relatório Comparador Athena x SPED Fiscal (PDF / XLS) para desambiguar notas fiscais com números repetidos de participantes/razões sociais diferentes, utilizando a Chave de Acesso e chave composta (Número + CNPJ/Razão Social) nas funções carregar_athena_por_nota_txt, carregar_sped_por_nota_cfop, carregar_pis_cofins_athena_por_nota_txt e carregar_pis_cofins_c100_por_nota, eliminando falsas diferenças e exibindo a Razão Social nas tabelas.
# - Atualizado o Relatório Comparador Athena x SPED Fiscal (PDF / XLS) na seção
#   "Notas com Diferenca de PIS/COFINS": adicionada a separação e agrupamento
#   por CFOP das notas com divergência nos campos 26 e 27 do C100 contra o TXT de vendas Athena.
# - Ajustada a função corrigir_frete_cst56_para_66_fornecedores (Bloco D): para D101/D105 (CST 56->66, fornecedores FOR000002709 e FOR000002427), mantido o campo 5 original inalterado, definido o campo 6 igual ao campo 3 (base) e calculados os impostos no campo 8 = campo 3 * alíquota (PIS 1,2375%, COFINS 5,7%).
# - Adicionada a função alterar_pis_cofins_d101_d105_por_txt (Bloco D) para importar TXT (numero_documento;pis;cofins) e atualizar os valores dos registros D101 e D105 vinculados ao D100.
# - Adicionado o campo Entrada/Saída (IND_OPER do registro pai C100) na função listar_campos_c170_cfop_pis_cofins (Análise - Matriz C170) para visualização na tabela e exportação em XLS.
# - Adicionado cabeçalho documentado de histórico e conformidade com AGENTS.md.
# - Adicionado o Relatório Comparador Athena x SPED Contribuições (PDF / XLS),
#   permitindo auditoria cruzada de PIS/COFINS, bases de cálculo e faturamento
#   por CFOP, por nota fiscal e com batimento do Bloco M (M200/M600).
# - Centralizados todos os imports de bibliotecas (ReportLab, SQLite3, Locale,
#   Threading, Pandas, Requests, etc.) no topo do arquivo.
# - Removidas funções duplicadas (carregar_athena_por_nota_txt e gerar_relatorio_totalizador).
# - Implementado carregamento assíncrono (threading) para leitura de arquivos SPED
#   extensos sem travar a interface gráfica.
# - Adicionada barra de progresso visual na barra de status durante operações longas.
# - Adicionado filtro rápido por tipo de registro/bloco no painel superior (C100, C170, 0200, D100, H010, etc.).
# - Adicionados atalhos globais de teclado (Ctrl+O, Ctrl+S, Ctrl+Shift+S, Ctrl+F, F3).
# - Adicionado controle de estado de edição (asterisco '*' no título e confirmação de salvamento ao fechar).
# - Aplicada estilização visual ttk.Style moderna com alternância de cores (zebrado).
# - Organizada e padronizada a barra de menus com nomes profissionais, agrupamento por blocos e ícones intuitivos.
# - Corrigido o zoom de fonte (Aumentar / Diminuir) ajustando o Treeview e a altura das linhas (rowheight) em tempo real sem mensagens popups.
# - Adicionado o Gerenciador de Leiautes de Registros SPED (janela para visualizar, pesquisar, cadastrar novos registros e editar campos com salvamento local).
# - Corrigida a identificação de registros numéricos com zeros à esquerda (ex: 0000, 0150, 0200, 0100) no evento de clique da tabela (on_tree_select).
# - Atualizado o mapeamento de NCM x CEST no método update_0200_cest a partir da tabela oficial CEST12.xlsx (742 correspondências).
# - Padronizadas todas as funções de processamento/edição para informarem a quantidade exata de alterações efetuadas via caixa de diálogo.
# - Implementado recálculo automático do Bloco 9 (9900, 9990, 9999) ao salvar arquivos.
# - Implementado Módulo de Pré-Validação Fiscal (validação de DV de NFe/CTe, CNPJ/CPF e totais C100 vs C170).
# - Adicionado Exportador de Registros Selecionados para Excel (.xlsx) com cabeçalhos oficiais.
# - Adicionado Filtro Combinado Avançado (Tipo de Registro + Busca por Texto simultâneo).
# - Adicionado Histórico de Desfazer (Undo Ctrl+Z) para reversão de ações.
# - Implementado o Dashboard Visual de Apuração Fiscal (indicadores de faturamento, impostos, Bloco E110 e M200/M600).
# - Implementado o Comparador de Dois Arquivos SPED (cruzamento ERP x Transmitido com notas faltantes e exportação Excel).
# - Implementado o Auditor de Matriz Tributária (CFOP x CST x NCM) com correção automática em lote em 1-clique.
# - Implementado o Gerenciador de Regras / Macros Personalizadas com salvamento em JSON (sped_rules_custom.json).
# - Adicionado o Módulo de Auditoria de Campos Obrigatórios em Branco (alertando quando descrições ou códigos essenciais estão vazios em registros como 0200, 0150, C100, C170, 0000).
# - Atualizada a ação de duplo-clique nas relatórios de auditoria para abrir diretamente a janela detalhada de edição de campos da linha (abrir_editor_linha).
# - Ajustadas as janelas de auditoria para permanecerem abertas ao dar duplo-clique, permitindo corrigir múltiplos itens sequencialmente sem fechar o relatório.
# - Adicionado o Menu Ajuda completo com o Manual de Uso Interativo (F1), Tabela de Atalhos de Teclado e documentação detalhada de todos os módulos.
# ==============================================================================

import os
import re
import time
import datetime
import json
import locale
import sqlite3
import threading
import unicodedata
from collections import defaultdict
from pathlib import Path

import tkinter as tk
from tkinter import Menu, filedialog, messagebox, ttk, scrolledtext, simpledialog
import pandas as pd
import requests
from fpdf import FPDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
def consolidar_totais_sped(dados):
    registros_c190 = []
    notas_canceladas = set()
    numeros_notas_saida = set()

    for line in dados:
        try:
            campos = line.strip().split('|')
            if not campos or len(campos) < 2:
                continue

            if campos[1] == 'C100' and len(campos) > 8:
                cod_sit = campos[6]
                ind_oper = campos[2]
                try:
                    numero_nota = int(campos[8])
                    if cod_sit == '02':
                        notas_canceladas.add(numero_nota)
                    if ind_oper == '1':
                        numeros_notas_saida.add(numero_nota)
                except (ValueError, IndexError):
                    continue
            elif campos[1] == 'C190':
                registros_c190.append(campos)
        except (ValueError, IndexError):
            continue

    if not registros_c190:
        return None

    faltantes = []
    if numeros_notas_saida:
        inicio = min(numeros_notas_saida)
        fim = max(numeros_notas_saida)
        numeros_considerados = numeros_notas_saida - notas_canceladas
        faltantes = sorted(list(set(range(inicio, fim + 1)) - numeros_considerados))

    cfop_data = agrupar_por_cfop(registros_c190)
    aplicar_pis_cofins_c100_por_cfop(dados, cfop_data)

    for prefixo in ('|D190|', '|D590|', '|C590|'):
        registros = [line.strip().split('|') for line in dados if line.startswith(prefixo)]
        for reg in registros:
            if len(reg) <= 3:
                continue
            cfop = reg[3]
            if cfop not in cfop_data:
                cfop_data[cfop] = {
                    'valor_contabil': 0,
                    'valor_icms': 0,
                    'valor_ipi': 0,
                    'valor_pis': 0,
                    'valor_cofins': 0,
                    'valor_st': 0
                }
            try:
                cfop_data[cfop]['valor_contabil'] += converter_valor_decimal(reg[5] if len(reg) > 5 else "")
                cfop_data[cfop]['valor_icms'] += converter_valor_decimal(reg[7] if len(reg) > 7 else "")
            except (ValueError, IndexError):
                continue

    return {
        'cfop_data': dict(sorted(cfop_data.items())),
        'notas_canceladas': sorted(notas_canceladas),
        'faltantes': faltantes,
        'total_linhas': len(dados),
    }


def consolidar_totais_sped_contribuicoes(dados):
    cfop_data = defaultdict(lambda: {
        'valor_contabil': 0.0,
        'bc_pis': 0.0,
        'valor_pis': 0.0,
        'bc_cofins': 0.0,
        'valor_cofins': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_st': 0.0
    })

    totais_m200_m600 = {
        'pis_apurado': 0.0,
        'cofins_apurado': 0.0,
        'tem_bloco_m': False
    }

    notas_canceladas = set()
    numeros_notas_saida = set()
    total_linhas = len(dados)

    c100_cancelado = False
    d100_cancelado = False
    d100_vlr_doc = 0.0
    ind_oper_atual = "1"

    for line in dados:
        campos = line.strip().split('|')
        if len(campos) < 2:
            continue
        reg = campos[1]

        if reg == 'C100':
            c100_cancelado = (campos[6].strip() if len(campos) > 6 else "") == "02"
            ind_oper_atual = campos[2].strip() if len(campos) > 2 else "1"
            try:
                num_doc_str = campos[8].strip() if len(campos) > 8 else ""
                if num_doc_str.isdigit():
                    num_doc = int(num_doc_str)
                    if c100_cancelado:
                        notas_canceladas.add(num_doc)
                    elif ind_oper_atual == '1':
                        numeros_notas_saida.add(num_doc)
            except (ValueError, IndexError):
                pass
            continue

        if reg == 'C170' and not c100_cancelado:
            cfop = campos[11].strip() if len(campos) > 11 else ""
            if cfop:
                vlr_item = converter_valor_decimal(campos[7] if len(campos) > 7 else "")
                bc_pis = converter_valor_decimal(campos[26] if len(campos) > 26 else "")
                vlr_pis = converter_valor_decimal(campos[30] if len(campos) > 30 else "")
                bc_cofins = converter_valor_decimal(campos[32] if len(campos) > 32 else "")
                vlr_cofins = converter_valor_decimal(campos[36] if len(campos) > 36 else "")

                cfop_data[cfop]['valor_contabil'] += vlr_item
                cfop_data[cfop]['bc_pis'] += bc_pis
                cfop_data[cfop]['valor_pis'] += vlr_pis
                cfop_data[cfop]['bc_cofins'] += bc_cofins
                cfop_data[cfop]['valor_cofins'] += vlr_cofins
            continue

        if reg == 'C175' and not c100_cancelado:
            cfop = campos[2].strip() if len(campos) > 2 else ""
            if cfop:
                vlr_opr = converter_valor_decimal(campos[3] if len(campos) > 3 else "")
                bc_pis = converter_valor_decimal(campos[6] if len(campos) > 6 else "")
                vlr_pis = converter_valor_decimal(campos[8] if len(campos) > 8 else "")
                bc_cofins = converter_valor_decimal(campos[10] if len(campos) > 10 else "")
                vlr_cofins = converter_valor_decimal(campos[12] if len(campos) > 12 else "")

                cfop_data[cfop]['valor_contabil'] += vlr_opr
                cfop_data[cfop]['bc_pis'] += bc_pis
                cfop_data[cfop]['valor_pis'] += vlr_pis
                cfop_data[cfop]['bc_cofins'] += bc_cofins
                cfop_data[cfop]['valor_cofins'] += vlr_cofins
            continue

        if reg == 'A170':
            vlr_item = converter_valor_decimal(campos[5] if len(campos) > 5 else "")
            bc_pis = converter_valor_decimal(campos[10] if len(campos) > 10 else "")
            vlr_pis = converter_valor_decimal(campos[12] if len(campos) > 12 else "")
            bc_cofins = converter_valor_decimal(campos[14] if len(campos) > 14 else "")
            vlr_cofins = converter_valor_decimal(campos[16] if len(campos) > 16 else "")

            cfop_data['SERVICOS']['valor_contabil'] += vlr_item
            cfop_data['SERVICOS']['bc_pis'] += bc_pis
            cfop_data['SERVICOS']['valor_pis'] += vlr_pis
            cfop_data['SERVICOS']['bc_cofins'] += bc_cofins
            cfop_data['SERVICOS']['valor_cofins'] += vlr_cofins
            continue

        if reg == 'D100':
            d100_cancelado = (campos[6].strip() if len(campos) > 6 else "") == "02"
            d100_vlr_doc = converter_valor_decimal(campos[15] if len(campos) > 15 else "")
            continue

        if reg == 'D101' and not d100_cancelado:
            bc_pis = converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            vlr_pis = converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            cfop_data['1352']['valor_contabil'] += d100_vlr_doc
            cfop_data['1352']['bc_pis'] += bc_pis
            cfop_data['1352']['valor_pis'] += vlr_pis
            continue

        if reg == 'D105' and not d100_cancelado:
            bc_cofins = converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            vlr_cofins = converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            cfop_data['1352']['bc_cofins'] += bc_cofins
            cfop_data['1352']['valor_cofins'] += vlr_cofins
            continue

        if reg == 'F100':
            vlr_oper = converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            bc_pis = converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            vlr_pis = converter_valor_decimal(campos[10] if len(campos) > 10 else "")
            bc_cofins = converter_valor_decimal(campos[12] if len(campos) > 12 else "")
            vlr_cofins = converter_valor_decimal(campos[14] if len(campos) > 14 else "")

            cfop_data['F100']['valor_contabil'] += vlr_oper
            cfop_data['F100']['bc_pis'] += bc_pis
            cfop_data['F100']['valor_pis'] += vlr_pis
            cfop_data['F100']['bc_cofins'] += bc_cofins
            cfop_data['F100']['valor_cofins'] += vlr_cofins
            continue

        if reg == 'M200':
            totais_m200_m600['tem_bloco_m'] = True
            vlr = converter_valor_decimal(campos[12] if len(campos) > 12 else (campos[8] if len(campos) > 8 else ""))
            totais_m200_m600['pis_apurado'] += vlr
            continue

        if reg == 'M600':
            totais_m200_m600['tem_bloco_m'] = True
            vlr = converter_valor_decimal(campos[12] if len(campos) > 12 else (campos[8] if len(campos) > 8 else ""))
            totais_m200_m600['cofins_apurado'] += vlr
            continue

    if not cfop_data:
        for line in dados:
            campos = line.strip().split('|')
            if len(campos) > 2 and campos[1] == 'C100':
                if (campos[6].strip() if len(campos) > 6 else "") == "02":
                    continue
                vlr_contabil = converter_valor_decimal(campos[12] if len(campos) > 12 else "")
                vlr_pis = converter_valor_decimal(campos[26] if len(campos) > 26 else "")
                vlr_cofins = converter_valor_decimal(campos[27] if len(campos) > 27 else "")
                cfop_data['C100']['valor_contabil'] += vlr_contabil
                cfop_data['C100']['valor_pis'] += vlr_pis
                cfop_data['C100']['valor_cofins'] += vlr_cofins

    faltantes = []
    if numeros_notas_saida:
        inicio = min(numeros_notas_saida)
        fim = max(numeros_notas_saida)
        numeros_considerados = numeros_notas_saida - notas_canceladas
        faltantes = sorted(list(set(range(inicio, fim + 1)) - numeros_considerados))

    return {
        'cfop_data': dict(sorted(cfop_data.items())),
        'totais_bloco_m': totais_m200_m600,
        'notas_canceladas': sorted(notas_canceladas),
        'faltantes': faltantes,
        'total_linhas': total_linhas
    }


def carregar_sped_contribuicoes_por_nota(dados):
    totais = defaultdict(lambda: {
        'nota': '',
        'cfop': '',
        'chave': '',
        'origem_sped': set(),
        'valor_contabil': 0.0,
        'bc_pis': 0.0,
        'valor_pis': 0.0,
        'bc_cofins': 0.0,
        'valor_cofins': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_st': 0.0
    })

    nota_atual = ''
    chave_atual = ''
    situacao_atual = ''
    d100_vlr_doc = 0.0

    for line in dados:
        campos = line.strip().split('|')
        if len(campos) < 2:
            continue
        reg = campos[1]

        if reg == 'C100':
            nota_atual = campos[8].strip() if len(campos) > 8 else ''
            chave_atual = campos[9].strip() if len(campos) > 9 else ''
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            continue

        if reg == 'A100':
            nota_atual = campos[8].strip() if len(campos) > 8 else ''
            chave_atual = campos[9].strip() if len(campos) > 9 else ''
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            continue

        if reg == 'D100':
            nota_atual = campos[9].strip() if len(campos) > 9 else ''
            chave_atual = campos[10].strip() if len(campos) > 10 else ''
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            d100_vlr_doc = converter_valor_decimal(campos[15] if len(campos) > 15 else "")
            continue

        if situacao_atual == '02' or not nota_atual:
            continue

        if reg == 'C170':
            cfop = campos[11].strip() if len(campos) > 11 else ''
            if not cfop:
                continue
            identificador = (nota_atual, cfop)
            totais[identificador]['nota'] = nota_atual
            totais[identificador]['cfop'] = cfop
            totais[identificador]['chave'] = chave_atual
            totais[identificador]['origem_sped'].add('C170')
            totais[identificador]['valor_contabil'] += converter_valor_decimal(campos[7] if len(campos) > 7 else "")
            totais[identificador]['bc_pis'] += converter_valor_decimal(campos[26] if len(campos) > 26 else "")
            totais[identificador]['valor_pis'] += converter_valor_decimal(campos[30] if len(campos) > 30 else "")
            totais[identificador]['bc_cofins'] += converter_valor_decimal(campos[32] if len(campos) > 32 else "")
            totais[identificador]['valor_cofins'] += converter_valor_decimal(campos[36] if len(campos) > 36 else "")
            continue

        if reg == 'C175':
            cfop = campos[2].strip() if len(campos) > 2 else ''
            if not cfop:
                continue
            identificador = (nota_atual, cfop)
            totais[identificador]['nota'] = nota_atual
            totais[identificador]['cfop'] = cfop
            totais[identificador]['chave'] = chave_atual
            totais[identificador]['origem_sped'].add('C175')
            totais[identificador]['valor_contabil'] += converter_valor_decimal(campos[3] if len(campos) > 3 else "")
            totais[identificador]['bc_pis'] += converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            totais[identificador]['valor_pis'] += converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            totais[identificador]['bc_cofins'] += converter_valor_decimal(campos[10] if len(campos) > 10 else "")
            totais[identificador]['valor_cofins'] += converter_valor_decimal(campos[12] if len(campos) > 12 else "")
            continue

        if reg == 'A170':
            identificador = (nota_atual, 'SERVICOS')
            totais[identificador]['nota'] = nota_atual
            totais[identificador]['cfop'] = 'SERVICOS'
            totais[identificador]['chave'] = chave_atual
            totais[identificador]['origem_sped'].add('A170')
            totais[identificador]['valor_contabil'] += converter_valor_decimal(campos[5] if len(campos) > 5 else "")
            totais[identificador]['bc_pis'] += converter_valor_decimal(campos[10] if len(campos) > 10 else "")
            totais[identificador]['valor_pis'] += converter_valor_decimal(campos[12] if len(campos) > 12 else "")
            totais[identificador]['bc_cofins'] += converter_valor_decimal(campos[14] if len(campos) > 14 else "")
            totais[identificador]['valor_cofins'] += converter_valor_decimal(campos[16] if len(campos) > 16 else "")
            continue

        if reg == 'D101':
            identificador = (nota_atual, '1352')
            totais[identificador]['nota'] = nota_atual
            totais[identificador]['cfop'] = '1352'
            totais[identificador]['chave'] = chave_atual
            totais[identificador]['origem_sped'].add('D101')
            totais[identificador]['valor_contabil'] += d100_vlr_doc
            totais[identificador]['bc_pis'] += converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            totais[identificador]['valor_pis'] += converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            continue

        if reg == 'D105':
            identificador = (nota_atual, '1352')
            totais[identificador]['nota'] = nota_atual
            totais[identificador]['cfop'] = '1352'
            totais[identificador]['chave'] = chave_atual
            totais[identificador]['origem_sped'].add('D105')
            totais[identificador]['bc_cofins'] += converter_valor_decimal(campos[6] if len(campos) > 6 else "")
            totais[identificador]['valor_cofins'] += converter_valor_decimal(campos[8] if len(campos) > 8 else "")
            continue

    if not totais:
        for line in dados:
            campos = line.strip().split('|')
            if len(campos) > 2 and campos[1] == 'C100':
                if (campos[6].strip() if len(campos) > 6 else "") == "02":
                    continue
                nota = campos[8].strip() if len(campos) > 8 else ''
                chave = campos[9].strip() if len(campos) > 9 else ''
                if not nota:
                    continue
                identificador = (nota, 'C100')
                totais[identificador]['nota'] = nota
                totais[identificador]['cfop'] = 'C100'
                totais[identificador]['chave'] = chave
                totais[identificador]['origem_sped'].add('C100')
                totais[identificador]['valor_contabil'] += converter_valor_decimal(campos[12] if len(campos) > 12 else "")
                totais[identificador]['valor_pis'] += converter_valor_decimal(campos[26] if len(campos) > 26 else "")
                totais[identificador]['valor_cofins'] += converter_valor_decimal(campos[27] if len(campos) > 27 else "")

    totais_formatados = {}
    for chave, valores in sorted(totais.items()):
        valores['origem_sped'] = "/".join(sorted(valores['origem_sped']))
        totais_formatados[chave] = valores
    return totais_formatados


def montar_linhas_comparacao_contribuicoes(sped_cfop, athena_cfop):
    linhas = []
    todos_cfops = sorted(set(sped_cfop.keys()) | set(athena_cfop.keys()))
    for cfop in todos_cfops:
        sped = sped_cfop.get(cfop, {})
        athena = athena_cfop.get(cfop, {})
        linha = {'cfop': cfop}
        for campo in ('valor_contabil', 'bc_pis', 'valor_pis', 'bc_cofins', 'valor_cofins', 'valor_icms', 'valor_ipi', 'valor_st'):
            valor_sped = float(sped.get(campo, 0.0) or 0.0)
            valor_athena = float(athena.get(campo, 0.0) or 0.0)
            linha[f'{campo}_sped'] = valor_sped
            linha[f'{campo}_athena'] = valor_athena
            linha[f'dif_{campo}'] = valor_sped - valor_athena
        linhas.append(linha)
    return linhas


def carregar_athena_contribuicoes_por_nota_txt(caminho_arquivo):
    conteudo = ler_arquivo_texto_multiencoding(caminho_arquivo)
    linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
    if len(linhas) < 2:
        raise ValueError("O arquivo Athena nao possui dados por nota para comparacao.")

    cabecalhos = [coluna.strip() for coluna in linhas[0].split('\t')]
    indices = {}

    try:
        indices['cfop'] = obter_indice_primeira_coluna(cabecalhos, ['CFOP', 'Codigo', 'Código'])
    except ValueError:
        indices['cfop'] = None

    try:
        indices['nota'] = obter_indice_primeira_coluna(cabecalhos, ['Nº NFe', 'N° NFe', 'No NFe', 'Numero NFe', 'Numero NF', 'NFe'])
    except ValueError:
        indices['nota'] = None

    try:
        indices['chave'] = obter_indice_primeira_coluna(cabecalhos, ['Chave de Acesso', 'Chave Acesso', 'Chave'])
    except ValueError:
        indices['chave'] = None

    try:
        indices['valor_contabil'] = obter_indice_primeira_coluna(cabecalhos, ['Vlr. Contabil', 'Vlr. Contábil', 'Valor Contabil', 'Valor Contábil', 'Vlr. Operacao', 'Vlr. Operação'])
    except ValueError:
        indices['valor_contabil'] = None

    try:
        indices['valor_pis'] = obter_indice_primeira_coluna(cabecalhos, ['Vlr. PIS', 'Vlr PIS', 'Valor PIS'])
    except ValueError:
        indices['valor_pis'] = None

    try:
        indices['valor_cofins'] = obter_indice_primeira_coluna(cabecalhos, ['Vlr. COFINS', 'Vlr COFINS', 'Valor COFINS'])
    except ValueError:
        indices['valor_cofins'] = None

    if indices['nota'] is None:
        raise ValueError("Coluna com o numero da nota fiscal nao encontrada no arquivo Athena.")

    totais = defaultdict(lambda: {
        'nota': '',
        'cfop': '',
        'chave': '',
        'valor_contabil': 0.0,
        'valor_pis': 0.0,
        'valor_cofins': 0.0
    })

    for linha in linhas[1:]:
        colunas = linha.split('\t')
        nota = colunas[indices['nota']].strip() if len(colunas) > indices['nota'] else ""
        cfop = colunas[indices['cfop']].strip() if (indices['cfop'] is not None and len(colunas) > indices['cfop']) else "0000"
        if not nota:
            continue
        identificador = (nota, cfop)
        totais[identificador]['nota'] = nota
        totais[identificador]['cfop'] = cfop
        totais[identificador]['chave'] = colunas[indices['chave']].strip() if (indices['chave'] is not None and len(colunas) > indices['chave']) else ""
        if indices['valor_contabil'] is not None and len(colunas) > indices['valor_contabil']:
            totais[identificador]['valor_contabil'] += converter_valor_decimal(colunas[indices['valor_contabil']])
        if indices['valor_pis'] is not None and len(colunas) > indices['valor_pis']:
            totais[identificador]['valor_pis'] += converter_valor_decimal(colunas[indices['valor_pis']])
        if indices['valor_cofins'] is not None and len(colunas) > indices['valor_cofins']:
            totais[identificador]['valor_cofins'] += converter_valor_decimal(colunas[indices['valor_cofins']])

    return dict(sorted(totais.items()))


def montar_linhas_comparacao_por_nota_contribuicoes(sped_notas, athena_notas):
    linhas = []
    todos = sorted(set(sped_notas.keys()) | set(athena_notas.keys()))
    for identificador in todos:
        sped = sped_notas.get(identificador, {})
        athena = athena_notas.get(identificador)
        if athena is None:
            cands = [v for k, v in athena_notas.items() if k[0] == identificador[0]]
            athena = cands[0] if cands else {}

        linha = {
            'nota': identificador[0],
            'cfop': identificador[1],
            'chave_sped': sped.get('chave', ''),
            'chave_athena': athena.get('chave', ''),
            'origem_sped': sped.get('origem_sped', ''),
        }
        possui_diferenca = False
        for campo in ('valor_contabil', 'valor_pis', 'valor_cofins'):
            valor_sped = float(sped.get(campo, 0.0) or 0.0)
            valor_athena = float(athena.get(campo, 0.0) or 0.0)
            diferenca = valor_sped - valor_athena
            linha[f'{campo}_sped'] = valor_sped
            linha[f'{campo}_athena'] = valor_athena
            linha[f'dif_{campo}'] = diferenca
            if abs(diferenca) > 0.0001:
                possui_diferenca = True
        linha['possui_diferenca'] = possui_diferenca
        linhas.append(linha)
    return linhas


def exportar_comparacao_xls_contribuicoes(caminho_arquivo, linhas_comparacao, totais_sped, totais_athena, totais_por_tipo_sped, totais_por_tipo_athena, linhas_por_nota=None, totais_bloco_m=None):
    colunas = [
        ('CFOP', 'cfop'),
        ('SPED Valor Contabil', 'valor_contabil_sped'),
        ('Athena Valor Contabil', 'valor_contabil_athena'),
        ('Dif Valor Contabil', 'dif_valor_contabil'),
        ('SPED PIS', 'valor_pis_sped'),
        ('Athena PIS', 'valor_pis_athena'),
        ('Dif PIS', 'dif_valor_pis'),
        ('SPED COFINS', 'valor_cofins_sped'),
        ('Athena COFINS', 'valor_cofins_athena'),
        ('Dif COFINS', 'dif_valor_cofins'),
    ]
    def formatar_excel(valor):
        if isinstance(valor, (int, float)):
            return f"{valor:.2f}".replace('.', ',')
        return str(valor)

    def classe_diferenca(chave, valor):
        if not chave.startswith('dif_'):
            return ""
        if isinstance(valor, (int, float)) and abs(valor) > 0.0001:
            return ' class="dif"'
        return ""

    html = [
        '<html><head><meta charset="utf-8">',
        '<style>',
        'body { font-family: Calibri, Arial, sans-serif; color: #243447; padding: 18px; }',
        'h1 { color: #17324d; margin: 0 0 6px 0; font-size: 20pt; }',
        'h2 { color: #17324d; margin: 22px 0 8px 0; font-size: 13pt; }',
        '.subtitulo { color: #5b6b79; margin-bottom: 18px; font-size: 10pt; }',
        '.painel { display: inline-block; min-width: 180px; margin: 0 12px 12px 0; padding: 10px 12px; border: 1px solid #d9e2ec; background: #f8fbff; }',
        '.painel .rotulo { display: block; color: #5b6b79; font-size: 9pt; margin-bottom: 4px; }',
        '.painel .valor { display: block; color: #17324d; font-size: 12pt; font-weight: bold; }',
        'table { border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 10pt; width: 100%; }',
        'th, td { border: 1px solid #b8c4cf; padding: 5px 7px; text-align: center; }',
        'th { background: #17324d; color: #ffffff; font-weight: bold; }',
        'tr:nth-child(even) td { background: #f8fafc; }',
        '.dif { color: #c00000; font-weight: bold; background: #fde9e7; }',
        '.resumo { font-weight: bold; }',
        '.secao { margin-top: 18px; }',
        '</style></head><body>',
        '<h1>Relatorio Comparador Athena x SPED Contribuicoes</h1>',
        f'<div class="subtitulo">Arquivo Athena: {os.path.basename(caminho_arquivo).replace(".xls", ".txt")} | Gerado em: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</div>',
        '<div class="painel"><span class="rotulo">Total de CFOPs comparados</span><span class="valor">' + str(len(linhas_comparacao)) + '</span></div>',
        '<div class="painel"><span class="rotulo">CFOPs com diferenca</span><span class="valor">' + str(sum(1 for linha in linhas_comparacao if any(abs(linha.get(chave, 0.0)) > 0.0001 for chave in ("dif_valor_contabil", "dif_valor_pis", "dif_valor_cofins")))) + '</span></div>',
        '<div class="painel"><span class="rotulo">Notas com diferenca</span><span class="valor">' + str(sum(1 for linha in (linhas_por_nota or []) if linha.get("possui_diferenca"))) + '</span></div>'
    ]

    if totais_bloco_m and totais_bloco_m.get('tem_bloco_m'):
        html.append('<div class="painel"><span class="rotulo">PIS Apurado Bloco M (M200)</span><span class="valor">' + formatar_excel(totais_bloco_m.get('pis_apurado', 0.0)) + '</span></div>')
        html.append('<div class="painel"><span class="rotulo">COFINS Apurado Bloco M (M600)</span><span class="valor">' + formatar_excel(totais_bloco_m.get('cofins_apurado', 0.0)) + '</span></div>')

    html.extend([
        '<h2>Resumo por CFOP (PIS / COFINS)</h2>',
        '<table>',
        '<tr>' + ''.join(f'<th>{titulo}</th>' for titulo, _ in colunas) + '</tr>'
    ])

    for linha in linhas_comparacao:
        html.append('<tr>' + ''.join(
            f'<td{classe_diferenca(chave, linha.get(chave, ""))}>{formatar_excel(linha.get(chave, ""))}</td>'
            for _, chave in colunas
        ) + '</tr>')

    html.extend([
        '</table>',
        '<div class="secao"></div>',
        '<h2>Resumo Geral (Entradas vs Saídas)</h2>',
        '<table>',
        '<tr><th>Tipo</th><th>Base</th><th>Valor Contábil</th><th>PIS</th><th>COFINS</th></tr>'
    ])

    for tipo in ('Entrada', 'Saida'):
        totais_tipo_sped = totais_por_tipo_sped.get(tipo, {})
        totais_tipo_athena = totais_por_tipo_athena.get(tipo, {})
        for titulo, totais in (('SPED Contribuições', totais_tipo_sped), ('Athena', totais_tipo_athena)):
            html.append(
                '<tr class="resumo">' +
                ''.join([
                    f'<td>{tipo}</td>',
                    f'<td>{titulo}</td>',
                    f'<td>{formatar_excel(totais.get("valor_contabil", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_pis", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_cofins", 0.0))}</td>',
                ]) +
                '</tr>'
            )

        diferencas_resumo = {
            'valor_contabil': totais_tipo_sped.get('valor_contabil', 0.0) - totais_tipo_athena.get('valor_contabil', 0.0),
            'valor_pis': totais_tipo_sped.get('valor_pis', 0.0) - totais_tipo_athena.get('valor_pis', 0.0),
            'valor_cofins': totais_tipo_sped.get('valor_cofins', 0.0) - totais_tipo_athena.get('valor_cofins', 0.0),
        }
        html.append(
            '<tr class="resumo">' +
            ''.join([
                f'<td>{tipo}</td>',
                '<td>Diferenca (SPED - Athena)</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_contabil"])}</td>' if abs(diferencas_resumo["valor_contabil"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_contabil"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_pis"])}</td>' if abs(diferencas_resumo["valor_pis"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_pis"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_cofins"])}</td>' if abs(diferencas_resumo["valor_cofins"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_cofins"])}</td>',
            ]) +
            '</tr>'
        )

    if linhas_por_nota:
        divergencias_por_cfop = agrupar_divergencias_por_cfop(linhas_por_nota)
        colunas_notas = [
            ('Nota Fiscal', 'nota'),
            ('CFOP', 'cfop'),
            ('Origem SPED', 'origem_sped'),
            ('SPED Valor Contábil', 'valor_contabil_sped'),
            ('Athena Valor Contábil', 'valor_contabil_athena'),
            ('Dif Valor Contábil', 'dif_valor_contabil'),
            ('SPED PIS', 'valor_pis_sped'),
            ('Athena PIS', 'valor_pis_athena'),
            ('Dif PIS', 'dif_valor_pis'),
            ('SPED COFINS', 'valor_cofins_sped'),
            ('Athena COFINS', 'valor_cofins_athena'),
            ('Dif COFINS', 'dif_valor_cofins'),
            ('Chave SPED', 'chave_sped'),
            ('Chave Athena', 'chave_athena'),
        ]
        html.extend([
            '<div class="secao"></div>',
            '<h2>Notas Fiscais com Diferenças no SPED Contribuições</h2>',
        ])
        for cfop, linhas_cfop in divergencias_por_cfop:
            html.extend([
                f'<h2>CFOP {cfop}</h2>',
                '<table>',
                '<tr>' + ''.join(f'<th>{titulo}</th>' for titulo, _ in colunas_notas) + '</tr>'
            ])
            for linha in linhas_cfop:
                html.append('<tr>' + ''.join(
                    f'<td{classe_diferenca(chave, linha.get(chave, ""))}>{formatar_excel(linha.get(chave, ""))}</td>'
                    for _, chave in colunas_notas
                ) + '</tr>')
            html.append('</table>')

    html.extend(['</body></html>'])

    with open(caminho_arquivo, 'w', encoding='utf-8-sig') as arquivo:
        arquivo.write('\n'.join(html))


def gerar_pdf_comparacao_athena_sped_contribuicoes(linhas_comparacao, totais_sped, totais_athena, totais_por_tipo_sped, totais_por_tipo_athena, output_path, caminho_txt, linhas_por_nota=None, totais_bloco_m=None):
    doc = SimpleDocTemplate(output_path, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("Relatorio Comparador Athena x SPED Contribuicoes", styles['Title']))
    elements.append(Paragraph(f"Arquivo Athena: {os.path.basename(caminho_txt)} | Gerado em: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
    elements.append(Spacer(1, 15))

    def formatar_valor(valor):
        return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    if totais_bloco_m and totais_bloco_m.get('tem_bloco_m'):
        elements.append(Paragraph(f"<b>Apuracao Bloco M:</b> PIS Apurado (M200): R$ {formatar_valor(totais_bloco_m.get('pis_apurado', 0.0))} | COFINS Apurado (M600): R$ {formatar_valor(totais_bloco_m.get('cofins_apurado', 0.0))}", styles['Normal']))
        elements.append(Spacer(1, 10))

    data_resumo = [['Tipo', 'Base', 'Valor Contabil', 'PIS', 'COFINS']]
    for tipo in ('Entrada', 'Saida'):
        totais_tipo_sped = totais_por_tipo_sped.get(tipo, {})
        totais_tipo_athena = totais_por_tipo_athena.get(tipo, {})

        data_resumo.append([
            tipo,
            'SPED Contribuicoes',
            formatar_valor(totais_tipo_sped.get('valor_contabil', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_pis', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_cofins', 0.0)),
        ])
        data_resumo.append([
            tipo,
            'Athena',
            formatar_valor(totais_tipo_athena.get('valor_contabil', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_pis', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_cofins', 0.0)),
        ])
        dif_contabil = totais_tipo_sped.get('valor_contabil', 0.0) - totais_tipo_athena.get('valor_contabil', 0.0)
        dif_pis = totais_tipo_sped.get('valor_pis', 0.0) - totais_tipo_athena.get('valor_pis', 0.0)
        dif_cofins = totais_tipo_sped.get('valor_cofins', 0.0) - totais_tipo_athena.get('valor_cofins', 0.0)

        data_resumo.append([
            tipo,
            'Diferenca (SPED - Athena)',
            formatar_valor(dif_contabil),
            formatar_valor(dif_pis),
            formatar_valor(dif_cofins),
        ])

    table_resumo = Table(data_resumo, colWidths=[80, 150, 120, 120, 120])
    table_resumo.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#b8c4cf')),
    ]))
    elements.append(Paragraph("<b>Resumo Geral por Operacao</b>", styles['Heading2']))
    elements.append(table_resumo)
    elements.append(Spacer(1, 15))

    header_cfop = ['CFOP', 'SPED Contabil', 'Athena Contabil', 'Dif Contabil', 'SPED PIS', 'Athena PIS', 'Dif PIS', 'SPED COFINS', 'Athena COFINS', 'Dif COFINS']
    data_cfop = [header_cfop]

    for linha in linhas_comparacao:
        data_cfop.append([
            str(linha['cfop']),
            formatar_valor(linha['valor_contabil_sped']),
            formatar_valor(linha['valor_contabil_athena']),
            formatar_valor(linha['dif_valor_contabil']),
            formatar_valor(linha['valor_pis_sped']),
            formatar_valor(linha['valor_pis_athena']),
            formatar_valor(linha['dif_valor_pis']),
            formatar_valor(linha['valor_cofins_sped']),
            formatar_valor(linha['valor_cofins_athena']),
            formatar_valor(linha['dif_valor_cofins']),
        ])

    table_cfop = Table(data_cfop, colWidths=[55, 75, 75, 75, 75, 75, 75, 75, 75, 75])
    table_cfop.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#b8c4cf')),
    ]))
    elements.append(Paragraph("<b>Detalhamento por CFOP (PIS / COFINS)</b>", styles['Heading2']))
    elements.append(table_cfop)

    if linhas_por_nota:
        divergencias_por_cfop = agrupar_divergencias_por_cfop(linhas_por_nota)
        if divergencias_por_cfop:
            elements.append(Spacer(1, 15))
            elements.append(Paragraph("<b>Notas Fiscais com Diferencas no SPED Contribuicoes</b>", styles['Heading2']))
            for cfop, linhas_cfop in divergencias_por_cfop:
                elements.append(Paragraph(f"<b>CFOP {cfop}</b>", styles['Heading3']))
                header_nota = ['Nota', 'Origem', 'SPED Contabil', 'Athena Contabil', 'Dif Contabil', 'SPED PIS', 'Athena PIS', 'Dif PIS', 'SPED COFINS', 'Athena COFINS', 'Dif COFINS']
                data_nota = [header_nota]
                for l in linhas_cfop:
                    data_nota.append([
                        str(l.get('nota', '')),
                        str(l.get('origem_sped', '')),
                        formatar_valor(l.get('valor_contabil_sped', 0.0)),
                        formatar_valor(l.get('valor_contabil_athena', 0.0)),
                        formatar_valor(l.get('dif_valor_contabil', 0.0)),
                        formatar_valor(l.get('valor_pis_sped', 0.0)),
                        formatar_valor(l.get('valor_pis_athena', 0.0)),
                        formatar_valor(l.get('dif_valor_pis', 0.0)),
                        formatar_valor(l.get('valor_cofins_sped', 0.0)),
                        formatar_valor(l.get('valor_cofins_athena', 0.0)),
                        formatar_valor(l.get('dif_valor_cofins', 0.0)),
                    ])
                table_nota = Table(data_nota, colWidths=[45, 45, 68, 68, 68, 68, 68, 68, 68, 68, 68])
                table_nota.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7.5),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#b8c4cf')),
                ]))
                elements.append(table_nota)
                elements.append(Spacer(1, 10))

    doc.build(elements)


def ler_arquivo_texto_multiencoding(caminho_arquivo):
    for encoding in ('utf-8-sig', 'cp1252', 'latin-1'):
        try:
            with open(caminho_arquivo, 'r', encoding=encoding) as arquivo:
                return arquivo.read()
        except UnicodeDecodeError:
            continue
    with open(caminho_arquivo, 'r', encoding='utf-8', errors='ignore') as arquivo:
        return arquivo.read()


def normalizar_cabecalho(texto):
    texto = unicodedata.normalize('NFKD', str(texto or '').strip())
    texto = ''.join(ch for ch in texto if not unicodedata.combining(ch))
    return ''.join(ch for ch in texto.lower() if ch.isalnum())


def obter_indices_colunas(cabecalhos, mapa_colunas):
    cabecalhos_normalizados = {normalizar_cabecalho(coluna): indice for indice, coluna in enumerate(cabecalhos)}
    indices = {}
    for nome_coluna, alias in mapa_colunas.items():
        chave_normalizada = normalizar_cabecalho(nome_coluna)
        if chave_normalizada not in cabecalhos_normalizados:
            raise ValueError(f"Coluna obrigatoria nao encontrada no Athena: {nome_coluna}")
        indices[alias] = cabecalhos_normalizados[chave_normalizada]
    return indices


def obter_indice_primeira_coluna(cabecalhos, nomes_coluna):
    cabecalhos_normalizados = {normalizar_cabecalho(coluna): indice for indice, coluna in enumerate(cabecalhos)}
    for nome_coluna in nomes_coluna:
        chave_normalizada = normalizar_cabecalho(nome_coluna)
        if chave_normalizada in cabecalhos_normalizados:
            return cabecalhos_normalizados[chave_normalizada]
    raise ValueError(f"Coluna obrigatoria nao encontrada no Athena: {' / '.join(nomes_coluna)}")


def carregar_totais_athena_txt(caminho_arquivo):
    conteudo = ler_arquivo_texto_multiencoding(caminho_arquivo)
    linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
    if len(linhas) < 2:
        raise ValueError("O arquivo Athena nao possui dados para comparacao.")

    cabecalhos = [coluna.strip() for coluna in linhas[0].split('\t')]
    colunas_necessarias = {
        'Código': 'cfop',
        'Vlr. Contábil': 'valor_contabil',
        'Vlr. ICMS': 'valor_icms',
        'Vlr. IPI': 'valor_ipi',
        'Vlr. ICMS-ST': 'valor_icms_st',
        'Vlr. FCP': 'valor_fcp',
    }

    indices = {}
    for nome_coluna, alias in colunas_necessarias.items():
        if nome_coluna not in cabecalhos:
            raise ValueError(f"Coluna obrigatoria nao encontrada no Athena: {nome_coluna}")
        indices[alias] = cabecalhos.index(nome_coluna)

    totais = defaultdict(lambda: {
        'valor_contabil': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_pis': 0.0,
        'valor_cofins': 0.0,
        'valor_st': 0.0
    })

    for linha in linhas[1:]:
        colunas = linha.split('\t')
        cfop = colunas[indices['cfop']].strip() if len(colunas) > indices['cfop'] else ""
        if not cfop:
            continue

        somar_valores_cfop(
            totais,
            cfop,
            valor_contabil=converter_valor_decimal(colunas[indices['valor_contabil']] if len(colunas) > indices['valor_contabil'] else ""),
            valor_icms=converter_valor_decimal(colunas[indices['valor_icms']] if len(colunas) > indices['valor_icms'] else ""),
            valor_ipi=converter_valor_decimal(colunas[indices['valor_ipi']] if len(colunas) > indices['valor_ipi'] else ""),
            valor_st=converter_valor_decimal(colunas[indices['valor_icms_st']] if len(colunas) > indices['valor_icms_st'] else "") +
                     converter_valor_decimal(colunas[indices['valor_fcp']] if len(colunas) > indices['valor_fcp'] else "")
        )

    return dict(sorted(totais.items()))


def carregar_totais_athena_txt_v2(caminho_arquivo):
    conteudo = ler_arquivo_texto_multiencoding(caminho_arquivo)
    linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
    if len(linhas) < 2:
        raise ValueError("O arquivo Athena nao possui dados para comparacao.")

    cabecalhos = [coluna.strip() for coluna in linhas[0].split('\t')]
    indices = obter_indices_colunas(cabecalhos, {
        'Vlr. Contabil': 'valor_contabil',
        'Vlr. ICMS': 'valor_icms',
        'Vlr. IPI': 'valor_ipi',
        'Vlr. PIS': 'valor_pis',
        'Vlr. COFINS': 'valor_cofins',
        'Vlr. ICMS-ST': 'valor_icms_st',
        'Vlr. FCP-ST': 'valor_fcp_st',
    })
    indices['cfop'] = obter_indice_primeira_coluna(cabecalhos, ['CFOP', 'Codigo'])

    totais = defaultdict(lambda: {
        'valor_contabil': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_pis': 0.0,
        'valor_cofins': 0.0,
        'valor_st': 0.0
    })

    for linha in linhas[1:]:
        colunas = linha.split('\t')
        cfop = colunas[indices['cfop']].strip() if len(colunas) > indices['cfop'] else ""
        if not cfop:
            continue
        somar_valores_cfop(
            totais,
            cfop,
            valor_contabil=converter_valor_decimal(colunas[indices['valor_contabil']] if len(colunas) > indices['valor_contabil'] else ""),
            valor_icms=converter_valor_decimal(colunas[indices['valor_icms']] if len(colunas) > indices['valor_icms'] else ""),
            valor_ipi=converter_valor_decimal(colunas[indices['valor_ipi']] if len(colunas) > indices['valor_ipi'] else ""),
            valor_pis=converter_valor_decimal(colunas[indices['valor_pis']] if len(colunas) > indices['valor_pis'] else ""),
            valor_cofins=converter_valor_decimal(colunas[indices['valor_cofins']] if len(colunas) > indices['valor_cofins'] else ""),
            valor_st=converter_valor_decimal(colunas[indices['valor_icms_st']] if len(colunas) > indices['valor_icms_st'] else "") +
                     converter_valor_decimal(colunas[indices['valor_fcp_st']] if len(colunas) > indices['valor_fcp_st'] else "")
        )

    return dict(sorted(totais.items()))


def carregar_participantes_sped(dados):
    participantes = {}
    for line in dados:
        campos = line.strip().split('|')
        if len(campos) > 3 and campos[1] == '0150':
            cod_part = campos[2].strip()
            nome = campos[3].strip() if len(campos) > 3 else ''
            cnpj = ''.join(ch for ch in (campos[5] if len(campos) > 5 else '') if ch.isdigit())
            cpf = ''.join(ch for ch in (campos[6] if len(campos) > 6 else '') if ch.isdigit())
            participantes[cod_part] = {
                'nome': nome,
                'cnpj': cnpj or cpf
            }
    return participantes


def normalizar_numero_nota(nota):
    texto = str(nota or '').strip()
    if texto.isdigit():
        return str(int(texto))
    return texto


def gerar_chave_documento(chave, nota, cnpj='', razao=''):
    chv = str(chave or '').strip()
    if chv:
        return chv
    num = normalizar_numero_nota(nota)
    doc_id = ''.join(ch for ch in str(cnpj or '') if ch.isdigit())
    if doc_id:
        return f"{num}_{doc_id}"
    rz = str(razao or '').strip().upper()
    if rz:
        return f"{num}_{rz}"
    return num


def carregar_athena_por_nota_txt(caminho_arquivo):
    conteudo = ler_arquivo_texto_multiencoding(caminho_arquivo)
    linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
    if len(linhas) < 2:
        raise ValueError("O arquivo Athena nao possui dados por nota para comparacao.")

    cabecalhos = [coluna.strip() for coluna in linhas[0].split('\t')]
    indices = obter_indices_colunas(cabecalhos, {
        'CFOP': 'cfop',
        'Vlr. Contabil': 'valor_contabil',
        'Vlr. ICMS': 'valor_icms',
        'Vlr. IPI': 'valor_ipi',
        'Vlr. ICMS-ST': 'valor_icms_st',
        'Vlr. FCP-ST': 'valor_fcp_st',
    })
    indices['nota'] = obter_indice_primeira_coluna(cabecalhos, ['Nº NFe', 'No NFe', 'Numero NFe', 'Numero NF', 'NFe'])
    try:
        indices['chave'] = obter_indice_primeira_coluna(cabecalhos, ['Chave de Acesso', 'Chave Acesso', 'Chave'])
    except ValueError:
        indices['chave'] = None
    try:
        indices['razao'] = obter_indice_primeira_coluna(cabecalhos, ['Razão Social', 'Razao Social', 'Razão', 'Razao', 'Nome'])
    except ValueError:
        indices['razao'] = None
    try:
        indices['cnpj'] = obter_indice_primeira_coluna(cabecalhos, ['CNPJ', 'CPF', 'CNPJ/CPF'])
    except ValueError:
        indices['cnpj'] = None

    totais = defaultdict(lambda: {
        'nota': '',
        'razao_social': '',
        'cnpj': '',
        'cfop': '',
        'chave': '',
        'valor_contabil': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_st': 0.0
    })

    for linha in linhas[1:]:
        colunas = linha.split('\t')
        nota_raw = colunas[indices['nota']].strip() if len(colunas) > indices['nota'] else ""
        cfop = colunas[indices['cfop']].strip() if len(colunas) > indices['cfop'] else ""
        if not nota_raw or not cfop:
            continue
        nota = normalizar_numero_nota(nota_raw)
        chave = colunas[indices['chave']].strip() if indices['chave'] is not None and len(colunas) > indices['chave'] else ""
        razao = colunas[indices['razao']].strip() if indices['razao'] is not None and len(colunas) > indices['razao'] else ""
        cnpj = ''.join(ch for ch in (colunas[indices['cnpj']] if indices['cnpj'] is not None and len(colunas) > indices['cnpj'] else "") if ch.isdigit())

        doc_key = gerar_chave_documento(chave, nota, cnpj, razao)
        identificador = (doc_key, cfop)
        totais[identificador]['nota'] = nota
        totais[identificador]['razao_social'] = razao
        totais[identificador]['cnpj'] = cnpj
        totais[identificador]['cfop'] = cfop
        totais[identificador]['chave'] = chave
        totais[identificador]['valor_contabil'] += converter_valor_decimal(colunas[indices['valor_contabil']] if len(colunas) > indices['valor_contabil'] else "")
        totais[identificador]['valor_icms'] += converter_valor_decimal(colunas[indices['valor_icms']] if len(colunas) > indices['valor_icms'] else "")
        totais[identificador]['valor_ipi'] += converter_valor_decimal(colunas[indices['valor_ipi']] if len(colunas) > indices['valor_ipi'] else "")
        totais[identificador]['valor_st'] += (
            converter_valor_decimal(colunas[indices['valor_icms_st']] if len(colunas) > indices['valor_icms_st'] else "") +
            converter_valor_decimal(colunas[indices['valor_fcp_st']] if len(colunas) > indices['valor_fcp_st'] else "")
        )

    return dict(sorted(totais.items()))


def carregar_sped_por_nota_cfop(dados):
    participantes = carregar_participantes_sped(dados)
    totais = defaultdict(lambda: {
        'nota': '',
        'razao_social': '',
        'cnpj': '',
        'cfop': '',
        'chave': '',
        'origem_sped': set(),
        'valor_contabil': 0.0,
        'valor_icms': 0.0,
        'valor_ipi': 0.0,
        'valor_st': 0.0
    })

    nota_atual = ''
    doc_key_atual = ''
    chave_atual = ''
    razao_atual = ''
    cnpj_atual = ''
    situacao_atual = ''
    registro_pai_atual = ''

    for line in dados:
        campos = line.strip().split('|')
        if len(campos) < 2:
            continue
        registro = campos[1]
        if registro == 'C100':
            num_raw = campos[8].strip() if len(campos) > 8 else ''
            nota_atual = normalizar_numero_nota(num_raw)
            chave_atual = campos[9].strip() if len(campos) > 9 else ''
            part_cod = campos[4].strip() if len(campos) > 4 else ''
            part_info = participantes.get(part_cod, {})
            razao_atual = part_info.get('nome', '')
            cnpj_atual = part_info.get('cnpj', '')
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            registro_pai_atual = 'C100'
            doc_key_atual = gerar_chave_documento(chave_atual, nota_atual, cnpj_atual, razao_atual)
            continue
        if registro == 'D100':
            num_raw = campos[9].strip() if len(campos) > 9 else ''
            nota_atual = normalizar_numero_nota(num_raw)
            chave_atual = campos[10].strip() if len(campos) > 10 else ''
            part_cod = campos[4].strip() if len(campos) > 4 else ''
            part_info = participantes.get(part_cod, {})
            razao_atual = part_info.get('nome', '')
            cnpj_atual = part_info.get('cnpj', '')
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            registro_pai_atual = 'D100'
            doc_key_atual = gerar_chave_documento(chave_atual, nota_atual, cnpj_atual, razao_atual)
            continue
        if registro == 'C500':
            num_raw = campos[10].strip() if len(campos) > 10 else ''
            nota_atual = normalizar_numero_nota(num_raw)
            chave_atual = ''
            part_cod = campos[4].strip() if len(campos) > 4 else ''
            part_info = participantes.get(part_cod, {})
            razao_atual = part_info.get('nome', '')
            cnpj_atual = part_info.get('cnpj', '')
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            registro_pai_atual = 'C500'
            doc_key_atual = gerar_chave_documento(chave_atual, nota_atual, cnpj_atual, razao_atual)
            continue
        if registro == 'D500':
            num_raw = campos[9].strip() if len(campos) > 9 else ''
            nota_atual = normalizar_numero_nota(num_raw)
            chave_atual = ''
            part_cod = campos[4].strip() if len(campos) > 4 else ''
            part_info = participantes.get(part_cod, {})
            razao_atual = part_info.get('nome', '')
            cnpj_atual = part_info.get('cnpj', '')
            situacao_atual = campos[6].strip() if len(campos) > 6 else ''
            registro_pai_atual = 'D500'
            doc_key_atual = gerar_chave_documento(chave_atual, nota_atual, cnpj_atual, razao_atual)
            continue
        if registro not in ('C190', 'D190', 'C590', 'D590') or not nota_atual or situacao_atual == '02':
            continue

        indice_cfop = 3
        indice_valor_contabil = 5
        indice_valor_icms = 7
        indice_valor_ipi = 11 if registro in ('C190', 'C590') else None
        indice_valor_st = 9 if registro in ('C190', 'C590') else None

        cfop = campos[indice_cfop].strip() if len(campos) > indice_cfop else ''
        if not cfop:
            continue
        identificador = (doc_key_atual, cfop)
        totais[identificador]['nota'] = nota_atual
        totais[identificador]['razao_social'] = razao_atual
        totais[identificador]['cnpj'] = cnpj_atual
        totais[identificador]['cfop'] = cfop
        totais[identificador]['chave'] = chave_atual
        totais[identificador]['origem_sped'].add(registro)
        totais[identificador]['valor_contabil'] += converter_valor_decimal(campos[indice_valor_contabil] if len(campos) > indice_valor_contabil else "")
        totais[identificador]['valor_icms'] += converter_valor_decimal(campos[indice_valor_icms] if len(campos) > indice_valor_icms else "")
        if indice_valor_ipi is not None:
            totais[identificador]['valor_ipi'] += converter_valor_decimal(campos[indice_valor_ipi] if len(campos) > indice_valor_ipi else "")
        if indice_valor_st is not None:
            totais[identificador]['valor_st'] += converter_valor_decimal(campos[indice_valor_st] if len(campos) > indice_valor_st else "")

    totais_formatados = {}
    for chave, valores in sorted(totais.items()):
        valores['origem_sped'] = "/".join(sorted(valores['origem_sped']))
        totais_formatados[chave] = valores
    return totais_formatados


def montar_linhas_comparacao_por_nota(sped_notas, athena_notas):
    linhas = []
    todos = sorted(set(sped_notas.keys()) | set(athena_notas.keys()))
    for identificador in todos:
        sped = sped_notas.get(identificador, {})
        athena = athena_notas.get(identificador, {})
        nota = sped.get('nota') or athena.get('nota') or ''
        razao = sped.get('razao_social') or athena.get('razao_social') or ''
        cfop = identificador[1] if isinstance(identificador, tuple) and len(identificador) > 1 else (sped.get('cfop') or athena.get('cfop') or '')
        linha = {
            'nota': nota,
            'razao_social': razao,
            'cfop': cfop,
            'chave_sped': sped.get('chave', ''),
            'chave_athena': athena.get('chave', ''),
            'origem_sped': sped.get('origem_sped', ''),
        }
        possui_diferenca = False
        for campo in ('valor_contabil', 'valor_icms', 'valor_ipi', 'valor_st'):
            valor_sped = float(sped.get(campo, 0.0) or 0.0)
            valor_athena = float(athena.get(campo, 0.0) or 0.0)
            diferenca = valor_sped - valor_athena
            linha[f'{campo}_sped'] = valor_sped
            linha[f'{campo}_athena'] = valor_athena
            linha[f'dif_{campo}'] = diferenca
            if abs(diferenca) > 0.0001:
                possui_diferenca = True
        linha['possui_diferenca'] = possui_diferenca
        linhas.append(linha)
    return linhas


def carregar_pis_cofins_athena_por_nota_txt(caminho_arquivo):
    conteudo = ler_arquivo_texto_multiencoding(caminho_arquivo)
    linhas = [linha for linha in conteudo.splitlines() if linha.strip()]
    if len(linhas) < 2:
        return {}

    cabecalhos = [coluna.strip() for coluna in linhas[0].split('\t')]
    indices = {
        'nota': obter_indice_primeira_coluna(cabecalhos, ['NÂº NFe', 'Nº NFe', 'N° NFe', 'No NFe', 'Numero NFe', 'Numero NF', 'NFe']),
        'pis': obter_indice_primeira_coluna(cabecalhos, ['Vlr. PIS', 'Vlr PIS', 'Valor PIS']),
        'cofins': obter_indice_primeira_coluna(cabecalhos, ['Vlr. COFINS', 'Vlr COFINS', 'Valor COFINS']),
    }
    try:
        indices['chave'] = obter_indice_primeira_coluna(cabecalhos, ['Chave de Acesso', 'Chave Acesso', 'Chave'])
    except ValueError:
        indices['chave'] = None
    try:
        indices['cfop'] = obter_indice_primeira_coluna(cabecalhos, ['CFOP', 'Codigo', 'Código', 'Cód. CFOP', 'Cod. CFOP', 'Cod CFOP'])
    except ValueError:
        indices['cfop'] = None
    try:
        indices['razao'] = obter_indice_primeira_coluna(cabecalhos, ['Razão Social', 'Razao Social', 'Razão', 'Razao', 'Nome'])
    except ValueError:
        indices['razao'] = None
    try:
        indices['cnpj'] = obter_indice_primeira_coluna(cabecalhos, ['CNPJ', 'CPF', 'CNPJ/CPF'])
    except ValueError:
        indices['cnpj'] = None

    totais = defaultdict(lambda: {
        'nota': '',
        'razao_social': '',
        'cnpj': '',
        'cfop': '',
        'cfops': set(),
        'chave': '',
        'valor_pis': 0.0,
        'valor_cofins': 0.0,
    })

    for linha in linhas[1:]:
        colunas = linha.split('\t')
        nota_raw = colunas[indices['nota']].strip() if len(colunas) > indices['nota'] else ""
        if not nota_raw:
            continue
        nota = normalizar_numero_nota(nota_raw)
        chave = colunas[indices['chave']].strip() if indices['chave'] is not None and len(colunas) > indices['chave'] else ""
        razao = colunas[indices['razao']].strip() if indices['razao'] is not None and len(colunas) > indices['razao'] else ""
        cnpj = ''.join(ch for ch in (colunas[indices['cnpj']] if indices['cnpj'] is not None and len(colunas) > indices['cnpj'] else "") if ch.isdigit())

        doc_key = gerar_chave_documento(chave, nota, cnpj, razao)

        totais[doc_key]['nota'] = nota
        totais[doc_key]['razao_social'] = razao
        totais[doc_key]['cnpj'] = cnpj
        if chave:
            totais[doc_key]['chave'] = chave
        if indices['cfop'] is not None and len(colunas) > indices['cfop']:
            cfop = colunas[indices['cfop']].strip()
            if cfop:
                totais[doc_key]['cfops'].add(cfop)
        totais[doc_key]['valor_pis'] += converter_valor_decimal(colunas[indices['pis']] if len(colunas) > indices['pis'] else "")
        totais[doc_key]['valor_cofins'] += converter_valor_decimal(colunas[indices['cofins']] if len(colunas) > indices['cofins'] else "")

    for doc_key, dados_nota in totais.items():
        dados_nota['cfop'] = " / ".join(sorted(dados_nota['cfops'])) if dados_nota['cfops'] else "Sem CFOP"

    return dict(sorted(totais.items()))


def carregar_pis_cofins_c100_por_nota(dados):
    participantes = carregar_participantes_sped(dados)
    totais = {}
    doc_key_atual = None
    cfops_nota = set()

    for line in dados:
        campos = line.strip().split('|')
        if len(campos) < 2:
            continue
        registro = campos[1]
        if registro == 'C100':
            if doc_key_atual and doc_key_atual in totais:
                totais[doc_key_atual]['cfop'] = " / ".join(sorted(cfops_nota)) if cfops_nota else "Sem CFOP"
            cod_sit = campos[6].strip() if len(campos) > 6 else ""
            if cod_sit == "02":
                doc_key_atual = None
                cfops_nota = set()
                continue
            nota_raw = campos[8].strip() if len(campos) > 8 else ""
            if not nota_raw:
                doc_key_atual = None
                cfops_nota = set()
                continue
            nota = normalizar_numero_nota(nota_raw)
            chave = campos[9].strip() if len(campos) > 9 else ""
            part_cod = campos[4].strip() if len(campos) > 4 else ""
            part_info = participantes.get(part_cod, {})
            razao = part_info.get('nome', '')
            cnpj = part_info.get('cnpj', '')

            doc_key = gerar_chave_documento(chave, nota, cnpj, razao)
            doc_key_atual = doc_key
            cfops_nota = set()
            totais[doc_key] = {
                'nota': nota,
                'razao_social': razao,
                'cnpj': cnpj,
                'cfop': 'Sem CFOP',
                'chave': chave,
                'valor_pis': converter_valor_decimal(campos[26] if len(campos) > 26 else ""),
                'valor_cofins': converter_valor_decimal(campos[27] if len(campos) > 27 else ""),
            }
            continue

        if doc_key_atual and doc_key_atual in totais:
            if registro == 'C190' and len(campos) > 3:
                cfop = campos[3].strip()
                if cfop:
                    cfops_nota.add(cfop)
            elif registro == 'C170' and len(campos) > 11:
                cfop = campos[11].strip()
                if cfop:
                    cfops_nota.add(cfop)
            elif not registro.startswith('C'):
                totais[doc_key_atual]['cfop'] = " / ".join(sorted(cfops_nota)) if cfops_nota else "Sem CFOP"
                doc_key_atual = None
                cfops_nota = set()

    if doc_key_atual and doc_key_atual in totais:
        totais[doc_key_atual]['cfop'] = " / ".join(sorted(cfops_nota)) if cfops_nota else "Sem CFOP"

    return dict(sorted(totais.items()))


def montar_linhas_pis_cofins_por_nota(sped_pis_cofins, athena_pis_cofins):
    linhas = []
    todos_docs = sorted(set(sped_pis_cofins.keys()) | set(athena_pis_cofins.keys()))
    for doc_key in todos_docs:
        sped = sped_pis_cofins.get(doc_key, {})
        athena = athena_pis_cofins.get(doc_key, {})
        nota = sped.get('nota') or athena.get('nota') or ''
        razao = sped.get('razao_social') or athena.get('razao_social') or ''
        valor_pis_sped = float(sped.get('valor_pis', 0.0) or 0.0)
        valor_pis_athena = float(athena.get('valor_pis', 0.0) or 0.0)
        valor_cofins_sped = float(sped.get('valor_cofins', 0.0) or 0.0)
        valor_cofins_athena = float(athena.get('valor_cofins', 0.0) or 0.0)
        dif_pis = valor_pis_sped - valor_pis_athena
        dif_cofins = valor_cofins_sped - valor_cofins_athena
        cfop = sped.get('cfop') or athena.get('cfop') or 'Sem CFOP'
        if cfop == 'Sem CFOP' and athena.get('cfop'):
            cfop = athena.get('cfop')
        linhas.append({
            'nota': nota,
            'razao_social': razao,
            'cfop': cfop,
            'chave_sped': sped.get('chave', ''),
            'chave_athena': athena.get('chave', ''),
            'valor_pis_sped': valor_pis_sped,
            'valor_pis_athena': valor_pis_athena,
            'dif_valor_pis': dif_pis,
            'valor_cofins_sped': valor_cofins_sped,
            'valor_cofins_athena': valor_cofins_athena,
            'dif_valor_cofins': dif_cofins,
            'possui_diferenca': abs(dif_pis) > 0.0001 or abs(dif_cofins) > 0.0001,
        })
    return linhas


def agrupar_divergencias_por_cfop(linhas_por_nota):
    grupos = defaultdict(list)
    for linha in linhas_por_nota or []:
        if not linha.get('possui_diferenca'):
            continue
        cfop = str(linha.get('cfop') or 'Sem CFOP')
        grupos[cfop].append(linha)
    return [
        (cfop, sorted(linhas, key=lambda linha: (str(linha.get('nota') or ''), str(linha.get('razao_social') or ''), str(linha.get('chave_sped') or linha.get('chave_athena') or ''))))
        for cfop, linhas in sorted(grupos.items(), key=lambda item: item[0])
    ]


def montar_linhas_comparacao(sped_cfop, athena_cfop):
    linhas = []
    todos_cfops = sorted(set(sped_cfop.keys()) | set(athena_cfop.keys()))

    for cfop in todos_cfops:
        sped = sped_cfop.get(cfop, {})
        athena = athena_cfop.get(cfop, {})
        linha = {'cfop': cfop}
        for campo in ('valor_contabil', 'valor_icms', 'valor_ipi', 'valor_pis', 'valor_cofins', 'valor_st'):
            valor_sped = float(sped.get(campo, 0.0) or 0.0)
            valor_athena = float(athena.get(campo, 0.0) or 0.0)
            linha[f'{campo}_sped'] = valor_sped
            linha[f'{campo}_athena'] = valor_athena
            linha[f'dif_{campo}'] = valor_sped - valor_athena
        linhas.append(linha)

    return linhas


def exportar_comparacao_xls(caminho_arquivo, linhas_comparacao, totais_sped, totais_athena, totais_por_tipo_sped, totais_por_tipo_athena, linhas_por_nota=None, linhas_pis_cofins=None):
    colunas = [
        ('CFOP', 'cfop'),
        ('SPED Valor Contabil', 'valor_contabil_sped'),
        ('Athena Valor Contabil', 'valor_contabil_athena'),
        ('Dif Valor Contabil', 'dif_valor_contabil'),
        ('SPED ICMS', 'valor_icms_sped'),
        ('Athena ICMS', 'valor_icms_athena'),
        ('Dif ICMS', 'dif_valor_icms'),
        ('SPED IPI', 'valor_ipi_sped'),
        ('Athena IPI', 'valor_ipi_athena'),
        ('Dif IPI', 'dif_valor_ipi'),
        ('SPED PIS', 'valor_pis_sped'),
        ('Athena PIS', 'valor_pis_athena'),
        ('Dif PIS', 'dif_valor_pis'),
        ('SPED COFINS', 'valor_cofins_sped'),
        ('Athena COFINS', 'valor_cofins_athena'),
        ('Dif COFINS', 'dif_valor_cofins'),
        ('SPED Subst. Trib.', 'valor_st_sped'),
        ('Athena Subst. Trib.', 'valor_st_athena'),
        ('Dif Subst. Trib.', 'dif_valor_st'),
    ]
    def formatar_excel(valor):
        if isinstance(valor, (int, float)):
            return f"{valor:.2f}".replace('.', ',')
        return str(valor)

    def classe_diferenca(chave, valor):
        if not chave.startswith('dif_'):
            return ""
        if isinstance(valor, (int, float)) and abs(valor) > 0.0001:
            return ' class="dif"'
        return ""

    html = [
        '<html><head><meta charset="utf-8">',
        '<style>',
        'body { font-family: Calibri, Arial, sans-serif; color: #243447; padding: 18px; }',
        'h1 { color: #17324d; margin: 0 0 6px 0; font-size: 20pt; }',
        'h2 { color: #17324d; margin: 22px 0 8px 0; font-size: 13pt; }',
        '.subtitulo { color: #5b6b79; margin-bottom: 18px; font-size: 10pt; }',
        '.painel { display: inline-block; min-width: 180px; margin: 0 12px 12px 0; padding: 10px 12px; border: 1px solid #d9e2ec; background: #f8fbff; }',
        '.painel .rotulo { display: block; color: #5b6b79; font-size: 9pt; margin-bottom: 4px; }',
        '.painel .valor { display: block; color: #17324d; font-size: 12pt; font-weight: bold; }',
        'table { border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: 10pt; width: 100%; }',
        'th, td { border: 1px solid #b8c4cf; padding: 5px 7px; text-align: center; }',
        'th { background: #17324d; color: #ffffff; font-weight: bold; }',
        'tr:nth-child(even) td { background: #f8fafc; }',
        '.dif { color: #c00000; font-weight: bold; background: #fde9e7; }',
        '.resumo { font-weight: bold; }',
        '.secao { margin-top: 18px; }',
        '</style></head><body>',
        '<h1>Relatorio Comparador Athena x SPED</h1>',
        f'<div class="subtitulo">Arquivo Athena: {os.path.basename(caminho_arquivo).replace(".xls", ".txt")} | Gerado em: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</div>',
        '<div class="painel"><span class="rotulo">Total de CFOPs comparados</span><span class="valor">' + str(len(linhas_comparacao)) + '</span></div>',
        '<div class="painel"><span class="rotulo">CFOPs com diferenca</span><span class="valor">' + str(sum(1 for linha in linhas_comparacao if any(abs(linha[chave]) > 0.0001 for chave in ("dif_valor_contabil", "dif_valor_icms", "dif_valor_ipi", "dif_valor_pis", "dif_valor_cofins", "dif_valor_st")))) + '</span></div>',
        '<div class="painel"><span class="rotulo">Notas com diferenca</span><span class="valor">' + str(sum(1 for linha in (linhas_por_nota or []) if linha.get("possui_diferenca"))) + '</span></div>',
        '<div class="painel"><span class="rotulo">Notas com diferenca PIS/COFINS</span><span class="valor">' + str(sum(1 for linha in (linhas_pis_cofins or []) if linha.get("possui_diferenca"))) + '</span></div>',
        '<h2>Resumo por CFOP</h2>',
        '<table>',
        '<tr>' + ''.join(f'<th>{titulo}</th>' for titulo, _ in colunas) + '</tr>'
    ]

    for linha in linhas_comparacao:
        html.append('<tr>' + ''.join(
            f'<td{classe_diferenca(chave, linha.get(chave, ""))}>{formatar_excel(linha.get(chave, ""))}</td>'
            for _, chave in colunas
        ) + '</tr>')

    html.extend([
        '</table>',
        '<div class="secao"></div>',
        '<h2>Resumo Geral</h2>',
        '<table>',
        '<tr><th>Tipo</th><th>Base</th><th>Valor Contabil</th><th>ICMS</th><th>IPI</th><th>PIS</th><th>COFINS</th><th>Subst. Trib.</th></tr>'
    ])
    for tipo in ('Entrada', 'Saida'):
        totais_tipo_sped = totais_por_tipo_sped.get(tipo, {})
        totais_tipo_athena = totais_por_tipo_athena.get(tipo, {})
        for titulo, totais in (('SPED', totais_tipo_sped), ('Athena', totais_tipo_athena)):
            html.append(
                '<tr class="resumo">' +
                ''.join([
                    f'<td>{tipo}</td>',
                    f'<td>{titulo}</td>',
                    f'<td>{formatar_excel(totais.get("valor_contabil", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_icms", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_ipi", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_pis", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_cofins", 0.0))}</td>',
                    f'<td>{formatar_excel(totais.get("valor_st", 0.0))}</td>',
                ]) +
                '</tr>'
            )

        diferencas_resumo = {
            'valor_contabil': totais_tipo_sped.get('valor_contabil', 0.0) - totais_tipo_athena.get('valor_contabil', 0.0),
            'valor_icms': totais_tipo_sped.get('valor_icms', 0.0) - totais_tipo_athena.get('valor_icms', 0.0),
            'valor_ipi': totais_tipo_sped.get('valor_ipi', 0.0) - totais_tipo_athena.get('valor_ipi', 0.0),
            'valor_pis': totais_tipo_sped.get('valor_pis', 0.0) - totais_tipo_athena.get('valor_pis', 0.0),
            'valor_cofins': totais_tipo_sped.get('valor_cofins', 0.0) - totais_tipo_athena.get('valor_cofins', 0.0),
            'valor_st': totais_tipo_sped.get('valor_st', 0.0) - totais_tipo_athena.get('valor_st', 0.0),
        }
        html.append(
            '<tr class="resumo">' +
            ''.join([
                f'<td>{tipo}</td>',
                '<td>Diferenca (SPED - Athena)</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_contabil"])}</td>' if abs(diferencas_resumo["valor_contabil"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_contabil"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_icms"])}</td>' if abs(diferencas_resumo["valor_icms"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_icms"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_ipi"])}</td>' if abs(diferencas_resumo["valor_ipi"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_ipi"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_pis"])}</td>' if abs(diferencas_resumo["valor_pis"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_pis"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_cofins"])}</td>' if abs(diferencas_resumo["valor_cofins"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_cofins"])}</td>',
                f'<td class="dif">{formatar_excel(diferencas_resumo["valor_st"])}</td>' if abs(diferencas_resumo["valor_st"]) > 0.0001 else f'<td>{formatar_excel(diferencas_resumo["valor_st"])}</td>',
            ]) +
            '</tr>'
        )
    if linhas_por_nota:
        divergencias_por_cfop = agrupar_divergencias_por_cfop(linhas_por_nota)
        colunas_notas = [
            ('Nota Fiscal', 'nota'),
            ('Razao Social', 'razao_social'),
            ('CFOP', 'cfop'),
            ('Origem SPED', 'origem_sped'),
            ('SPED Valor Contabil', 'valor_contabil_sped'),
            ('Athena Valor Contabil', 'valor_contabil_athena'),
            ('Dif Valor Contabil', 'dif_valor_contabil'),
            ('SPED ICMS', 'valor_icms_sped'),
            ('Athena ICMS', 'valor_icms_athena'),
            ('Dif ICMS', 'dif_valor_icms'),
            ('SPED IPI', 'valor_ipi_sped'),
            ('Athena IPI', 'valor_ipi_athena'),
            ('Dif IPI', 'dif_valor_ipi'),
            ('SPED Subst. Trib.', 'valor_st_sped'),
            ('Athena Subst. Trib.', 'valor_st_athena'),
            ('Dif Subst. Trib.', 'dif_valor_st'),
            ('Chave SPED', 'chave_sped'),
            ('Chave Athena', 'chave_athena'),
        ]
        html.extend([
            '<div class="secao"></div>',
            '<h2>Notas Fiscais com Diferencas</h2>',
        ])
        for cfop, linhas_cfop in divergencias_por_cfop:
            html.extend([
                f'<h2>CFOP {cfop}</h2>',
                '<table>',
                '<tr>' + ''.join(f'<th>{titulo}</th>' for titulo, _ in colunas_notas) + '</tr>'
            ])
            for linha in linhas_cfop:
                html.append('<tr>' + ''.join(
                    f'<td{classe_diferenca(chave, linha.get(chave, ""))}>{formatar_excel(linha.get(chave, ""))}</td>'
                    for _, chave in colunas_notas
                ) + '</tr>')
            html.append('</table>')

    if linhas_pis_cofins:
        divergencias_pis_cofins_por_cfop = agrupar_divergencias_por_cfop(linhas_pis_cofins)
        colunas_pis_cofins = [
            ('Nota Fiscal', 'nota'),
            ('Razao Social', 'razao_social'),
            ('CFOP', 'cfop'),
            ('SPED PIS C100 Campo 26', 'valor_pis_sped'),
            ('Vendas PIS', 'valor_pis_athena'),
            ('Dif PIS', 'dif_valor_pis'),
            ('SPED COFINS C100 Campo 27', 'valor_cofins_sped'),
            ('Vendas COFINS', 'valor_cofins_athena'),
            ('Dif COFINS', 'dif_valor_cofins'),
            ('Chave SPED', 'chave_sped'),
            ('Chave Vendas', 'chave_athena'),
        ]
        html.extend([
            '<div class="secao"></div>',
            '<h2>Notas com Diferenca de PIS/COFINS</h2>',
        ])
        for cfop, linhas_cfop in divergencias_pis_cofins_por_cfop:
            html.extend([
                f'<h2>CFOP {cfop}</h2>',
                '<table>',
                '<tr>' + ''.join(f'<th>{titulo}</th>' for titulo, _ in colunas_pis_cofins) + '</tr>'
            ])
            for linha in linhas_cfop:
                html.append('<tr>' + ''.join(
                    f'<td{classe_diferenca(chave, linha.get(chave, ""))}>{formatar_excel(linha.get(chave, ""))}</td>'
                    for _, chave in colunas_pis_cofins
                ) + '</tr>')
            html.append('</table>')
        if not divergencias_pis_cofins_por_cfop:
            html.append('<table><tr><td colspan="11">Nenhuma diferenca encontrada em PIS/COFINS.</td></tr></table>')

    html.extend(['</body></html>'])

    with open(caminho_arquivo, 'w', encoding='utf-8-sig') as arquivo:
        arquivo.write('\n'.join(html))


def gerar_pdf_comparacao_athena_sped(linhas_comparacao, totais_sped, totais_athena, totais_por_tipo_sped, totais_por_tipo_athena, output_path, caminho_txt, linhas_por_nota=None, linhas_pis_cofins=None):
    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter))
    elements = []

    elements.append(Paragraph("Relatorio Comparador Athena x SPED", styles['Title']))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(f"Arquivo Athena: {os.path.basename(caminho_txt)}", styles['Normal']))
    elements.append(Paragraph(f"Gerado em: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
    elements.append(Spacer(1, 12))

    total_cfops_diferenca = sum(
        1 for linha in linhas_comparacao
        if any(abs(linha[chave]) > 0.0001 for chave in ('dif_valor_contabil', 'dif_valor_icms', 'dif_valor_ipi', 'dif_valor_pis', 'dif_valor_cofins', 'dif_valor_st'))
    )
    total_notas_diferenca = sum(1 for linha in (linhas_por_nota or []) if linha.get('possui_diferenca'))
    total_pis_cofins_diferenca = sum(1 for linha in (linhas_pis_cofins or []) if linha.get('possui_diferenca'))
    indicadores = [
        ['Indicadores', 'Quantidade'],
        ['CFOPs comparados', str(len(linhas_comparacao))],
        ['CFOPs com diferenca', str(total_cfops_diferenca)],
        ['Notas com diferenca', str(total_notas_diferenca)],
        ['Notas com diferenca PIS/COFINS', str(total_pis_cofins_diferenca)],
    ]
    tabela_indicadores = Table(indicadores, colWidths=[220, 90])
    tabela_indicadores.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#b8c4cf')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8fbff')),
    ]))
    elements.append(tabela_indicadores)
    elements.append(Spacer(1, 14))

    resumo = [['Tipo', 'Base', 'Valor Contabil', 'ICMS', 'IPI', 'PIS', 'COFINS', 'Subst. Trib.']]
    diferencas_resumo = []
    for tipo in ('Entrada', 'Saida'):
        totais_tipo_sped = totais_por_tipo_sped.get(tipo, {})
        totais_tipo_athena = totais_por_tipo_athena.get(tipo, {})
        resumo.append([
            tipo,
            'SPED',
            formatar_valor(totais_tipo_sped.get('valor_contabil', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_icms', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_ipi', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_pis', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_cofins', 0.0)),
            formatar_valor(totais_tipo_sped.get('valor_st', 0.0))
        ])
        resumo.append([
            tipo,
            'Athena',
            formatar_valor(totais_tipo_athena.get('valor_contabil', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_icms', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_ipi', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_pis', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_cofins', 0.0)),
            formatar_valor(totais_tipo_athena.get('valor_st', 0.0))
        ])
        diferenca_tipo = (
            totais_tipo_sped.get('valor_contabil', 0.0) - totais_tipo_athena.get('valor_contabil', 0.0),
            totais_tipo_sped.get('valor_icms', 0.0) - totais_tipo_athena.get('valor_icms', 0.0),
            totais_tipo_sped.get('valor_ipi', 0.0) - totais_tipo_athena.get('valor_ipi', 0.0),
            totais_tipo_sped.get('valor_pis', 0.0) - totais_tipo_athena.get('valor_pis', 0.0),
            totais_tipo_sped.get('valor_cofins', 0.0) - totais_tipo_athena.get('valor_cofins', 0.0),
            totais_tipo_sped.get('valor_st', 0.0) - totais_tipo_athena.get('valor_st', 0.0),
        )
        diferencas_resumo.append(diferenca_tipo)
        resumo.append([
            tipo,
            'Diferenca',
            formatar_valor(diferenca_tipo[0]),
            formatar_valor(diferenca_tipo[1]),
            formatar_valor(diferenca_tipo[2]),
            formatar_valor(diferenca_tipo[3]),
            formatar_valor(diferenca_tipo[4]),
            formatar_valor(diferenca_tipo[5])
        ])

    elements.append(Paragraph("Resumo Geral", styles['Heading2']))
    elements.append(Spacer(1, 6))
    tabela_resumo = Table(resumo, colWidths=[65, 65, 84, 70, 70, 70, 78, 84])
    estilos_resumo = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#b8c4cf')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
    ]
    for indice_linha in (3, 6):
        estilos_resumo.append(('BACKGROUND', (0, indice_linha), (-1, indice_linha), colors.HexColor('#eef4fb')))
        estilos_resumo.append(('FONTNAME', (0, indice_linha), (-1, indice_linha), 'Helvetica-Bold'))

    for indice_linha, diferenca_tipo in zip((3, 6), diferencas_resumo):
        for coluna, valor in enumerate(diferenca_tipo, start=2):
            if abs(valor) > 0.0001:
                estilos_resumo.append(('TEXTCOLOR', (coluna, indice_linha), (coluna, indice_linha), colors.red))
    tabela_resumo.setStyle(TableStyle(estilos_resumo))
    elements.append(tabela_resumo)
    elements.append(Spacer(1, 14))

    elements.append(Paragraph("Comparativo por CFOP", styles['Heading2']))
    elements.append(Spacer(1, 6))
    detalhes = [[
        'CFOP', 'SPED Contabil', 'Athena Contabil', 'Dif Contabil',
        'SPED ICMS', 'Athena ICMS', 'Dif ICMS',
        'SPED IPI', 'Athena IPI', 'Dif IPI',
        'SPED PIS', 'Athena PIS', 'Dif PIS',
        'SPED COFINS', 'Athena COFINS', 'Dif COFINS',
        'SPED ST', 'Athena ST', 'Dif ST'
    ]]
    for linha in linhas_comparacao:
        detalhes.append([
            linha['cfop'],
            formatar_valor(linha['valor_contabil_sped']),
            formatar_valor(linha['valor_contabil_athena']),
            formatar_valor(linha['dif_valor_contabil']),
            formatar_valor(linha['valor_icms_sped']),
            formatar_valor(linha['valor_icms_athena']),
            formatar_valor(linha['dif_valor_icms']),
            formatar_valor(linha['valor_ipi_sped']),
            formatar_valor(linha['valor_ipi_athena']),
            formatar_valor(linha['dif_valor_ipi']),
            formatar_valor(linha['valor_pis_sped']),
            formatar_valor(linha['valor_pis_athena']),
            formatar_valor(linha['dif_valor_pis']),
            formatar_valor(linha['valor_cofins_sped']),
            formatar_valor(linha['valor_cofins_athena']),
            formatar_valor(linha['dif_valor_cofins']),
            formatar_valor(linha['valor_st_sped']),
            formatar_valor(linha['valor_st_athena']),
            formatar_valor(linha['dif_valor_st']),
        ])

    tabela_detalhes = Table(
        detalhes,
        repeatRows=1,
        colWidths=[26, 34, 34, 34, 32, 32, 32, 32, 32, 32, 32, 32, 32, 36, 36, 36, 32, 32, 32]
    )
    estilos_detalhes = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 5.5),
        ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#b8c4cf')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
    ]
    colunas_dif = (3, 6, 9, 12, 15, 18)
    for indice_linha, linha in enumerate(linhas_comparacao, start=1):
        for coluna, chave in zip(colunas_dif, ('dif_valor_contabil', 'dif_valor_icms', 'dif_valor_ipi', 'dif_valor_pis', 'dif_valor_cofins', 'dif_valor_st')):
            if abs(linha[chave]) > 0.0001:
                estilos_detalhes.append(('TEXTCOLOR', (coluna, indice_linha), (coluna, indice_linha), colors.red))
                estilos_detalhes.append(('FONTNAME', (coluna, indice_linha), (coluna, indice_linha), 'Helvetica-Bold'))
                estilos_detalhes.append(('BACKGROUND', (coluna, indice_linha), (coluna, indice_linha), colors.HexColor('#FDE9E7')))
    tabela_detalhes.setStyle(TableStyle(estilos_detalhes))
    elements.append(tabela_detalhes)

    if linhas_por_nota:
        divergencias_por_cfop = agrupar_divergencias_por_cfop(linhas_por_nota)
        total_divergencias = sum(len(linhas_cfop) for _, linhas_cfop in divergencias_por_cfop)
        if divergencias_por_cfop:
            elements.append(PageBreak())
            elements.append(Paragraph("Notas Fiscais com Diferencas", styles['Heading2']))
            elements.append(Spacer(1, 12))
            limite_por_cfop = 40
            for cfop, linhas_cfop in divergencias_por_cfop:
                linhas_exibidas = linhas_cfop[:limite_por_cfop]
                elements.append(Paragraph(f"CFOP {cfop} ({len(linhas_cfop)} nota{'s' if len(linhas_cfop) > 1 else ''} com divergência)", styles['Heading3']))
                elements.append(Spacer(1, 6))
                tabela_notas = [[
                    'Nota', 'Razao Social', 'CFOP', 'Origem SPED', 'Dif Contabil', 'Dif ICMS', 'Dif IPI', 'Dif ST'
                ]]
                for linha in linhas_exibidas:
                    tabela_notas.append([
                        linha['nota'],
                        str(linha.get('razao_social', ''))[:26],
                        linha['cfop'],
                        linha.get('origem_sped', ''),
                        formatar_valor(linha['dif_valor_contabil']),
                        formatar_valor(linha['dif_valor_icms']),
                        formatar_valor(linha['dif_valor_ipi']),
                        formatar_valor(linha['dif_valor_st']),
                    ])

                tabela_divergencias = Table(tabela_notas, repeatRows=1, colWidths=[48, 140, 38, 52, 64, 58, 54, 54])
                estilos_notas = [
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7.5),
                    ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#b8c4cf')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
                ]
                for indice_linha, linha in enumerate(linhas_exibidas, start=1):
                    for coluna, chave in zip((4, 5, 6, 7), ('dif_valor_contabil', 'dif_valor_icms', 'dif_valor_ipi', 'dif_valor_st')):
                        if abs(linha[chave]) > 0.0001:
                            estilos_notas.append(('TEXTCOLOR', (coluna, indice_linha), (coluna, indice_linha), colors.red))
                            estilos_notas.append(('FONTNAME', (coluna, indice_linha), (coluna, indice_linha), 'Helvetica-Bold'))
                            estilos_notas.append(('BACKGROUND', (coluna, indice_linha), (coluna, indice_linha), colors.HexColor('#FDE9E7')))
                tabela_divergencias.setStyle(TableStyle(estilos_notas))
                elements.append(tabela_divergencias)
                if len(linhas_cfop) > limite_por_cfop:
                    elements.append(Spacer(1, 4))
                    elements.append(Paragraph(f"Exibindo {limite_por_cfop} de {len(linhas_cfop)} notas divergentes do CFOP {cfop} no PDF. O arquivo XLS contém todas as notas.", styles['Normal']))
                elements.append(Spacer(1, 10))

    if linhas_pis_cofins:
        divergencias_pis_cofins_por_cfop = agrupar_divergencias_por_cfop(linhas_pis_cofins)
        total_divergencias_pis_cofins = sum(len(linhas_cfop) for _, linhas_cfop in divergencias_pis_cofins_por_cfop)
        if divergencias_pis_cofins_por_cfop:
            elements.append(PageBreak())
            elements.append(Paragraph("Notas com Diferenca de PIS/COFINS", styles['Heading2']))
            elements.append(Spacer(1, 8))
            elements.append(Paragraph("Comparacao do Vlr. PIS e Vlr. COFINS do TXT de vendas contra os campos 26 e 27 do C100.", styles['Normal']))
            elements.append(Spacer(1, 8))
            limite_por_cfop = 40
            for cfop, linhas_cfop in divergencias_pis_cofins_por_cfop:
                linhas_exibidas = linhas_cfop[:limite_por_cfop]
                elements.append(Paragraph(f"CFOP {cfop} ({len(linhas_cfop)} nota{'s' if len(linhas_cfop) > 1 else ''} com divergência)", styles['Heading3']))
                elements.append(Spacer(1, 6))
                tabela_pis_cofins = [[
                    'Nota', 'Razao Social', 'CFOP', 'SPED PIS', 'Vendas PIS', 'Dif PIS',
                    'SPED COFINS', 'Vendas COFINS', 'Dif COFINS'
                ]]
                for linha in linhas_exibidas:
                    tabela_pis_cofins.append([
                        linha['nota'],
                        str(linha.get('razao_social', ''))[:26],
                        linha.get('cfop', ''),
                        formatar_valor(linha['valor_pis_sped']),
                        formatar_valor(linha['valor_pis_athena']),
                        formatar_valor(linha['dif_valor_pis']),
                        formatar_valor(linha['valor_cofins_sped']),
                        formatar_valor(linha['valor_cofins_athena']),
                        formatar_valor(linha['dif_valor_cofins']),
                    ])

                tabela_pis_cofins_pdf = Table(tabela_pis_cofins, repeatRows=1, colWidths=[46, 140, 38, 56, 56, 52, 66, 66, 56])
                estilos_pis_cofins = [
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#17324d')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                    ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#b8c4cf')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')]),
                ]
                for indice_linha, linha in enumerate(linhas_exibidas, start=1):
                    for coluna, chave in zip((5, 8), ('dif_valor_pis', 'dif_valor_cofins')):
                        if abs(linha[chave]) > 0.0001:
                            estilos_pis_cofins.append(('TEXTCOLOR', (coluna, indice_linha), (coluna, indice_linha), colors.red))
                            estilos_pis_cofins.append(('FONTNAME', (coluna, indice_linha), (coluna, indice_linha), 'Helvetica-Bold'))
                            estilos_pis_cofins.append(('BACKGROUND', (coluna, indice_linha), (coluna, indice_linha), colors.HexColor('#FDE9E7')))
                tabela_pis_cofins_pdf.setStyle(TableStyle(estilos_pis_cofins))
                elements.append(tabela_pis_cofins_pdf)
                if len(linhas_cfop) > limite_por_cfop:
                    elements.append(Spacer(1, 4))
                    elements.append(Paragraph(f"Exibindo {limite_por_cfop} de {len(linhas_cfop)} notas divergentes do CFOP {cfop} no PDF. O arquivo XLS contém todas as notas.", styles['Normal']))
                elements.append(Spacer(1, 10))

    pdf.build(elements)


import datetime
import os
import re
from pathlib import Path
import requests
import pandas as pd
# ... restante dos imports

# Configurar o locale para formato PT-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Definir estilos para os parágrafos do PDF
styles = getSampleStyleSheet()

# Conectar e criar o banco de dados
con = sqlite3.connect("referencias.db")
cur = con.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS municipios (
    codigo_cliente TEXT PRIMARY KEY,
    codigo_municipio TEXT
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS cst_substituicoes (
    fornecedor TEXT,
    cst_antigo TEXT,
    cst_novo TEXT,
    UNIQUE(fornecedor, cst_antigo)
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS cfop_conta (
    cfop TEXT PRIMARY KEY,
    conta TEXT
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS substituicoes (
    de TEXT PRIMARY KEY,
    para TEXT
)
""")

con.commit()
con.close()
print("Banco de dados 'referencias.db' criado/conectado com sucesso!")


# Funções auxiliares fora da classe
def formatar_valor(valor):
    """Formata um valor monetário no formato PT-BR."""
    try:
        return f"R$ {locale.currency(valor, grouping=True, symbol=None).replace('.', ',')}"
    except Exception:
        return f"R$ {valor}"


def converter_valor_decimal(valor):
    texto = str(valor or "").strip().replace("\xa0", "").replace(" ", "")
    if not texto:
        return 0.0
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def aplicar_pis_cofins_c100_por_cfop(dados, cfop_data):
    def garantir_cfop(cfop):
        if cfop not in cfop_data:
            cfop_data[cfop] = {
                'valor_contabil': 0.0,
                'valor_icms': 0.0,
                'valor_ipi': 0.0,
                'valor_pis': 0.0,
                'valor_cofins': 0.0,
                'valor_st': 0.0,
            }
        else:
            cfop_data[cfop].setdefault('valor_pis', 0.0)
            cfop_data[cfop].setdefault('valor_cofins', 0.0)

    def cfop_nao_gera_pis_venda(cfop):
        c = str(cfop or '').strip()
        if len(c) == 4 and c[1:] in ('901', '902', '903', '904', '908', '909', '910', '911', '912', '913', '914', '915', '916', '920', '921', '923', '924', '925', '949'):
            return True
        return False

    def fechar_c100(pis_c100, cofins_c100, c190_da_nota):
        if not c190_da_nota:
            return
        c190_tributados = [(cfop, valor) for cfop, valor in c190_da_nota if not cfop_nao_gera_pis_venda(cfop)]
        c190_destino = c190_tributados if c190_tributados else c190_da_nota
        total_contabil = sum(valor for _, valor in c190_destino)
        if abs(total_contabil) <= 0.0001:
            total_contabil = float(len(c190_destino))
            c190_destino = [(cfop, 1.0) for cfop, _ in c190_destino]
        for cfop, valor_contabil in c190_destino:
            garantir_cfop(cfop)
            proporcao = valor_contabil / total_contabil if total_contabil else 0.0
            cfop_data[cfop]['valor_pis'] += pis_c100 * proporcao
            cfop_data[cfop]['valor_cofins'] += cofins_c100 * proporcao

    pis_atual = 0.0
    cofins_atual = 0.0
    c190_atual = []
    c100_ativo = False
    c100_cancelado = False

    for line in dados:
        campos = line.strip().split('|')
        if len(campos) < 2:
            continue
        registro = campos[1]
        if registro == 'C100':
            if c100_ativo and not c100_cancelado:
                fechar_c100(pis_atual, cofins_atual, c190_atual)
            c100_ativo = True
            c100_cancelado = (campos[6].strip() if len(campos) > 6 else "") == "02"
            pis_atual = converter_valor_decimal(campos[26] if len(campos) > 26 else "")
            cofins_atual = converter_valor_decimal(campos[27] if len(campos) > 27 else "")
            c190_atual = []
            continue
        if registro == 'C190' and c100_ativo and not c100_cancelado:
            cfop = campos[3].strip() if len(campos) > 3 else ""
            if cfop:
                valor_contabil = converter_valor_decimal(campos[5] if len(campos) > 5 else "")
                c190_atual.append((cfop, valor_contabil))

    if c100_ativo and not c100_cancelado:
        fechar_c100(pis_atual, cofins_atual, c190_atual)


def somar_valores_cfop(destino, cfop, valor_contabil=0.0, valor_icms=0.0, valor_ipi=0.0, valor_pis=0.0, valor_cofins=0.0, valor_st=0.0):
    destino[cfop]['valor_contabil'] += valor_contabil
    destino[cfop]['valor_icms'] += valor_icms
    destino[cfop]['valor_ipi'] += valor_ipi
    destino[cfop]['valor_pis'] += valor_pis
    destino[cfop]['valor_cofins'] += valor_cofins
    destino[cfop]['valor_st'] += valor_st


def calcular_totais_cfop(cfop_data):
    return {
        'valor_contabil': sum(v.get('valor_contabil', 0.0) for v in cfop_data.values()),
        'valor_icms': sum(v.get('valor_icms', 0.0) for v in cfop_data.values()),
        'valor_ipi': sum(v.get('valor_ipi', 0.0) for v in cfop_data.values()),
        'valor_pis': sum(v.get('valor_pis', 0.0) for v in cfop_data.values()),
        'valor_cofins': sum(v.get('valor_cofins', 0.0) for v in cfop_data.values()),
        'valor_st': sum(v.get('valor_st', 0.0) for v in cfop_data.values())
    }


def classificar_tipo_cfop(cfop):
    try:
        cfop_num = int(str(cfop).strip())
    except (TypeError, ValueError):
        return None

    if 1000 <= cfop_num <= 3999:
        return 'Entrada'
    if 5000 <= cfop_num <= 7999:
        return 'Saida'
    return None


def calcular_totais_cfop_por_tipo(cfop_data):
    totais = {
        'Entrada': {'valor_contabil': 0.0, 'valor_icms': 0.0, 'valor_ipi': 0.0, 'valor_pis': 0.0, 'valor_cofins': 0.0, 'valor_st': 0.0},
        'Saida': {'valor_contabil': 0.0, 'valor_icms': 0.0, 'valor_ipi': 0.0, 'valor_pis': 0.0, 'valor_cofins': 0.0, 'valor_st': 0.0},
    }

    for cfop, valores in cfop_data.items():
        tipo = classificar_tipo_cfop(cfop)
        if not tipo:
            continue

        for campo in ('valor_contabil', 'valor_icms', 'valor_ipi', 'valor_pis', 'valor_cofins', 'valor_st'):
            totais[tipo][campo] += float(valores.get(campo, 0.0) or 0.0)

    return totais


BASE_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = BASE_DIR.parents[1]
TIPI_DB_CANDIDATOS = [
    WORKSPACE_DIR / "tipi.db",
    BASE_DIR / "tipi.db",
    WORKSPACE_DIR / "05 - SISTEMAS E MODULOS" / "CONSULTAS" / "tipi.db",
]
TIPI_XLSX_URL = "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/legislacao/documentos-e-arquivos/tipi.xlsx"


def normalizar_ncm_tipi(ncm):
    return "".join(ch for ch in str(ncm or "") if ch.isdigit())


def normalizar_aliquota_tipi(valor):
    texto = str(valor or "").strip().replace("%", "").replace(" ", "")
    if not texto:
        return None
    texto = texto.replace(".", "").replace(",", ".") if "," in texto else texto
    try:
        return float(texto)
    except ValueError:
        return None


def localizar_tipi_db():
    for caminho in TIPI_DB_CANDIDATOS:
        if caminho.exists():
            return caminho
    return None


def consultar_ncm_tipi(ncm):
    ncm_limpo = normalizar_ncm_tipi(ncm)
    caminho_db = localizar_tipi_db()
    if not ncm_limpo:
        return {"encontrado": False, "mensagem": "NCM nao informada.", "registro": None}
    if caminho_db is None:
        return {"encontrado": False, "mensagem": "Banco TIPI nao encontrado.", "registro": None}
    try:
        conn = sqlite3.connect(str(caminho_db))
        cursor = conn.cursor()
        cursor.execute(
            "SELECT NCM, EX, DESCRICAO, ALIQUOTA FROM Tipi WHERE REPLACE(REPLACE(NCM, '.', ''), ' ', '') = ?",
            (ncm_limpo,),
        )
        resultado = cursor.fetchone()
        conn.close()
    except Exception as e:
        return {"encontrado": False, "mensagem": f"Erro ao consultar TIPI: {e}", "registro": None}
    if not resultado:
        return {"encontrado": False, "mensagem": f"NCM {ncm_limpo} nao encontrada na TIPI.", "registro": None}
    return {
        "encontrado": True,
        "mensagem": f"NCM {resultado[0]} encontrada na TIPI.",
        "registro": {
            "NCM": resultado[0],
            "EX": resultado[1],
            "DESCRICAO": resultado[2],
            "ALIQUOTA": resultado[3],
        },
    }


def atualizar_banco_tipi_local(destino_db=None):
    caminho_db = Path(destino_db) if destino_db else (localizar_tipi_db() or WORKSPACE_DIR / "tipi.db")
    response = requests.get(TIPI_XLSX_URL, timeout=60)
    response.raise_for_status()
    caminho_xlsx = caminho_db.with_suffix(".xlsx")
    caminho_xlsx.write_bytes(response.content)
    df = pd.read_excel(caminho_xlsx, engine="openpyxl", skiprows=7)
    df.columns = [str(col).strip() for col in df.columns]

    conn = sqlite3.connect(str(caminho_db))
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS Tipi")
    cursor.execute("""
        CREATE TABLE Tipi (
            NCM TEXT,
            EX TEXT,
            DESCRICAO TEXT,
            ALIQUOTA TEXT
        )
    """)

    registros = 0
    for _, row in df.iterrows():
        ncm = str(row.get("NCM", row.get("NCM ", "")) or "").strip()
        ex = str(row.get("EX", "") or "").strip()
        descricao = str(
            row.get("DESCRIÇÃO", row.get("DESCRIÇÎÇŸO", row.get("DESCRIÇÃO ", row.get("DESCRIÇÎÇŸO ", ""))))
            or ""
        ).strip()
        aliquota = str(row.get("ALÍQUOTA (%)", row.get("ALÇ?QUOTA (%)", row.get("ALÍQUOTA", ""))) or "").strip()
        if not ncm:
            continue
        cursor.execute(
            "INSERT INTO Tipi (NCM, EX, DESCRICAO, ALIQUOTA) VALUES (?, ?, ?, ?)",
            (ncm, ex, descricao, aliquota)
        )
        registros += 1

    conn.commit()
    conn.close()
    return caminho_db, registros

def agrupar_por_cfop(registros):
    cfop_data = defaultdict(lambda: {
        'valor_contabil': 0,
        'valor_icms': 0,
        'valor_ipi': 0,
        'valor_pis': 0,
        'valor_cofins': 0,
        'valor_st': 0
    })

    for registro in registros:
        cfop = registro[3]
        
        def safe_float(value):
            return float(value.replace(',', '.')) if value else 0.0

        try:
            cfop_data[cfop]['valor_contabil'] += safe_float(registro[5])
            cfop_data[cfop]['valor_icms'] += safe_float(registro[7])
            cfop_data[cfop]['valor_ipi'] += safe_float(registro[11])
            cfop_data[cfop]['valor_st'] += safe_float(registro[9])
        except (ValueError, IndexError):
            continue

    return dict(sorted(cfop_data.items()))


# Adicione essa importação no início do seu código, se ainda não estiver lá

# E esta linha para obter os estilos padrões
styles = getSampleStyleSheet()


def gerar_pdf(cfop_data, notas_canceladas, faltantes, total_linhas, output_path, dados):
    # ========================
    # PROCESSAR D190
    # ========================
    registros_d190 = [line.strip().split('|') for line in dados if line.startswith('|D190|')]
    for reg in registros_d190:
        if len(reg) > 1:
            cfop_d190 = reg[3]  # CFOP é o campo 03
            try:
                valor_contabil = float(reg[5].replace(',', '.')) if len(reg) > 5 and reg[5] else 0
                valor_icms     = float(reg[7].replace(',', '.')) if len(reg) > 7 and reg[7] else 0

                # Garantir que a chave existe
                if cfop_d190 not in cfop_data:
                    cfop_data[cfop_d190] = {
                        'valor_contabil': 0,
                        'valor_icms': 0,
                        'valor_ipi': 0,
                        'valor_st': 0
                    }

                cfop_data[cfop_d190]['valor_contabil'] += valor_contabil
                cfop_data[cfop_d190]['valor_icms']     += valor_icms

            except (ValueError, IndexError):
                continue

    # ========================
    # PROCESSAR D590
    # ========================
    registros_d590 = [line.strip().split('|') for line in dados if line.startswith('|D590|')]
    for reg in registros_d590:
        if len(reg) > 1:
            cfop_d590 = reg[3]  # CFOP é o campo 03
            try:
                valor_contabil = float(reg[5].replace(',', '.')) if len(reg) > 5 and reg[5] else 0
                valor_icms     = float(reg[7].replace(',', '.')) if len(reg) > 7 and reg[7] else 0

                # Garantir que a chave existe
                if cfop_d590 not in cfop_data:
                    cfop_data[cfop_d590] = {
                        'valor_contabil': 0,
                        'valor_icms': 0,
                        'valor_ipi': 0,
                        'valor_st': 0
                    }

                cfop_data[cfop_d590]['valor_contabil'] += valor_contabil
                cfop_data[cfop_d590]['valor_icms']     += valor_icms

            except (ValueError, IndexError):
                continue
    # ========================
    # PROCESSAR C590
    # ========================
    registros_c590 = [line.strip().split('|') for line in dados if line.startswith('|C590|')]
    for reg in registros_c590:
        if len(reg) > 1:
            cfop_c590 = reg[3]  # CFOP é o campo 03
            try:
                valor_contabil = float(reg[5].replace(',', '.')) if len(reg) > 5 and reg[5] else 0
                valor_icms     = float(reg[7].replace(',', '.')) if len(reg) > 7 and reg[7] else 0

                # Garantir que a chave existe
                if cfop_c590 not in cfop_data:
                    cfop_data[cfop_c590] = {
                        'valor_contabil': 0,
                        'valor_icms': 0,
                        'valor_ipi': 0,
                        'valor_st': 0
                    }

                cfop_data[cfop_c590]['valor_contabil'] += valor_contabil
                cfop_data[cfop_c590]['valor_icms']     += valor_icms

            except (ValueError, IndexError):
                continue



    # ========================
    # SEPARAR ENTRADA / SAÍDA
    # ========================
    cfop_entrada = {cfop: valores for cfop, valores in cfop_data.items() if cfop.startswith(('1', '2', '3'))}
    cfop_saida   = {cfop: valores for cfop, valores in cfop_data.items() if cfop.startswith(('5', '6', '7'))}

    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []

    # Título principal
    elements.append(Paragraph("Relatório Totalizador SPED Fiscal", styles['Title']))
    elements.append(Spacer(1, 12))

    # ========================
    # ENTRADAS
    # ========================
    elements.append(Paragraph("Relatório de CFOPs de Entrada", styles['Heading2']))

    total_entrada = {
        'valor_contabil': sum(v['valor_contabil'] for v in cfop_entrada.values()),
        'valor_icms': sum(v['valor_icms'] for v in cfop_entrada.values()),
        'valor_ipi': sum(v['valor_ipi'] for v in cfop_entrada.values()),
        'valor_st': sum(v['valor_st'] for v in cfop_entrada.values())
    }

    data_entrada = [['CFOP', 'Valor Contábil', 'ICMS', 'IPI', 'Subst. Trib.']]
    for cfop, valores in sorted(cfop_entrada.items()):
        data_entrada.append([
            cfop,
            formatar_valor(valores['valor_contabil']),
            formatar_valor(valores['valor_icms']),
            formatar_valor(valores['valor_ipi']),
            formatar_valor(valores['valor_st'])
        ])
    data_entrada.append([
        'Total Entrada',
        formatar_valor(total_entrada['valor_contabil']),
        formatar_valor(total_entrada['valor_icms']),
        formatar_valor(total_entrada['valor_ipi']),
        formatar_valor(total_entrada['valor_st'])
    ])

    table_entrada = Table(data_entrada)
    table_entrada.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_entrada)
    elements.append(Spacer(1, 20))

    # ========================
    # SAÍDAS
    # ========================
    elements.append(Paragraph("Relatório de CFOPs de Saída", styles['Heading2']))

    total_saida = {
        'valor_contabil': sum(v['valor_contabil'] for v in cfop_saida.values()),
        'valor_icms': sum(v['valor_icms'] for v in cfop_saida.values()),
        'valor_ipi': sum(v['valor_ipi'] for v in cfop_saida.values()),
        'valor_st': sum(v['valor_st'] for v in cfop_saida.values())
    }

    data_saida = [['CFOP', 'Valor Contábil', 'ICMS', 'IPI', 'Subst. Trib.']]
    for cfop, valores in sorted(cfop_saida.items()):
        data_saida.append([
            cfop,
            formatar_valor(valores['valor_contabil']),
            formatar_valor(valores['valor_icms']),
            formatar_valor(valores['valor_ipi']),
            formatar_valor(valores['valor_st'])
        ])
    data_saida.append([
        'Total Saída',
        formatar_valor(total_saida['valor_contabil']),
        formatar_valor(total_saida['valor_icms']),
        formatar_valor(total_saida['valor_ipi']),
        formatar_valor(total_saida['valor_st'])
    ])

    table_saida = Table(data_saida)
    table_saida.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_saida)

    # ========================
    # RESUMO IMPOSTOS
    # ========================
    total_ipi_recolher  = total_saida['valor_ipi']  - total_entrada['valor_ipi']
    total_icms_recolher = total_saida['valor_icms'] - total_entrada['valor_icms']

    elements.append(Spacer(1, 20))
    recolher_data = [
        ['Resumo de Impostos a Recolher'],
        ['Total ICMS Saída', 'Total ICMS Entrada', 'ICMS a Recolher'],
        [formatar_valor(total_saida['valor_icms']), formatar_valor(total_entrada['valor_icms']), formatar_valor(total_icms_recolher)],
        ['Total IPI Saída', 'Total IPI Entrada', 'IPI a Recolher'],
        [formatar_valor(total_saida['valor_ipi']), formatar_valor(total_entrada['valor_ipi']), formatar_valor(total_ipi_recolher)]
    ]
    recolher_table = Table(recolher_data)
    recolher_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(recolher_table)

    # ========================
    # RESUMO FINAL
    # ========================
    elements.append(Spacer(1, 20))
    summary_data = [
        ['Resumo do Relatório'],
        [f"Total de linhas processadas: {total_linhas}"],
        [f"Total de registros CFOP: {len(cfop_data)}"],
        [f"Total de notas canceladas: {len(notas_canceladas)}"],
        [f"Total de notas faltantes: {len(faltantes)}"]
    ]
    summary_table = Table(summary_data)
    summary_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5)
    ]))
    elements.append(summary_table)

    # ========================
    # RODAPÉ
    # ========================
    data_relatorio = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    nome_arquivo = os.path.basename(output_path)
    footer_data = [[f"Data do Relatório: {data_relatorio}  |  Arquivo Gerado: {nome_arquivo}"]]
    footer_table = Table(footer_data, colWidths=[550])
    footer_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10)
    ]))
    elements.append(footer_table)

    pdf.build(elements)

    print(f"Relatório gerado com sucesso: {output_path}")

class SPEDFileEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("Editor de Arquivos SPED")
        self.root.geometry("1800x700")

        self.file_path = None
        self.data = []
        self.is_modified = False
        self.current_search_index = "1.0"
        self.last_search_text = ""
        self.frame_buttons = tk.Frame(root)
        self.frame_buttons.pack()
        
        self.current_font_size = 10 
        self.recent_files = []
        self.load_recent_config()
        self.load_layouts()
        self.setup_styles()
        self.create_widgets()
        self.create_menu()
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        self.style.configure("Treeview.Heading", font=("Calibri", 10, "bold"), background="#17324d", foreground="white")
        self.style.map("Treeview.Heading", background=[('active', '#24476b')])
        self.style.configure("Treeview", font=("Consolas", 10), rowheight=22)

    def bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.load_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_file_as())
        self.root.bind("<Control-f>", lambda e: self.focus_search_entry())
        self.root.bind("<F3>", lambda e: self.search_in_tree())

    def focus_search_entry(self):
        if hasattr(self, 'entry_search'):
            self.entry_search.focus_set()
            self.entry_search.select_range(0, tk.END)

    def mark_modified(self):
        self.is_modified = True
        self.update_window_title()

    def clear_modified(self):
        self.is_modified = False
        self.update_window_title()

    def update_window_title(self):
        title = "Editor de Arquivos SPED"
        if self.file_path:
            name = os.path.basename(self.file_path)
            title += f" - {name}"
        if getattr(self, "is_modified", False):
            title += " *"
        self.root.title(title)

    def on_close(self):
        if getattr(self, "is_modified", False):
            ans = messagebox.askyesnocancel("Salvar alterações", "Existem alterações não salvas no arquivo SPED.\nDeseja salvar antes de sair?")
            if ans is True:
                self.save_file()
                self.root.destroy()
            elif ans is False:
                self.root.destroy()
        else:
            self.root.destroy()

    def start_progress(self, msg="Processando..."):
        if hasattr(self, 'lbl_status'):
            self.lbl_status.config(text=f"Status: {msg}")
        if hasattr(self, 'progressbar'):
            self.progressbar.pack(side=tk.RIGHT, padx=10)
            self.progressbar.start(10)
        self.root.update_idletasks()

    def stop_progress(self, msg="Pronto"):
        if hasattr(self, 'progressbar'):
            self.progressbar.stop()
            self.progressbar.pack_forget()
        if hasattr(self, 'lbl_status') and msg:
            self.lbl_status.config(text=f"Status: {msg}")
        self.root.update_idletasks()

    def create_menu(self):
        """Cria e organiza a barra de menus principal com categorias limpas e padronizadas."""
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)

        # ----------------------------------------------------------------------
        # 1. MENU ARQUIVO
        # ----------------------------------------------------------------------
        file_menu = Menu(self.menubar, tearoff=0)
        file_menu.add_command(label="Abrir Arquivo SPED... (Ctrl+O)", command=self.load_file)
        
        self.recent_menu = Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Arquivos Recentes", menu=self.recent_menu)
        self.update_recent_menu()
        
        file_menu.add_separator()
        file_menu.add_command(label="Salvar Alterações (Ctrl+S)", command=self.save_file)
        file_menu.add_command(label="Salvar Como... (Ctrl+Shift+S)", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_close)
        self.menubar.add_cascade(label="Arquivo", menu=file_menu)

        # ----------------------------------------------------------------------
        # 2. MENU CADASTRO & INTEGRAÇÃO
        # ----------------------------------------------------------------------
        cadastro_menu = Menu(self.menubar, tearoff=0)
        cadastro_menu.add_command(label="Importar Tabela de Municípios (TXT)", command=self.importar_municipios_txt)
        cadastro_menu.add_command(label="Atualizar Tabela TIPI Online (Receita Federal)", command=self.atualizar_banco_tipi_menu)
        cadastro_menu.add_separator()
        cadastro_menu.add_command(label="Gerenciar Leiautes de Registros SPED...", command=self.abrir_gerenciador_layouts)
        self.menubar.add_cascade(label="Cadastros", menu=cadastro_menu)

        # ----------------------------------------------------------------------
        # 3. MENU VISUALIZAÇÃO
        # ----------------------------------------------------------------------
        view_menu = Menu(self.menubar, tearoff=0)
        view_menu.add_command(label="Visualizar em Grade Paginada (Modo Leitura)", command=self.visualizar_em_tabela)
        view_menu.add_command(label="Gerenciar Leiautes de Registros SPED...", command=self.abrir_gerenciador_layouts)
        view_menu.add_separator()
        view_menu.add_command(label="Aumentar Fonte (Zoom +)", command=self.increase_font_size)
        view_menu.add_command(label="Diminuir Fonte (Zoom -)", command=self.decrease_font_size)
        view_menu.add_command(label="Redefinir Tamanho da Fonte", command=self.reset_font_size)
        self.menubar.add_cascade(label="Exibir", menu=view_menu)

        # ----------------------------------------------------------------------
        # 4. BLOCO 0 - ABERTURA & CADASTROS
        # ----------------------------------------------------------------------
        bloco_0_menu = Menu(self.menubar, tearoff=0)
        bloco_0_menu.add_command(label="0150 / A170 - Corrigir Código do Município", command=self.corrigir_0150_e_a170)
        bloco_0_menu.add_command(label="0150 - Limpar Inscrição Municipal (Campo 6)", command=self.clear_field_6_if_condition)
        bloco_0_menu.add_command(label="0150 - Formatar IE Minas Gerais", command=self.pad_field_8_if_condition)
        bloco_0_menu.add_separator()
        bloco_0_menu.add_command(label="0200 - Corrigir NCM Vazio / Zerado '00' (Campo 8)", command=self.corrigir_0200_ncm_vazio_ou_zerado)
        bloco_0_menu.add_command(label="0200 - Atualizar Alíquota ICMS (Campo 7)", command=self.update_0200_field_7)
        bloco_0_menu.add_command(label="0200 - Preencher / Atualizar Código CEST", command=self.update_0200_cest)
        bloco_0_menu.add_command(label="0200 - Importar Itens Faltantes do H010 (Inventário)", command=self.verificar_e_importar_0200_faltantes_h010)
        bloco_0_menu.add_command(label="0220 - Processar Fator de Conversão de Unidades", command=self.process_0220_conversion)
        self.menubar.add_cascade(label="Bloco 0", menu=bloco_0_menu)

        # ----------------------------------------------------------------------
        # 5. BLOCO C - MERCADORIAS & NOTAS FISCAIS
        # ----------------------------------------------------------------------
        bloco_c_menu = Menu(self.menubar, tearoff=0)
        bloco_c_menu.add_command(label="C100 - Preencher Série Padrão (000)", command=self.fill_c100_field7)
        bloco_c_menu.add_command(label="C100 vs C190 - Validar e Ajustar Valores", command=self.verificar_c100_c190_valores)
        bloco_c_menu.add_separator()
        bloco_c_menu.add_command(label="C170 - Vincular Plano de Contas (Campo 38)", command=self.update_c170_field_38)
        bloco_c_menu.add_command(label="C170 - Alterar CST PIS / COFINS em Lote", command=self.trocar_cst_pis_cofins)
        bloco_c_menu.add_separator()
        bloco_c_menu.add_command(label="C190 - Totalizar Valores por CFOP e Alíquota", command=self.totalizar_c190)
        bloco_c_menu.add_command(label="C190 - Gerar Registros de Ajuste C195 / C197", command=self.process_c190)
        bloco_c_menu.add_command(label="C197 - Zerar Impostos nos Ajustes", command=self.process_c197_lines)
        bloco_c_menu.add_separator()
        bloco_c_menu.add_command(label="C112 - Gerar Guia de Arrecadação via Arquivo Externo", command=self.gerar_c112_externo)
        bloco_c_menu.add_command(label="C140 / C141 - Excluir Registros de Fatura", command=self.delete_c141)
        bloco_c_menu.add_separator()
        bloco_c_menu.add_command(label="Rateio - Ajustar Valores por Nota", command=self.ajustar_valores_por_nota)
        bloco_c_menu.add_command(label="Rateio - Ajustar Valores em Lote (Separado por Vírgula)", command=self.ajustar_valores_em_lote)
        self.menubar.add_cascade(label="Bloco C", menu=bloco_c_menu)

        # ----------------------------------------------------------------------
        # 6. BLOCO D - TRANSPORTES & COMUNICAÇÃO
        # ----------------------------------------------------------------------
        bloco_d_menu = Menu(self.menubar, tearoff=0)
        bloco_d_menu.add_command(label="D100 - Configurar Conta Contábil de Frete", command=self.update_d100_d101_d105_fields)
        bloco_d_menu.add_command(label="D100 - Corrigir Registros de Frete", command=self.run_validar_d100)
        bloco_d_menu.add_separator()
        bloco_d_menu.add_command(label="D101 / D105 - Recalcular Impostos PIS / COFINS", command=self.importar_alterar_sped)
        bloco_d_menu.add_command(label="D101 / D105 - Processar e Gerar D105", command=self.process_d101_d105)
        bloco_d_menu.add_command(label="D101 / D105 - Alterar PIS e COFINS por TXT (Num. Doc)", command=self.alterar_pis_cofins_d101_d105_por_txt)
        bloco_d_menu.add_command(label="D101 / D105 - Corrigir CST 56 para 66 (FOR000002709 / FOR000002427)", command=self.corrigir_frete_cst56_para_66_fornecedores)
        bloco_d_menu.add_separator()
        bloco_d_menu.add_command(label="D190 - Auditar e Corrigir Registros D190", command=self.run_validar_d190)
        self.menubar.add_cascade(label="Bloco D", menu=bloco_d_menu)

        # ----------------------------------------------------------------------
        # 7. BLOCO H - INVENTÁRIO FÍSICO
        # ----------------------------------------------------------------------
        bloco_h_menu = Menu(self.menubar, tearoff=0)
        bloco_h_menu.add_command(label="H010 - Atualizar Dados do Inventário", command=self.update_h010_inventory)
        bloco_h_menu.add_command(label="H010 - Sincronizar Unidades de Medida com 0200", command=self.sincronizar_unidade_h010_com_0200)
        bloco_h_menu.add_command(label="H010 - Padronizar Estrutura de Colunas", command=self.corrigir_h010_colunas)
        bloco_h_menu.add_separator()
        bloco_h_menu.add_command(label="H010 vs 0200 - Validar Itens / Importar Faltantes", command=self.verificar_e_importar_0200_faltantes_h010)
        self.menubar.add_cascade(label="Bloco H", menu=bloco_h_menu)

        # ----------------------------------------------------------------------
        # 8. AUDITORIA & RELATÓRIOS
        # ----------------------------------------------------------------------
        tools_menu = Menu(self.menubar, tearoff=0)
        tools_menu.add_command(label="Auditoria - Verificar Registros Duplicados", command=self.check_duplicates)
        tools_menu.add_command(label="Auditoria - Remover Duplicatas (0600)", command=self.update_field_5_correct_and_remove_duplicates)
        tools_menu.add_command(label="Auditoria - Conferir Totais (C100 vs C170)", command=self.validar_totais_c100_c170)
        tools_menu.add_command(label="Auditoria - 0200 com NCM Vazio ou '00' (Campo 8)", command=self.corrigir_0200_ncm_vazio_ou_zerado)
        tools_menu.add_command(label="Auditoria - Auditar NCM x Tabela TIPI", command=self.auditar_ncm_tipi_sped)
        tools_menu.add_command(label="Auditoria - Módulo de Pré-Validação Fiscal (PVA)", command=self.auditar_pre_validacao)
        tools_menu.add_command(label="Auditoria - Campos Obrigatórios em Branco (0200 / 0150 / C100)", command=self.auditar_campos_obrigatorios)
        tools_menu.add_command(label="Auditoria - Matriz Tributária (CFOP x CST)", command=self.auditar_matriz_tributaria)
        tools_menu.add_command(label="Auditoria - Comparador de Dois Arquivos SPED", command=self.abrir_comparador_speds)
        tools_menu.add_separator()
        tools_menu.add_command(label="Inteligência - Dashboard Executivo de Apuração", command=self.abrir_dashboard_apuracao)
        tools_menu.add_command(label="Inteligência - Gerenciador de Macros Personalizadas", command=self.abrir_gerenciador_macros)
        tools_menu.add_separator()
        tools_menu.add_command(label="Exportar Registro Selecionado para Excel (.xlsx)", command=self.exportar_registro_excel)
        tools_menu.add_command(label="Recalcular Bloco 9 (9900 / 9990 / 9999)", command=self.recalcular_bloco_9)
        tools_menu.add_command(label="Desfazer Última Alteração (Ctrl+Z)", command=self.undo_last_action)
        tools_menu.add_separator()
        tools_menu.add_command(label="Análise - Resumo C100 (Campos 25 e 26)", command=self.listar_campos_c100_25_26)
        tools_menu.add_command(label="Análise - Matriz C170 (CFOP / PIS / COFINS)", command=self.listar_campos_c170_cfop_pis_cofins)
        tools_menu.add_separator()
        tools_menu.add_command(label="Edição - Pesquisar e Substituir Ocorrências", command=self.pesquisar_e_substituir)
        tools_menu.add_command(label="Edição - Remover Espaços Extras (Limpar Pipes)", command=self.remove_spaces_between_pipes)
        tools_menu.add_command(label="Edição - Contar Colunas do Registro", command=self.contar_colunas)
        tools_menu.add_command(label="Edição - Aplicar Alterações Pendentes", command=self.apply_changes)
        tools_menu.add_separator()
        tools_menu.add_command(label="Relatório - Totalizador Simples (PDF)", command=self.gerar_relatorio_totalizador_simples)
        tools_menu.add_command(label="Relatório - Comparador Athena x SPED Fiscal (PDF / XLS)", command=self.gerar_relatorio_totalizador)
        tools_menu.add_command(label="Relatório - Comparador Athena x SPED Contribuições (PDF / XLS)", command=self.gerar_relatorio_totalizador_contribuicoes)
        self.menubar.add_cascade(label="Ferramentas", menu=tools_menu)

        # ----------------------------------------------------------------------
        # 9. MENU AJUDA & MANUAIS
        # ----------------------------------------------------------------------
        help_menu = Menu(self.menubar, tearoff=0)
        help_menu.add_command(label="Manual Completo de Uso (F1)", command=self.abrir_manual_uso)
        help_menu.add_command(label="Tabela de Atalhos de Teclado", command=self.abrir_atalhos_teclado)
        help_menu.add_separator()
        help_menu.add_command(label="Gerenciar Leiautes de Registros SPED...", command=self.abrir_gerenciador_layouts)
        help_menu.add_separator()
        help_menu.add_command(label="Desenvolvido por Wesley Rocha Raimundo", state=tk.DISABLED)
        self.menubar.add_cascade(label="Ajuda", menu=help_menu)

    def create_widgets(self):
        # --- TOPO (Busca) ---
        frame_top = tk.Frame(self.root)
        frame_top.pack(fill=tk.X, padx=10, pady=5)
        
        # Busca de Texto Simples
        self.entry_search = tk.Entry(frame_top)
        self.entry_search.pack(side=tk.LEFT, padx=3, fill=tk.X, expand=False)
        btn_search = tk.Button(frame_top, text="Localizar Texto", command=self.search_in_tree)
        btn_search.pack(side=tk.LEFT, padx=5)

        # Busca de Linha
        tk.Label(frame_top, text="Linha:").pack(side=tk.LEFT, padx=5)
        self.entry_search_line = tk.Entry(frame_top, width=10)
        self.entry_search_line.pack(side=tk.LEFT, padx=5)
        btn_search_line = tk.Button(frame_top, text="Ir", command=self.goto_line_tree)
        btn_search_line.pack(side=tk.LEFT, padx=5)

        # === NOVO BOTÃO DE BUSCA DE NOTA ===
        btn_nota = tk.Button(frame_top, text="Buscar Nota e Destacar", bg="#13BFD6", command=self.find_and_highlight_invoice)
        btn_nota.pack(side=tk.LEFT, padx=20)
        # ===================================

                # --- BARRA DE REGISTRO / FILTRO RÁPIDO ---
        tk.Label(frame_top, text="|  Registro:").pack(side=tk.LEFT, padx=3)
        self.combo_registro_filter = ttk.Combobox(
            frame_top, 
            values=["TODOS", "0000", "0150", "0200", "C100", "C170", "C190", "C197", "D100", "D190", "E110", "H010", "1010"], 
            width=8, 
            state="readonly"
        )
        self.combo_registro_filter.set("TODOS")
        self.combo_registro_filter.pack(side=tk.LEFT, padx=3)
        self.combo_registro_filter.bind("<<ComboboxSelected>>", lambda e: self.display_data())

        # --- BARRA DE STATUS COM PROGRESSBAR ---
        self.frame_status = tk.Frame(self.root, bg="#f0f4f8")
        self.frame_status.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=2)

        self.lbl_status = tk.Label(self.frame_status, text="Status: Nenhum arquivo carregado", anchor="w", bg="#f0f4f8", font=("Calibri", 9))
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.progressbar = ttk.Progressbar(self.frame_status, mode='indeterminate', length=140)

        # --- CENTRO (A Tabela Principal) ---
        frame_middle = tk.Frame(self.root)
        frame_middle.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree_scroll_y = ttk.Scrollbar(frame_middle, orient="vertical")
        self.tree_scroll_x = ttk.Scrollbar(frame_middle, orient="horizontal")

        self.tree = ttk.Treeview(
            frame_middle, 
            yscrollcommand=self.tree_scroll_y.set, 
            xscrollcommand=self.tree_scroll_x.set,
            selectmode="browse",
            show="headings"
        )
        
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.abrir_editor_linha)

        # --- RODAPÉ (Edição) ---
        frame_bottom = tk.Frame(self.root)
        frame_bottom.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_bottom, text="Linha:").pack(side=tk.LEFT, padx=5)
        self.entry_line = tk.Entry(frame_bottom, width=8)
        self.entry_line.pack(side=tk.LEFT, padx=5)

        tk.Label(frame_bottom, text="Coluna:").pack(side=tk.LEFT, padx=5)
        self.entry_column = tk.Entry(frame_bottom, width=5)
        self.entry_column.pack(side=tk.LEFT, padx=5)

        tk.Label(frame_bottom, text="Conteúdo:").pack(side=tk.LEFT, padx=5)
        self.entry_value = tk.Entry(frame_bottom, width=40)
        self.entry_value.pack(side=tk.LEFT, padx=5)

        btn_edit = tk.Button(frame_bottom, text="Salvar Edição", command=self.edit_value_tree)
        btn_edit.pack(side=tk.LEFT, padx=5)

        btn_delete = tk.Button(frame_bottom, text="Deletar Linha", command=self.delete_line)
        btn_delete.pack(side=tk.LEFT, padx=5)
        
        # Filtros rápidos
        tk.Label(frame_top, text="|  Filtro Rápido (Col):").pack(side=tk.LEFT, padx=5)
        self.entry_filter_column = tk.Entry(frame_top, width=5)
        self.entry_filter_column.pack(side=tk.LEFT, padx=5)
        tk.Label(frame_top, text="Valor:").pack(side=tk.LEFT, padx=2)
        self.entry_filter_value = tk.Entry(frame_top, width=15)
        self.entry_filter_value.pack(side=tk.LEFT, padx=5)
        btn_filter = tk.Button(frame_top, text="Filtrar", command=self.filter_data_tree)
        btn_filter.pack(side=tk.LEFT, padx=5)
        btn_reset = tk.Button(frame_top, text="Limpar", command=self.display_data)
        btn_reset.pack(side=tk.LEFT, padx=5)
    def find_and_highlight_invoice(self):
        """Busca uma nota fiscal (C100) e destaca ela e todos os seus filhos (itens/impostos) em amarelo."""
        target_num = simpledialog.askstring("Buscar Nota", "Digite o número da Nota Fiscal:")
        if not target_num:
            return
            
        target_num = target_num.strip()
        found_any = False
        first_found_id = None
        
        # 1. Limpar destaques anteriores (Restaura o padrão zebrado)
        # Vamos redefinir as tags de todos os itens visíveis
        for i, item_id in enumerate(self.tree.get_children()):
            # Recalcula se é par ou impar baseado na posição atual
            tag = 'even' if i % 2 == 0 else 'odd'
            self.tree.item(item_id, tags=(tag,))
            
        # 2. Configurar a cor de destaque (Amarelo)
        self.tree.tag_configure('highlight_nota', background='#FFFFAA', foreground='black') # Amarelo claro
        
        # 3. Lógica de Busca e Marcação
        # Vamos varrer os dados brutos para entender a hierarquia (quem pertence à nota)
        
        ids_para_destacar = []
        inside_target_note = False
        
        for i, line in enumerate(self.data):
            parts = line.strip().split('|')
            if len(parts) < 2: continue
            
            reg = parts[1]
            
            if reg == 'C100':
                # Verifica se é a nota alvo (Campo 8 é o número)
                if len(parts) > 8 and parts[8].strip() == target_num:
                    inside_target_note = True
                    found_any = True
                else:
                    # Se encontrou OUTRO C100, saiu da nota alvo
                    inside_target_note = False
            
            # Se estivermos dentro do escopo da nota alvo, marca a linha
            if inside_target_note:
                # O ID na Treeview é o número da linha (str(i+1))
                item_id = str(i + 1)
                
                # Verifica se a linha existe na tabela (pode estar filtrada/oculta)
                if self.tree.exists(item_id):
                    ids_para_destacar.append(item_id)
                    if first_found_id is None:
                        first_found_id = item_id

        # 4. Aplica o destaque visual
        if ids_para_destacar:
            for item_id in ids_para_destacar:
                self.tree.item(item_id, tags=('highlight_nota',))
            
            # Rola a tela até o C100 encontrado
            self.tree.see(first_found_id)
            self.tree.selection_set(first_found_id)
            messagebox.showinfo("Sucesso", f"Nota {target_num} encontrada!\n{len(ids_para_destacar)} registros destacados.")
        else:
            messagebox.showwarning("Não encontrado", f"Nota {target_num} não encontrada.")
            
    def abrir_editor_linha(self, event):
        """Abre uma janela pop-up para editar todos os campos da linha selecionada."""
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item_id = selected_item[0]
        valores = self.tree.item(item_id)['values']
        linha_num = int(valores[0])
        raw_line = self.data[linha_num - 1]
        partes = raw_line.strip().split('|')
        dados_uteis = partes[1:-1] if len(partes) > 1 else partes

        registro = dados_uteis[0] if len(dados_uteis) > 0 else "Dados"
        headers = self.layouts.get(registro, [])

        editor_win = tk.Toplevel(self.root)
        editor_win.title(f"Editor de Registro: Linha {linha_num} ({registro})")
        editor_win.geometry("500x600")

        canvas = tk.Canvas(editor_win)
        scrollbar = tk.Scrollbar(editor_win, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, pady=(0, 60))
        scrollbar.pack(side="right", fill="y")

        entries_list = []
        entries_por_header = {}
        cor_normal = "white"
        cor_alerta = "#fff1f2"
        cor_ok = "#ecfdf3"

        tk.Label(scrollable_frame, text=f"Editando Linha {linha_num}", font=("Arial", 12, "bold")).pack(pady=10)
        status_tipi = tk.Label(scrollable_frame, text="Validacao TIPI: pronta.", font=("Arial", 9, "bold"), fg="#475467", anchor="w", justify="left")
        status_tipi.pack(fill=tk.X, padx=10, pady=(0, 8))
        status_check_ncm = tk.Label(scrollable_frame, text="CHECK NCM: aguardando validacao.", font=("Arial", 9, "bold"), fg="#475467", anchor="w", justify="left")
        status_check_ncm.pack(fill=tk.X, padx=10, pady=(0, 8))

        frame_info_tipi = tk.LabelFrame(scrollable_frame, text="TIPI / NCM", font=("Arial", 9, "bold"))
        frame_info_tipi.pack(fill=tk.X, padx=10, pady=(0, 10))
        lbl_info_tipi = tk.Label(frame_info_tipi, text="Sem consulta realizada.", justify="left", anchor="w", font=("Arial", 9), wraplength=430)
        lbl_info_tipi.pack(fill=tk.X, padx=8, pady=8)

        for i, valor_atual in enumerate(dados_uteis):
            nome_campo = f"{i+1}. {headers[i]}" if i < len(headers) else f"Campo {i+1}"
            row_frame = tk.Frame(scrollable_frame)
            row_frame.pack(fill=tk.X, padx=10, pady=2)
            tk.Label(row_frame, text=nome_campo, width=25, anchor="w", font=("Arial", 9)).pack(side=tk.LEFT)
            ent = tk.Entry(row_frame, width=40)
            ent.insert(0, valor_atual)
            ent.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            if i == 0:
                ent.config(state='disabled')
            entries_list.append(ent)
            if i < len(headers):
                entries_por_header[headers[i]] = ent

        def focar_proximo_campo(event):
            try:
                indice_atual = entries_list.index(event.widget)
            except ValueError:
                return "break"

            for proximo_widget in entries_list[indice_atual + 1:]:
                if str(proximo_widget.cget("state")) != "disabled":
                    proximo_widget.focus_set()
                    proximo_widget.icursor(tk.END)
                    return "break"

            return "break"

        for ent in entries_list:
            if str(ent.cget("state")) != "disabled":
                ent.bind("<Return>", focar_proximo_campo)

        def obter_valor_header(nome_header):
            widget = entries_por_header.get(nome_header)
            return "" if not widget else widget.get().strip()

        def buscar_ncm_do_item(cod_item):
            cod_item = str(cod_item or "").strip()
            if not cod_item:
                return None
            for indice, linha in enumerate(self.data, start=1):
                partes_local = linha.strip().split('|')
                dados_local = partes_local[1:-1] if len(partes_local) > 1 else partes_local
                if len(dados_local) > 7 and dados_local[0] == "0200" and dados_local[1].strip() == cod_item:
                    return {"cod_item": cod_item, "ncm": dados_local[7].strip(), "linha": indice}
            return None

        def aplicar_cor(widget, tipo="normal"):
            if str(widget.cget("state")) == "disabled":
                return
            widget.config(bg=cor_alerta if tipo == "erro" else cor_ok if tipo == "ok" else cor_normal)

        def validar_tipi_editor():
            for widget in entries_por_header.values():
                aplicar_cor(widget, "normal")

            if registro == "0200":
                ncm = obter_valor_header("COD_NCM")
                resultado_tipi = consultar_ncm_tipi(ncm)
                widget_ncm = entries_por_header.get("COD_NCM")
                if widget_ncm:
                    aplicar_cor(widget_ncm, "ok" if resultado_tipi["encontrado"] else "erro")
                status_tipi.config(text=f"Validacao TIPI: {resultado_tipi['mensagem']}", fg="#027a48" if resultado_tipi["encontrado"] else "#b42318")
                if resultado_tipi["encontrado"]:
                    status_check_ncm.config(text="CHECK NCM: OK - NCM certa na TIPI.", fg="#027a48")
                    reg_tipi = resultado_tipi["registro"]
                    lbl_info_tipi.config(text=f"NCM: {reg_tipi['NCM']}\nEX: {reg_tipi['EX'] or '-'}\nDescricao: {reg_tipi['DESCRICAO']}", fg="#101828")
                else:
                    status_check_ncm.config(text="CHECK NCM: ERRO - NCM nao localizada na TIPI.", fg="#b42318")
                    lbl_info_tipi.config(text=resultado_tipi["mensagem"], fg="#b42318")
                return

            if registro == "C170":
                cod_item = obter_valor_header("COD_ITEM")
                aliq_ipi_digitada = obter_valor_header("ALIQ_IPI")
                widget_cod_item = entries_por_header.get("COD_ITEM")
                widget_aliq_ipi = entries_por_header.get("ALIQ_IPI")
                item_0200 = buscar_ncm_do_item(cod_item)
                if not item_0200:
                    if widget_cod_item:
                        aplicar_cor(widget_cod_item, "erro")
                    status_tipi.config(text=f"Validacao TIPI: COD_ITEM {cod_item or '(vazio)'} nao encontrado no registro 0200.", fg="#b42318")
                    status_check_ncm.config(text="CHECK NCM: ERRO - item sem vinculo com 0200.", fg="#b42318")
                    lbl_info_tipi.config(text=f"Nao foi possivel localizar o item {cod_item or '(vazio)'} no registro 0200 para obter a NCM.", fg="#b42318")
                    return

                resultado_tipi = consultar_ncm_tipi(item_0200["ncm"])
                if widget_cod_item:
                    aplicar_cor(widget_cod_item, "ok" if resultado_tipi["encontrado"] else "erro")
                if not resultado_tipi["encontrado"]:
                    status_tipi.config(text=f"Validacao TIPI: item {cod_item} vinculado a NCM {item_0200['ncm']} nao encontrada.", fg="#b42318")
                    status_check_ncm.config(text="CHECK NCM: ERRO - NCM vinculada nao encontrada na TIPI.", fg="#b42318")
                    lbl_info_tipi.config(text=f"COD_ITEM: {cod_item}\nNCM vinculada no 0200: {item_0200['ncm']}\n{resultado_tipi['mensagem']}", fg="#b42318")
                    return

                aliq_tipi = normalizar_aliquota_tipi(resultado_tipi["registro"]["ALIQUOTA"])
                aliq_c170 = normalizar_aliquota_tipi(aliq_ipi_digitada)
                aliquota_confere = aliq_tipi is not None and aliq_c170 is not None and abs(aliq_tipi - aliq_c170) < 0.0001
                if widget_aliq_ipi:
                    aplicar_cor(widget_aliq_ipi, "ok" if aliquota_confere else "erro")
                if aliquota_confere:
                    status_check_ncm.config(text="CHECK NCM: OK - NCM certa e aliquota IPI compativel.", fg="#027a48")
                    status_tipi.config(text=f"Validacao TIPI: item {cod_item} com NCM {item_0200['ncm']} valida e ALIQ_IPI compativel ({aliq_ipi_digitada}).", fg="#027a48")
                else:
                    status_check_ncm.config(text="CHECK NCM: PARCIAL - NCM certa, mas aliquota divergente.", fg="#b54708")
                    status_tipi.config(text=f"Validacao TIPI: item {cod_item} usa NCM {item_0200['ncm']}. TIPI={resultado_tipi['registro']['ALIQUOTA']}% e C170={aliq_ipi_digitada or '-'}.", fg="#b42318")
                reg_tipi = resultado_tipi["registro"]
                lbl_info_tipi.config(
                    text=f"COD_ITEM: {cod_item}\nNCM no 0200: {item_0200['ncm']}\nEX: {reg_tipi['EX'] or '-'}\nAliquota TIPI: {reg_tipi['ALIQUOTA']}%\nAliquota no C170: {aliq_ipi_digitada or '-'}\nDescricao: {reg_tipi['DESCRICAO']}",
                    fg="#101828" if aliquota_confere else "#b42318"
                )
                return

            status_tipi.config(text="Validacao TIPI: disponivel para registros 0200 e C170.", fg="#475467")
            status_check_ncm.config(text="CHECK NCM: disponivel para registros 0200 e C170.", fg="#475467")
            lbl_info_tipi.config(text="Esse painel mostra a descricao da NCM e a aliquota IPI retornadas da TIPI para registros 0200 e C170.", fg="#475467")

        if registro in ("0200", "C170"):
            tk.Button(scrollable_frame, text="Verificar NCM / TIPI", bg="#1d4ed8", fg="white", font=("Arial", 9, "bold"), command=validar_tipi_editor).pack(pady=(0, 10))

        for chave_header in ("COD_NCM", "COD_ITEM", "ALIQ_IPI"):
            widget = entries_por_header.get(chave_header)
            if widget:
                widget.bind("<KeyRelease>", lambda e: validar_tipi_editor())
                widget.bind("<FocusOut>", lambda e: validar_tipi_editor())

        validar_tipi_editor()

        def salvar_edicao():
            novos_valores = [ent.get() for ent in entries_list]
            nova_linha_str = "|" + "|".join(novos_valores) + "|\n"
            self.data[linha_num - 1] = nova_linha_str
            valores_tree = [linha_num] + novos_valores
            colunas_atuais = self.tree["columns"]
            if len(valores_tree) < len(colunas_atuais):
                valores_tree += [""] * (len(colunas_atuais) - len(valores_tree))
            self.tree.item(item_id, values=valores_tree)
            messagebox.showinfo("Sucesso", "Registro atualizado!")
            editor_win.destroy()

        btn_frame = tk.Frame(editor_win, bg="#ddd")
        btn_frame.place(relx=0, rely=0.9, relwidth=1, relheight=0.1)
        tk.Button(btn_frame, text="SALVAR ALTERACOES", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), command=salvar_edicao).pack(pady=15)

        def _on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except tk.TclError:
                pass

        def _bind_mouse(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mouse(event):
            canvas.unbind_all("<MouseWheel>")

        editor_win.bind("<Enter>", _bind_mouse)
        editor_win.bind("<Leave>", _unbind_mouse)
        editor_win.bind("<Destroy>", _unbind_mouse)

    def atualizar_banco_tipi_menu(self):
        try:
            self.lbl_status.config(text="Atualizando banco TIPI...")
            self.root.update_idletasks()
            caminho_db, total = atualizar_banco_tipi_local()
            self.lbl_status.config(text=f"Banco TIPI atualizado: {caminho_db}")
            messagebox.showinfo("Sucesso", f"Banco TIPI atualizado com sucesso.\n\nArquivo: {caminho_db}\nRegistros importados: {total}")
        except Exception as e:
            self.lbl_status.config(text="Falha ao atualizar banco TIPI.")
            messagebox.showerror("Erro", f"Nao foi possivel atualizar o banco TIPI:\n{e}")

    def _mapear_itens_0200(self):
        itens = {}
        for idx, linha in enumerate(self.data, start=1):
            partes = linha.strip().split('|')
            dados = partes[1:-1] if len(partes) > 1 else partes
            if len(dados) > 7 and dados[0] == "0200":
                cod_item = dados[1].strip()
                itens[cod_item] = {
                    "linha": idx,
                    "cod_item": cod_item,
                    "ncm": dados[7].strip(),
                    "aliq_icms": dados[11].strip() if len(dados) > 11 else "",
                    "cest": dados[12].strip() if len(dados) > 12 else "",
                }
        return itens

    def auditar_ncm_tipi_sped(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Carregue um arquivo SPED antes de auditar as NCMs.")
            return

        for item_id in self.tree.get_children():
            valores_linha = self.tree.item(item_id, "values")
            try:
                numero_linha = int(valores_linha[0])
            except (ValueError, IndexError, TypeError):
                numero_linha = 0
            tag_base = 'even' if numero_linha % 2 == 1 else 'odd'
            self.tree.item(item_id, tags=(tag_base,))

        self.tree.tag_configure('audit_ok', background='#ecfdf3', foreground='black')
        self.tree.tag_configure('audit_erro', background='#fff1f2', foreground='black')
        self.tree.tag_configure('audit_divergente', background='#fff7ed', foreground='black')

        itens_0200 = self._mapear_itens_0200()
        relatorio = []
        totais = {"0200_ok": 0, "0200_erro": 0, "c170_ok": 0, "c170_divergente": 0, "c170_sem_item": 0, "c170_ncm_nao_encontrada": 0}

        for cod_item, info in sorted(itens_0200.items()):
            consulta = consultar_ncm_tipi(info["ncm"])
            if consulta["encontrado"]:
                totais["0200_ok"] += 1
                if self.tree.exists(str(info["linha"])):
                    self.tree.item(str(info["linha"]), tags=('audit_ok',))
                relatorio.append({"tipo": "0200", "linha": info["linha"], "referencia": cod_item, "ncm": info["ncm"], "status": "OK", "detalhe": "NCM encontrada na TIPI."})
            else:
                totais["0200_erro"] += 1
                if self.tree.exists(str(info["linha"])):
                    self.tree.item(str(info["linha"]), tags=('audit_erro',))
                relatorio.append({"tipo": "0200", "linha": info["linha"], "referencia": cod_item, "ncm": info["ncm"], "status": "ERRO", "detalhe": consulta["mensagem"]})

        for idx, linha in enumerate(self.data, start=1):
            partes = linha.strip().split('|')
            dados = partes[1:-1] if len(partes) > 1 else partes
            if len(dados) <= 22 or dados[0] != "C170":
                continue
            cod_item = dados[2].strip()
            aliq_ipi_c170 = dados[22].strip()
            item_0200 = itens_0200.get(cod_item)

            if not item_0200:
                totais["c170_sem_item"] += 1
                if self.tree.exists(str(idx)):
                    self.tree.item(str(idx), tags=('audit_erro',))
                relatorio.append({"tipo": "C170", "linha": idx, "referencia": cod_item, "ncm": "", "status": "ERRO", "detalhe": "COD_ITEM nao encontrado no 0200."})
                continue

            consulta = consultar_ncm_tipi(item_0200["ncm"])
            if not consulta["encontrado"]:
                totais["c170_ncm_nao_encontrada"] += 1
                if self.tree.exists(str(idx)):
                    self.tree.item(str(idx), tags=('audit_erro',))
                relatorio.append({"tipo": "C170", "linha": idx, "referencia": cod_item, "ncm": item_0200["ncm"], "status": "ERRO", "detalhe": consulta["mensagem"]})
                continue

            aliq_tipi = normalizar_aliquota_tipi(consulta["registro"]["ALIQUOTA"])
            aliq_c170 = normalizar_aliquota_tipi(aliq_ipi_c170)
            if aliq_tipi is not None and aliq_c170 is not None and abs(aliq_tipi - aliq_c170) < 0.0001:
                totais["c170_ok"] += 1
                if self.tree.exists(str(idx)):
                    self.tree.item(str(idx), tags=('audit_ok',))
                relatorio.append({"tipo": "C170", "linha": idx, "referencia": cod_item, "ncm": item_0200["ncm"], "status": "OK", "detalhe": f"ALIQ_IPI compativel com TIPI ({consulta['registro']['ALIQUOTA']}%)."})
            else:
                totais["c170_divergente"] += 1
                if self.tree.exists(str(idx)):
                    self.tree.item(str(idx), tags=('audit_divergente',))
                relatorio.append({"tipo": "C170", "linha": idx, "referencia": cod_item, "ncm": item_0200["ncm"], "status": "DIVERGENTE", "detalhe": f"TIPI={consulta['registro']['ALIQUOTA']}% / C170={aliq_ipi_c170 or '-'}"})

        relatorio_win = tk.Toplevel(self.root)
        relatorio_win.title("Relatorio de Auditoria NCM / TIPI")
        relatorio_win.geometry("1100x650")
        resumo = (
            f"0200 OK: {totais['0200_ok']}   |   0200 com erro: {totais['0200_erro']}   |   "
            f"C170 OK: {totais['c170_ok']}   |   C170 divergente: {totais['c170_divergente']}   |   "
            f"C170 sem 0200: {totais['c170_sem_item']}   |   C170 NCM nao encontrada: {totais['c170_ncm_nao_encontrada']}"
        )
        tk.Label(relatorio_win, text="Auditoria de NCM / TIPI", font=("Arial", 11, "bold")).pack(pady=(8, 4))
        tk.Label(relatorio_win, text=resumo, font=("Arial", 9), wraplength=1060, justify="left").pack(padx=10, pady=(0, 8))

        frame_table = tk.Frame(relatorio_win)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        columns = ("tipo", "linha", "referencia", "ncm", "status", "detalhe")
        tree = ttk.Treeview(frame_table, columns=columns, show="headings")
        for col, titulo, largura in [("tipo", "Registro", 90), ("linha", "Linha", 70), ("referencia", "Referencia", 180), ("ncm", "NCM", 110), ("status", "Status", 110), ("detalhe", "Detalhe", 500)]:
            tree.heading(col, text=titulo)
            tree.column(col, width=largura, anchor="w" if col in ("referencia", "detalhe") else "center")
        scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree.tag_configure("ok", background="#ecfdf3")
        tree.tag_configure("erro", background="#fff1f2")
        tree.tag_configure("divergente", background="#fff7ed")

        for item in relatorio:
            tag = "ok" if item["status"] == "OK" else "divergente" if item["status"] == "DIVERGENTE" else "erro"
            tree.insert("", tk.END, values=(item["tipo"], item["linha"], item["referencia"], item["ncm"], item["status"], item["detalhe"]), tags=(tag,))

        def exportar_relatorio_ncm():
            save_path = filedialog.asksaveasfilename(title="Exportar Relatorio NCM / TIPI", defaultextension=".txt", filetypes=[("Arquivo Texto", "*.txt")])
            if not save_path:
                return
            try:
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write("RELATORIO AUDITORIA NCM / TIPI\n")
                    f.write(resumo + "\n\n")
                    for item in relatorio:
                        f.write(f"{item['tipo']} | Linha {item['linha']} | Ref {item['referencia']} | NCM {item['ncm'] or '-'} | {item['status']} | {item['detalhe']}\n")
                messagebox.showinfo("Sucesso", f"Relatorio exportado para:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar relatorio:\n{e}")

        tk.Button(relatorio_win, text="Exportar Relatorio TXT", command=exportar_relatorio_ncm, height=2).pack(fill=tk.X, padx=20, pady=10)

        primeira_linha_problema = next((item["linha"] for item in relatorio if item["status"] in ("ERRO", "DIVERGENTE")), None)
        if primeira_linha_problema and self.tree.exists(str(primeira_linha_problema)):
            self.tree.see(str(primeira_linha_problema))
            self.tree.selection_set(str(primeira_linha_problema))


    def visualizar_em_tabela(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        # 1. Preparação da Janela
        tabela_win = tk.Toplevel(self.root)
        tabela_win.title("Visualização em Grade (Modo Leitura)")
        tabela_win.geometry("1200x600")

        # Frame de Filtro no Topo
        frame_top = tk.Frame(tabela_win)
        frame_top.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(frame_top, text="Filtrar por Registro (ex: C100):").pack(side=tk.LEFT)
        entry_filtro = tk.Entry(frame_top)
        entry_filtro.pack(side=tk.LEFT, padx=5)

        # 2. Análise das Colunas
        # Descobre qual é o número máximo de colunas no arquivo para criar o cabeçalho
        max_colunas = 0
        dados_processados = []
        
        for i, line in enumerate(self.data):
            # Removemos quebras de linha e separamos por pipe
            # O SPED começa e termina com pipe, então o split gera strings vazias nas pontas
            # Vamos limpar isso para ficar bonito na tabela
            partes = line.strip().split('|')
            
            # Remove o primeiro e último item se forem vazios (comum no sped |C100|...|)
            if len(partes) > 0 and partes[0] == '': partes.pop(0)
            if len(partes) > 0 and partes[-1] == '': partes.pop(-1)
            
            if len(partes) > max_colunas:
                max_colunas = len(partes)
            
            # Guardamos o índice original, o registro e os dados
            dados_processados.append((i + 1, partes))

        # 3. Configuração da Treeview (Tabela)
        frame_table = tk.Frame(tabela_win)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # As colunas serão: Linha + C0, C1, C2... até o máximo encontrado
        colunas = ["Linha"] + [f"C{k+1}" for k in range(max_colunas)]
        
        tree = ttk.Treeview(frame_table, columns=colunas, show="headings")
        
        # Configurar Cabeçalhos
        tree.heading("Linha", text="Linha")
        tree.column("Linha", width=60, anchor="center")
        
        for col in colunas[1:]:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="w")

        # Scrollbars (Horizontal e Vertical)
        vsb = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame_table, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 4. Inserção de Dados
        def preencher_tabela(filtro=None):
            # Limpa tudo
            for item in tree.get_children():
                tree.delete(item)
            
            # Insere linhas
            for idx_linha, dados in dados_processados:
                # Se tiver filtro, verifica se o registro (primeira coluna de dados) bate
                if filtro:
                    if len(dados) > 0 and filtro.upper() not in dados[0].upper():
                        continue

                # Prepara os valores. Se a linha for menor que o max de colunas, completa com vazio
                valores = [idx_linha] + dados
                if len(valores) < len(colunas):
                    valores += [""] * (len(colunas) - len(valores))
                
                tree.insert("", "end", values=valores)

        # Carrega dados iniciais
        preencher_tabela()

        # Botão de Filtrar
        def aplicar_filtro():
            texto = entry_filtro.get().strip()
            preencher_tabela(texto if texto else None)

        btn_filtrar = tk.Button(frame_top, text="Filtrar", command=aplicar_filtro)
        btn_filtrar.pack(side=tk.LEFT, padx=5)

        btn_limpar = tk.Button(frame_top, text="Limpar Filtro", command=lambda: [entry_filtro.delete(0, tk.END), preencher_tabela()])
        btn_limpar.pack(side=tk.LEFT, padx=5)

        # Dica: Duplo clique para ver linha no editor principal (Opcional, mas útil)
        def on_double_click(event):
            item = tree.selection()[0]
            valores = tree.item(item, "values")
            linha_num = valores[0] # A primeira coluna é o número da linha
            
            # Fecha tabela e foca no editor
            tabela_win.destroy()
            
            # Navega no editor principal
            self.entry_search_line.delete(0, tk.END)
            self.entry_search_line.insert(0, linha_num)
            self.search_by_line()

        tree.bind("<Double-1>", on_double_click)

    def gerar_c112_externo(self):
        # 1. Selecionar o arquivo externo
        path_externo = filedialog.askopenfilename(
            title="Selecione o arquivo 'Cons-DBGrid'",
            filetypes=[("Texto", "*.txt"), ("CSV", "*.csv"), ("Todos", "*.*")]
        )
        if not path_externo:
            return

        # 2. Configurações de Leitura
        separador = simpledialog.askstring("Configuração", "Qual o separador do arquivo?\n(Deixe vazio para usar | )", initialvalue="|")
        if not separador: separador = "|"
        if separador == "\\t": separador = "\t"

        # --- COLUNAS FIXADAS ---
        # Ordem: Nota, Cod Arrec, UF, Num Doc, Auth, Valor, Dt Venc, Dt Pagto
        # Índices: 1, 0, 5, 4, 11, 9, 8, 9
        indices = [1, 0, 5, 4, 11, 9, 8, 8]

        idx_nota, idx_cod, idx_uf, idx_num_doc, idx_auth, idx_vlr, idx_dt_vct, idx_dt_pgt = indices

        # 3. Ler arquivo externo e carregar em memória
        dados_externos = {}
        try:
            with open(path_externo, "r", encoding="latin-1") as f:
                for linha in f:
                    if not linha.strip(): continue
                    partes = linha.strip().split(separador)
                    
                    if len(partes) > 11:
                        numero_nota = partes[idx_nota].strip()
                        # Remove zeros à esquerda
                        chave_nota = str(int(numero_nota)) if numero_nota.isdigit() else numero_nota
                        
                        dados_externos[chave_nota] = {
                            'cod': partes[idx_cod].strip(),
                            'uf': partes[idx_uf].strip(),
                            'num_doc': partes[idx_num_doc].strip(),
                            'auth': partes[idx_auth].strip(),
                            'vlr': partes[idx_vlr].strip(),
                            'dt_vct': partes[idx_dt_vct].strip(),
                            'dt_pgt': partes[idx_dt_pgt].strip()
                        }
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler arquivo externo: {e}")
            return

        # 4. Processar o SPED
        novas_linhas = []
        c100_nota_atual = None
        c112_inserido_para_esta_nota = False
        insercoes_count = 0

        i = 0
        while i < len(self.data):
            linha = self.data[i]
            campos = linha.strip().split('|')
            
            # Se for C100
            if len(campos) > 8 and campos[1] == 'C100':
                ind_oper = campos[2] # Campo 2: 0=Entrada, 1=Saída
                
                if ind_oper == '1':
                    c100_nota_atual = campos[8].strip()
                    if c100_nota_atual.isdigit():
                        c100_nota_atual = str(int(c100_nota_atual))
                    c112_inserido_para_esta_nota = False
                else:
                    c100_nota_atual = None 
                
                novas_linhas.append(linha)
            
            # === CORREÇÃO AQUI: Adicionado verificação 'len(campos) > 1' ===
            elif len(campos) > 1 and campos[1] == 'C190' and c100_nota_atual and not c112_inserido_para_esta_nota:
                if c100_nota_atual in dados_externos:
                    d = dados_externos[c100_nota_atual]
                    
                    valor_formatado = d['vlr'].replace('.', ',')
                    dt_vct_fmt = d['dt_vct'].replace('/', '').replace('-', '')
                    dt_pgt_fmt = d['dt_pgt'].replace('/', '').replace('-', '')

                    # Monta o registro C112
                    registro_c112 = f"|C112|{d['cod']}|{d['uf']}|{d['num_doc']}|{d['auth']}|{valor_formatado}|{dt_vct_fmt}|{dt_pgt_fmt}|\n"
                    
                    novas_linhas.append(registro_c112)
                    insercoes_count += 1
                    c112_inserido_para_esta_nota = True 
                
                novas_linhas.append(linha)
            
            else:
                novas_linhas.append(linha)
            
            i += 1

        # 5. Atualizar e Finalizar
        self.data = novas_linhas
        self.display_data()
        
        if insercoes_count > 0:
            messagebox.showinfo("Sucesso", f"Processamento concluído!\nForam inseridos {insercoes_count} registros C112 em notas de SAÍDA.")
        else:
            messagebox.showwarning("Aviso", "Nenhum registro inserido.\nVerifique se as notas são de Saída (1) e se os números coincidem.")
    ## MODIFICAÇÃO: Funções para controle de tamanho da fonte (Zoom)
    def load_recent_config(self):
        """Carrega a lista de arquivos recentes de um JSON local."""
        try:
            if os.path.exists("recent_files.json"):
                with open("recent_files.json", "r") as f:
                    self.recent_files = json.load(f)
            else:
                self.recent_files = []
        except Exception:
            self.recent_files = []

    def save_recent_config(self):
        """Salva a lista atual no JSON."""
        try:
            with open("recent_files.json", "w") as f:
                json.dump(self.recent_files, f)
        except Exception as e:
            print(f"Erro ao salvar config: {e}")

    def update_recent_menu(self):
        """Reconstroi o submenu de arquivos recentes."""
        self.recent_menu.delete(0, tk.END) # Limpa
        if not self.recent_files:
            self.recent_menu.add_command(label="(Vazio)", state=tk.DISABLED)
        else:
            for path in self.recent_files:
                # Cria uma função lambda para capturar o caminho específico
                self.recent_menu.add_command(
                    label=path, 
                    command=lambda p=path: self.load_file(p)
                )

    def add_to_recent(self, file_path):
        """Adiciona um arquivo ao topo da lista e mantém apenas 5."""
        # Remove se já existir para mover pro topo
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        
        self.recent_files.insert(0, file_path)
        
        # Mantém só os 5 primeiros
        if len(self.recent_files) > 5:
            self.recent_files = self.recent_files[:5]
            
        self.save_recent_config()
        self.update_recent_menu()

    def increase_font_size(self):
        """Aumenta o tamanho da fonte da tabela em tempo real."""
        if self.current_font_size < 28: 
            self.current_font_size += 1
            self._apply_new_font_size()

    def decrease_font_size(self):
        """Diminui o tamanho da fonte da tabela em tempo real."""
        if self.current_font_size > 7: 
            self.current_font_size -= 1
            self._apply_new_font_size()

    def reset_font_size(self):
        """Redefine o tamanho da fonte da tabela para o padrão (10)."""
        self.current_font_size = 10
        self._apply_new_font_size()
        
    def _apply_new_font_size(self):
        """Aplica a nova configuração de fonte e altura de linha na tabela (Treeview) em tempo real."""
        row_height = max(18, int(self.current_font_size * 2.2))
        self.style.configure("Treeview", font=("Consolas", self.current_font_size), rowheight=row_height)
        self.style.configure("Treeview.Heading", font=("Calibri", max(9, self.current_font_size), "bold"))
        if hasattr(self, 'text_display'):
            try:
                self.text_display.config(font=("Consolas", self.current_font_size))
            except Exception:
                pass
        if hasattr(self, 'lbl_status'):
            self.lbl_status.config(text=f"Status: Zoom da tabela ajustado para {self.current_font_size}pt")
        
    # FIM DAS MODIFICAÇÕES DE ZOOM
    
    def fill_c100_field7(self):
        count = 0
        for i in range(len(self.data)):
            fields = self.data[i].strip().split("|")
            if len(fields) > 7 and (fields[1] == "C100" or fields[0] == "C100") and fields[7] == "":
                fields[7] = "000"
                self.data[i] = "|".join(fields) + "\n"
                count += 1
        self.display_data()
        self.is_modified = True
        self.update_window_title()
        messagebox.showinfo("Sucesso", f"Total de {count} registro(s) C100 com Série preenchida (000).")

    def search_by_line(self):
        try:
            line_num = int(self.entry_search_line.get().strip())
            if line_num < 1 or line_num > len(self.data):
                raise ValueError("Número de linha inválido.")
            start_idx = f"{line_num}.0"
            self.text_display.see(start_idx)
            self.text_display.tag_remove("highlight", "1.0", tk.END)
            end_idx = f"{line_num}.end"
            self.text_display.tag_add("highlight", start_idx, end_idx)
            self.text_display.tag_config("highlight", background="lightblue", foreground="black")
            messagebox.showinfo("Sucesso", f"Você está na linha {line_num}.")
        except ValueError as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    def ajustar_valores_em_lote(self):
        # 1. Solicitar os números das notas
        entrada_notas = simpledialog.askstring(
            "Ajuste em Lote", 
            "Digite os números das Notas Fiscais separados por vírgula:\nEx: 100, 105, 230"
        )
        if not entrada_notas:
            return

        # Limpar e criar lista de notas alvo
        # O split(',') separa pela vírgula e o strip() remove espaços em branco extras
        notas_alvo = [n.strip() for n in entrada_notas.split(',') if n.strip()]
        
        if not notas_alvo:
            return

        notas_encontradas = set()
        total_linhas_alteradas = 0
        
        # Função auxiliar de rateio (Interna)
        def aplicar_rateio_lote(indices, indice_campo_valor, valor_alvo_total):
            if not indices: return 0
            
            soma_atual = 0.0
            valores_originais = []
            
            # Coleta valores atuais
            for idx in indices:
                linha = self.data[idx].strip().split('|')
                try:
                    val = float(linha[indice_campo_valor].replace(',', '.'))
                except (ValueError, IndexError):
                    val = 0.0
                valores_originais.append(val)
                soma_atual += val

            if soma_atual == 0: soma_atual = 1 

            soma_novos_valores = 0.0
            modificacoes = 0

            # Aplica regra de 3 em todos, exceto o último
            for k, idx in enumerate(indices[:-1]):
                linha = self.data[idx].strip().split('|')
                peso = valores_originais[k] / soma_atual
                novo_valor = valor_alvo_total * peso
                
                linha[indice_campo_valor] = f"{novo_valor:.2f}".replace('.', ',')
                self.data[idx] = "|".join(linha) + "\n"
                
                soma_novos_valores += novo_valor
                modificacoes += 1

            # Ajuste de centavos no último item
            ultimo_idx = indices[-1]
            linha_ultima = self.data[ultimo_idx].strip().split('|')
            valor_restante = valor_alvo_total - soma_novos_valores
            if valor_restante < 0: valor_restante = 0 
            
            linha_ultima[indice_campo_valor] = f"{valor_restante:.2f}".replace('.', ',')
            self.data[ultimo_idx] = "|".join(linha_ultima) + "\n"
            modificacoes += 1
            
            return modificacoes

        # 2. Percorrer o arquivo UMA vez procurando as notas da lista
        i = 0
        while i < len(self.data):
            line = self.data[i]
            campos = line.strip().split('|')
            
            # Verifica se é um C100 e se o número (campo 8) está na nossa lista de alvos
            if len(campos) > 8 and campos[1] == 'C100' and campos[8] in notas_alvo:
                numero_nota = campos[8]
                notas_encontradas.add(numero_nota)
                
                try:
                    # Campo 12 = Valor Total do Documento
                    valor_total_nota = float(campos[12].replace(',', '.'))
                    
                    # Identificar filhos (C170 e C190)
                    indices_c170 = []
                    indices_c190 = []
                    
                    # Loop interno para capturar os filhos desta nota específica
                    j = i + 1
                    while j < len(self.data):
                        sub_campos = self.data[j].strip().split('|')
                        if len(sub_campos) < 2: 
                            j += 1
                            continue
                            
                        tipo_reg = sub_campos[1]
                        
                        if tipo_reg == 'C100': # Acabou a nota atual
                            break
                        
                        if tipo_reg == 'C170':
                            indices_c170.append(j)
                        elif tipo_reg == 'C190':
                            indices_c190.append(j)
                        
                        j += 1
                    
                    # Aplicar correções
                    if indices_c170:
                        total_linhas_alteradas += aplicar_rateio_lote(indices_c170, 7, valor_total_nota)
                    
                    if indices_c190:
                        total_linhas_alteradas += aplicar_rateio_lote(indices_c190, 5, valor_total_nota) # Valor Operação
                        total_linhas_alteradas += aplicar_rateio_lote(indices_c190, 10, valor_total_nota) # Valor Total
                
                except ValueError:
                    print(f"Erro ao processar valores da nota {numero_nota}")
            
            i += 1

        # 3. Relatório Final
        self.display_data()
        
        notas_nao_encontradas = set(notas_alvo) - notas_encontradas
        
        msg = f"Processamento concluído!\n\n"
        msg += f"Notas ajustadas: {len(notas_encontradas)}\n"
        msg += f"Total de registros alterados: {total_linhas_alteradas}\n"
        
        if notas_nao_encontradas:
            msg += f"\nATENÇÃO: As seguintes notas não foram encontradas:\n"
            msg += ", ".join(notas_nao_encontradas)
            messagebox.showwarning("Relatório Parcial", msg)
        else:
            messagebox.showinfo("Sucesso Total", msg)
            
    def show_changes_and_confirm(self, changes, action_label):
        """
        Mostra uma tela de pré-visualização das alterações e pede confirmação.
        :param changes: Uma lista de tuplas (linha_original, linha_alterada)
        :param action_label: O nome da ação a ser confirmada (ex: "Corrigir D190")
        """
        if not changes:
            messagebox.showinfo("Nenhuma Alteração", f"Nenhuma alteração encontrada para a ação '{action_label}'.")
            return

        confirm_window = tk.Toplevel(self.root)
        confirm_window.title(f"Pré-visualização e Confirmação: {action_label}")
        confirm_window.geometry("1200x600")

        frame = tk.Frame(confirm_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Usando um Text widget para exibir as mudanças
        text_display = scrolledtext.ScrolledText(frame, wrap=tk.NONE, undo=False, height=25)
        text_display.pack(fill=tk.BOTH, expand=True)
        
        # Inserindo as linhas originais e alteradas no widget de texto
        text_display.insert(tk.END, f"Alterações para a ação: '{action_label}'\n\n", "header")
        
        for original_line, new_line in changes:
            text_display.insert(tk.END, "--------------------------------------------------------\n\n")
            text_display.insert(tk.END, "Linha Original:\n", "original_tag")
            text_display.insert(tk.END, f"{original_line}\n")
            text_display.insert(tk.END, "Linha Corrigida:\n", "new_tag")
            text_display.insert(tk.END, f"{new_line}\n\n")

        # Configurando as tags para cores
        text_display.tag_config("header", foreground="blue", font=("Helvetica", 12, "bold"))
        text_display.tag_config("original_tag", foreground="red", font=("Helvetica", 10, "bold"))
        text_display.tag_config("new_tag", foreground="green", font=("Helvetica", 10, "bold"))
        
        # Desabilita a edição do texto
        text_display.config(state=tk.DISABLED)

        # Botões de confirmação e cancelamento
        button_frame = tk.Frame(confirm_window)
        button_frame.pack(pady=10)

        def apply_changes():
            # Itera sobre as alterações e aplica no self.data
            for i in range(len(self.data)):
                # Encontra a linha original para a qual há uma alteração
                original_line_to_find = changes[0][0].strip()
                if self.data[i].strip() == original_line_to_find:
                    self.data[i] = changes[0][1].strip() + "\n"
                    changes.pop(0) # Remove a alteração já aplicada
                    if not changes:
                        break # Sai do loop se todas as alterações foram aplicadas
            
            self.display_data()
            messagebox.showinfo("Concluído", f"Total de {len(changes)} alteração(ões) aplicada(s) com sucesso para a ação '{action_label}'!")
            confirm_window.destroy()

        def cancel_changes():
            messagebox.showinfo("Cancelado", "As alterações foram canceladas.")
            confirm_window.destroy()

        btn_confirm = tk.Button(button_frame, text="Aplicar Alterações", command=apply_changes)
        btn_confirm.pack(side=tk.LEFT, padx=10)

        btn_cancel = tk.Button(button_frame, text="Cancelar", command=cancel_changes)
        btn_cancel.pack(side=tk.LEFT, padx=10)

    def importar_alterar_sped(self):
        skip_to_d101 = {"F01720", "F06588", "F05755", "F01209", "21914"}
        def ask_cst():
            cst_window = tk.Toplevel(self.root)
            cst_window.title("Informe o CST")
            tk.Label(cst_window, text="Digite o CST para substituir os campos 4 dos registros D101 e D105:").pack(padx=10, pady=10)
            cst_entry = tk.Entry(cst_window)
            cst_entry.pack(padx=10, pady=5)
            cst_entry.focus_set()
            cst_value = {"value": None}
            def submit_cst():
                valor = cst_entry.get().strip()
                if not valor:
                    messagebox.showerror("Erro", "O valor do CST não pode ser vazio.")
                    return
                cst_value["value"] = valor
                cst_window.destroy()
            tk.Button(cst_window, text="Confirmar", command=submit_cst).pack(padx=10, pady=10)
            self.root.wait_window(cst_window)
            return cst_value["value"]

        novo_cst = ask_cst()
        if not novo_cst:
            messagebox.showwarning("Aviso", "Processamento cancelado. CST não informado.")
            return
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 4:
                if fields[1] == "D100":
                    if fields[4] in skip_to_d101:
                        return "|".join(fields) + "\n"
                if fields[1] in ["D101", "D105"]:
                    if len(fields) > 8:
                        fields[4] = novo_cst
                        if fields[4] != "98":
                            if fields[1] == "D101":
                                fields[7] = "1,2375"
                            elif fields[1] == "D105":
                                fields[7] = "5,70"
                            try:
                                valor6 = float(fields[6].replace(",", "."))
                                valor7 = float(fields[7].replace(",", ".")) / 100
                                fields[8] = f"{(valor6 * valor7):.2f}".replace(".", ",")
                            except ValueError:
                                fields[8] = "0"
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Regras aplicadas nos registros D100, D101 e D105.")

    def update_field_5_correct_and_remove_duplicates(self):
        mapping = {
            "03.10": "Gerencia Industrial", "03.11": "Foles e Malhas", "03.12": "Tubos Flexiveis",
            "03.13": "Suporte de Mola", "03.14": "PTFE", "03.15": "Junta de Borracha",
            "03.16": "Usinagem", "03.17": "Caldeiraria", "03.19": "Acabamento e Pintura",
            "03.20": "Carpintaria e Expedi o", "03.21": "Manutencao e Ferramentaria",
            "03.22": "Obras Civis", "03.23": "CGQ Controle de Qualidade", "03.24": "PCP",
        }
        seen = set()
        def correct_field_4(field):
            if len(field) == 4 and field.isdigit():
                return f"{field[:2]}.{field[2:]}"
            return field
        def modify_and_filter_lines(line):
            fields = line.strip().split("|")
            if len(fields) > 5 and fields[1] == "0600":
                fields[3] = correct_field_4(fields[3])
                if fields[3] in mapping:
                    fields[4] = mapping[fields[3]]
                duplicate_key = (fields[1], fields[3])
                if duplicate_key in seen:
                    return None
                seen.add(duplicate_key)
            return "|".join(fields) + "\n"
        self.data = [line for line in map(modify_and_filter_lines, self.data) if line is not None]
        self.display_data()
        messagebox.showinfo("Sucesso", "Campo 5 atualizado, campo 4 corrigido e duplicatas removidas.")

    def validar_d190(self):
        """
        Coleta as alterações nos registros D190 e as retorna para pré-visualização.
        """
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return []

        changes = []
        for i, line in enumerate(self.data):
            campos = line.strip().split('|')
            if len(campos) > 7 and campos[1] == 'D190':
                try:
                    valor_campo7 = float(campos[7].replace(',', '.'))
                    if valor_campo7 > 0 and campos[6] != campos[5]:
                        original_line = line.strip()
                        campos[6] = campos[5] 
                        new_line = "|".join(campos)
                        changes.append((original_line, new_line))
                except (ValueError, IndexError):
                    pass
        
        return changes

    def run_validar_d190(self):
        """
        Ponto de entrada para a ação de correção do D190 com pré-visualização.
        """
        changes = self.validar_d190()
        self.show_changes_and_confirm(changes, "Corrigir D190")

    def update_c170_field_38(self):
        mapping_c3 = {
            "1": "111011", "2": "111012", "3": "111013", "7": "31121", "R": "31114",
            "0": "212499", "AN": "212499",
        }
        mapping_c12 = {
            ("5101", "5124", "5401"): "31111",
            ("6101", "6124", "6401"): "31112",
            ("7101", "7124"): "31113",
            ("5102", "6102", "6403", "6108"): "31114",
            ("1202"): "31121",
        }
        def get_mapping_c12(value):
            for keys, result in mapping_c12.items():
                if value in keys:
                    return result
            return None
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 38 and fields[1] == "C170":
                if fields[3] and fields[3][0] in mapping_c3:
                    fields[37] = mapping_c3[fields[3][0]]
                elif len(fields[3]) > 1 and fields[3][:2] in mapping_c3:
                    fields[37] = mapping_c3[fields[3][:2]]
                mapped_value = get_mapping_c12(fields[11])
                if mapped_value:
                    fields[37] = mapped_value
            return "|".join(fields) + "\n"
        def modify_line2(line):
            fields = line.strip().split("|")
            if len(fields) > 16:
                if fields[1] == 'A170' and fields[3][:2] in ['J0', 'W0', 'J3', 'A0', 'R0', 'B0', 'S0', 'R3']:
                    fields[17] = '31115'
                if fields[1] == 'A170' and fields[3][:2] in ['S2', '62', '60', '63', '64', '65', '29', '61', '69', '20', '21', '22', '23', '24', '25', '26', '27', '28', '30', '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53', '54', '55', '56', '57', '58', '59', '00']:
                    fields[17] = '41330'
                if fields[1] == 'C170' and fields[3][:2] in ['90']:
                    fields[37] = '212499'
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.data = [modify_line2(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Campo 38 atualizado para as condições do registro C170.")

    def verificar_c100_c190_valores(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Carregue um arquivo SPED antes de verificar.")
            return

        # Lista para armazenar as divergências encontradas
        divergencias = []

        # Vamos percorrer as linhas e agrupar
        c100_atual = None
        c100_linha_idx = -1
        c100_campos = []
        c190_indices = []

        def processar_nota_anterior():
            if c100_atual is not None and c190_indices:
                # Obter valor de VL_DOC (campo 12)
                try:
                    vl_doc_str = c100_campos[12].replace(',', '.') if len(c100_campos) > 12 else '0'
                    vl_doc = float(vl_doc_str) if vl_doc_str.strip() else 0.0
                except ValueError:
                    vl_doc = 0.0

                # Obter soma de VL_OPR (campo 5) dos registros C190 associados
                vl_opr_soma = 0.0
                for idx in c190_indices:
                    c190_campos = self.data[idx].strip().split('|')
                    try:
                        vl_opr_str = c190_campos[5].replace(',', '.') if len(c190_campos) > 5 else '0'
                        vl_opr = float(vl_opr_str) if vl_opr_str.strip() else 0.0
                    except ValueError:
                        vl_opr = 0.0
                    vl_opr_soma += vl_opr

                # Calcular diferença
                diferenca = vl_doc - vl_opr_soma
                if abs(diferenca) > 0.005:  # Diferença maior que meio centavo
                    divergencias.append({
                        'linha_c100': c100_linha_idx + 1,
                        'nota': c100_campos[8] if len(c100_campos) > 8 else '',
                        'chave': c100_campos[9] if len(c100_campos) > 9 else '',
                        'vl_doc': vl_doc,
                        'vl_opr_soma': vl_opr_soma,
                        'diferenca': diferenca,
                        'c190_indices': list(c190_indices)
                    })

        for i, line in enumerate(self.data):
            campos = line.strip().split('|')
            if len(campos) < 2:
                continue
            
            registro = campos[1]

            if registro == 'C100':
                # Processa a nota anterior antes de iniciar a nova
                processar_nota_anterior()
                
                # Zera/reinicia controle
                c100_linha_idx = i
                c100_campos = campos
                c190_indices = []
                
                # Situação da nota (campo 6)
                cod_sit = campos[6] if len(campos) > 6 else ''
                if cod_sit == '02':  # Cancelada
                    c100_atual = None
                else:
                    c100_atual = campos[8] if len(campos) > 8 else ''
            
            elif registro == 'C190':
                if c100_atual is not None:
                    c190_indices.append(i)

        # Processar a última nota
        processar_nota_anterior()

        if not divergencias:
            messagebox.showinfo("Auditoria Concluída", "Nenhuma divergência de valores foi encontrada entre os registros C100 e C190.")
            return

        # Se houver divergências, mostramos na tela em uma janela Toplevel
        janela_dif = tk.Toplevel(self.root)
        janela_dif.title("Divergências C190 vs C100")
        janela_dif.geometry("900x500")
        janela_dif.transient(self.root)
        janela_dif.grab_set()

        lbl_info = tk.Label(janela_dif, text=f"Foram encontradas {len(divergencias)} notas com diferenças entre o C100 (VL_DOC) e o C190 (Soma de VL_OPR).", font=("Arial", 10, "bold"), fg="#b42318", pady=10)
        lbl_info.pack()

        # Frame da Tabela
        frame_tabela = tk.Frame(janela_dif)
        frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        colunas = ("linha", "nota", "chave", "vl_doc", "vl_opr_soma", "diferenca", "c190_qtde")
        tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings")

        tree.heading("linha", text="Linha C100")
        tree.heading("nota", text="Nota")
        tree.heading("chave", text="Chave de Acesso")
        tree.heading("vl_doc", text="VL_DOC (C100)")
        tree.heading("vl_opr_soma", text="Soma VL_OPR (C190)")
        tree.heading("diferenca", text="Diferença")
        tree.heading("c190_qtde", text="Qtd C190")

        tree.column("linha", width=80, anchor="center")
        tree.column("nota", width=100, anchor="center")
        tree.column("chave", width=250, anchor="w")
        tree.column("vl_doc", width=110, anchor="e")
        tree.column("vl_opr_soma", width=110, anchor="e")
        tree.column("diferenca", width=100, anchor="e")
        tree.column("c190_qtde", width=80, anchor="center")

        scrollbar = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for item in divergencias:
            tree.insert("", tk.END, values=(
                item['linha_c100'],
                item['nota'],
                item['chave'],
                f"{item['vl_doc']:.2f}",
                f"{item['vl_opr_soma']:.2f}",
                f"{item['diferenca']:.2f}",
                len(item['c190_indices'])
            ))

        def executar_correcao():
            # Executa a atualização
            alteracoes_count = 0
            for item in divergencias:
                vl_total_c100 = item['vl_doc']
                indices = item['c190_indices']
                
                if not indices:
                    continue

                if len(indices) == 1:
                    # Apenas um C190: o valor do C190 Campo 5 = C100 Campo 12
                    idx = indices[0]
                    campos_c190 = self.data[idx].strip().split('|')
                    campos_c190[5] = f"{vl_total_c100:.2f}".replace('.', ',')
                    self.data[idx] = "|".join(campos_c190) + "\n"
                    alteracoes_count += 1
                else:
                    # Múltiplos C190: rateio proporcional ao valor original
                    # Coletar valores originais
                    soma_atual = 0.0
                    valores_originais = []
                    for idx in indices:
                        campos_c190 = self.data[idx].strip().split('|')
                        try:
                            val = float(campos_c190[5].replace(',', '.'))
                        except (ValueError, IndexError):
                            val = 0.0
                        valores_originais.append(val)
                        soma_atual += val

                    if soma_atual == 0.0:
                        # Se todos forem zero, rateia igualmente
                        peso_igual = 1.0 / len(indices)
                        valores_originais = [peso_igual] * len(indices)
                        soma_atual = 1.0

                    soma_novos_valores = 0.0
                    for k, idx in enumerate(indices[:-1]):
                        campos_c190 = self.data[idx].strip().split('|')
                        peso = valores_originais[k] / soma_atual
                        novo_valor = round(vl_total_c100 * peso, 2)
                        campos_c190[5] = f"{novo_valor:.2f}".replace('.', ',')
                        self.data[idx] = "|".join(campos_c190) + "\n"
                        soma_novos_valores += novo_valor
                        alteracoes_count += 1

                    # Último registro fica com o resto
                    ultimo_idx = indices[-1]
                    campos_c190_ultimo = self.data[ultimo_idx].strip().split('|')
                    valor_restante = round(vl_total_c100 - soma_novos_valores, 2)
                    if valor_restante < 0.0:
                        valor_restante = 0.0
                    campos_c190_ultimo[5] = f"{valor_restante:.2f}".replace('.', ',')
                    self.data[ultimo_idx] = "|".join(campos_c190_ultimo) + "\n"
                    alteracoes_count += 1

            self.display_data()
            messagebox.showinfo("Sucesso", f"Atualização concluída!\n\nForam corrigidos {len(divergencias)} documentos (C100), totalizando {alteracoes_count} registros C190 alterados.")
            janela_dif.destroy()

        def cancelar():
            janela_dif.destroy()

        # Frame de Botões
        frame_botoes = tk.Frame(janela_dif)
        frame_botoes.pack(pady=15)

        btn_atualizar = tk.Button(frame_botoes, text="Atualizar C190 de acordo com C100", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), padx=10, pady=5, command=executar_correcao)
        btn_atualizar.pack(side=tk.LEFT, padx=10)

        btn_cancelar = tk.Button(frame_botoes, text="Cancelar", bg="#f44336", fg="white", font=("Arial", 10, "bold"), padx=10, pady=5, command=cancelar)
        btn_cancelar.pack(side=tk.LEFT, padx=10)

    def trocar_cst_pis_cofins(self):
        """
        Troca o CST de PIS (campo 25) e COFINS (campo 31) de todos os registros C170.
        """
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        cst_pis = simpledialog.askstring("Alterar CST", "Novo CST PIS (C170 - Campo 25):")
        if cst_pis is None:
            return
        cst_pis = cst_pis.strip()
        if not cst_pis:
            messagebox.showwarning("Aviso", "Informe o CST de PIS.")
            return
        
        cst_cofins = simpledialog.askstring("Alterar CST", "Novo CST COFINS (C170 - Campo 31):")
        if cst_cofins is None:
            return
        cst_cofins = cst_cofins.strip()
        if not cst_cofins:
            messagebox.showwarning("Aviso", "Informe o CST de COFINS.")
            return

        count = 0
        new_data = []
        csts_validos = {str(cst) for cst in range(50, 57)}
        for line in self.data:
            fields = line.strip().split("|")
            if len(fields) > 1:
                reg = fields[1]
                altered = False
                
                # Verifica registros C170, D101, D105 e F120
                if reg == "C170" and len(fields) > 31:
                    if fields[25].strip() in csts_validos:
                        fields[25] = cst_pis
                        altered = True
                    if fields[31].strip() in csts_validos:
                        fields[31] = cst_cofins
                        altered = True
                elif reg == "D101" and len(fields) > 4:
                    if fields[4].strip() in csts_validos:
                        fields[4] = cst_pis
                        altered = True
                elif reg == "D105" and len(fields) > 4:
                    if fields[4].strip() in csts_validos:
                        fields[4] = cst_cofins
                        altered = True
                elif reg == "F120" and len(fields) > 12:
                    if fields[8].strip() in csts_validos:
                        fields[8] = cst_pis
                        altered = True
                    if fields[12].strip() in csts_validos:
                        fields[12] = cst_cofins
                        altered = True
                
                if altered:
                    new_line = "|".join(fields) + "\n"
                    new_data.append(new_line)
                    count += 1
                else:
                    new_data.append(line)
            else:
                new_data.append(line)
        
        self.data = new_data
        self.display_data()
        messagebox.showinfo("Sucesso", f"CST PIS e COFINS alterados em {count} registros (C170/D101/D105/F120).")

    def importar_municipios_txt(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo TXT de municípios", filetypes=[("Text Files", "*.txt")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as file:
                linhas = file.readlines()
            inseridos = 0
            con = sqlite3.connect("referencias.db")
            cur = con.cursor()
            for linha in linhas:
                partes = linha.strip().split("|")
                if len(partes) >= 2:
                    codigo_cliente = partes[0].strip()
                    codigo_municipio = partes[1].strip()
                    if codigo_cliente and codigo_municipio:
                        cur.execute(
                            "INSERT OR REPLACE INTO municipios (codigo_cliente, codigo_municipio) VALUES (?, ?)",
                            (codigo_cliente, codigo_municipio)
                        )
                        inseridos += 1
            con.commit()
            con.close()
            messagebox.showinfo("Importação Concluída", f"{inseridos} municípios importados com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar TXT: {e}")

    def buscar_municipio_bd(self, codigo_cliente):
        try:
            con = sqlite3.connect("referencias.db")
            cur = con.cursor()
            cur.execute("SELECT codigo_municipio FROM municipios WHERE codigo_cliente = ?", (codigo_cliente,))
            row = cur.fetchone()
            return row[0] if row else None
        except Exception as e:
            messagebox.showerror("Erro BD", f"Erro ao consultar banco de dados: {e}")
            return None
        finally:
            con.close()

    def corrigir_0150_e_a170(self):
        alteracoes_0150 = 0
        alteracoes_a170 = 0
        for i, line in enumerate(self.data):
            fields = line.strip().split("|")
            if len(fields) > 8 and fields[1] == "0150" and fields[8] == "0":
                cod_cliente = fields[2]
                municipio = self.buscar_municipio_bd(cod_cliente)
                if not municipio:
                    municipio = simpledialog.askstring("Código Município", f"Cliente: {cod_cliente}\nDigite o código do município:")
                    if municipio and municipio.strip().isdigit():
                        try:
                            con = sqlite3.connect("referencias.db")
                            cur = con.cursor()
                            cur.execute("INSERT OR REPLACE INTO municipios (codigo_cliente, codigo_municipio) VALUES (?, ?)", (cod_cliente, municipio.strip()))
                            con.commit()
                            con.close()
                        except Exception as e:
                            messagebox.showerror("Erro BD", f"Erro ao salvar no banco: {e}")
                if municipio:
                    fields[8] = municipio.strip()
                    self.data[i] = "|".join(fields) + "\n"
                    alteracoes_0150 += 1
            if fields[1] == "A170":
                modificado = False
                if len(fields) > 15 and fields[15] == "1":
                    fields.pop(15)
                    modificado = True
                if len(fields) > 10 and fields[10] == "1":
                    fields.pop(10)
                    modificado = True
                if modificado:
                    self.data[i] = "|".join(fields) + "\n"
                    alteracoes_a170 += 1
        self.display_data()
        self.is_modified = True
        self.update_window_title()
        messagebox.showinfo("Concluído", f"Total de alterações:\n- {alteracoes_0150} registro(s) 0150 atualizado(s).\n- {alteracoes_a170} registro(s) A170 ajustado(s).")

    def update_d100_d101_d105_fields(self):
        default_value = "41332"
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 24:
                if fields[1] == "D100":
                    fields[23] = default_value
                elif fields[1] in ["D101", "D105"]:
                    if len(fields) > 10:
                        fields[9] = default_value
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Campos atualizados para os registros D100, D101 e D105.")

    def remove_spaces_between_pipes(self):
        def clean_line(line):
            fields = line.strip().split("|")
            cleaned_fields = [field.strip() for field in fields]
            return "|".join(cleaned_fields) + "\n"
        self.data = [clean_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Espaços extras entre delimitadores removidos.")

    def delete_c141(self):
        qtd_antes = len(self.data)
        self.data = [line for line in self.data if not (line.startswith("|C140|") or line.startswith("|C141|") or line.startswith("C140|") or line.startswith("C141|"))]
        qtd_depois = len(self.data)
        alteracoes = qtd_antes - qtd_depois
        self.display_data()
        self.is_modified = True
        self.update_window_title()
        messagebox.showinfo("Sucesso", f"Total de {alteracoes} registro(s) C140/C141 removido(s) do arquivo.")

    def contar_colunas(self):
        try:
            line_num = simpledialog.askinteger("Contar Colunas", "Digite o número da linha:")
            if line_num is None or line_num < 1 or line_num > len(self.data):
                messagebox.showerror("Erro", "Número de linha inválido.")
                return
            selected_line = self.data[line_num - 1].strip()
            num_columns = len(selected_line.split("|")) - 1
            messagebox.showinfo("Resultado", f"A linha {line_num} tem {num_columns} colunas.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

    def process_d101_d105(self):
        new_data = []
        for line in self.data:
            if not line:
                continue
            fields = line.strip().split("|")
            if len(fields) < 2:
                new_data.append(line)
                continue
            if fields[1] == "D101":
                if len(fields) < 10:
                    fields += [""] * (10 - len(fields))
                fields[9] = "4.1.01.03.0032|"
                new_line = "|".join(fields[:10]) + "\n"
                new_data.append(new_line)
                if len(fields) < 8:
                    continue
                campo2 = fields[2] if len(fields) > 2 else ""
                campo3 = fields[3] if len(fields) > 3 else ""
                campo4 = fields[4] if len(fields) > 4 else ""
                campo5 = fields[5] if len(fields) > 5 else ""
                campo6 = fields[6] if len(fields) > 6 else "0"
                campo7 = fields[7] if len(fields) > 7 else "0"
                if campo7.replace(",", ".") == "1.65":
                    campo7 = "7,60"
                try:
                    valor6 = float(campo6.replace(",", "."))
                    valor7 = float(campo7.replace(",", "."))
                    campo8 = f"{(valor6 * (valor7 / 100)):.2f}".replace(".", ",")
                except ValueError:
                    campo8 = "0"
                d105 = f"|D105|{campo2}|{campo3}|{campo4}|{campo5}|{campo6}|{campo7}|{campo8}|4.1.01.03.0032|\n"
                new_data.append(d105)
            else:
                new_data.append(line)
        self.data = new_data
        self.display_data()
        messagebox.showinfo("Sucesso", "Processamento de D101 e criação de D105 concluído.")

    def alterar_pis_cofins_d101_d105_por_txt(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado.")
            return

        file_path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT (numero_documento;pis;cofins)",
            filetypes=[("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not file_path:
            return

        mapa_valores = {}
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line_clean = line.strip()
                    if not line_clean or line_clean.startswith("#"):
                        continue

                    if ";" in line_clean:
                        parts = line_clean.split(";")
                    elif "\t" in line_clean:
                        parts = line_clean.split("\t")
                    elif "|" in line_clean:
                        parts = [p for p in line_clean.split("|") if p]
                    elif "," in line_clean and line_clean.count(",") >= 2:
                        parts = line_clean.split(",")
                    else:
                        continue

                    if len(parts) >= 3:
                        num_doc_raw = parts[0].strip().strip('"').strip("'")
                        pis_raw = parts[1].strip().strip('"').strip("'")
                        cofins_raw = parts[2].strip().strip('"').strip("'")

                        def formatar_valor(val_str):
                            if not val_str:
                                return "0,00"
                            val_clean = val_str.replace("R$", "").replace(" ", "").strip()
                            try:
                                if "," in val_clean and "." in val_clean:
                                    val_clean = val_clean.replace(".", "").replace(",", ".")
                                elif "," in val_clean:
                                    val_clean = val_clean.replace(",", ".")
                                val_float = float(val_clean)
                                return f"{val_float:.2f}".replace(".", ",")
                            except ValueError:
                                return val_str

                        pis_fmt = formatar_valor(pis_raw)
                        cofins_fmt = formatar_valor(cofins_raw)

                        num_doc_norm = num_doc_raw.lstrip("0")
                        val_dict = {"pis": pis_fmt, "cofins": cofins_fmt}
                        if num_doc_raw:
                            mapa_valores[num_doc_raw] = val_dict
                        if num_doc_norm:
                            mapa_valores[num_doc_norm] = val_dict
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo TXT:\n{e}")
            return

        if not mapa_valores:
            messagebox.showwarning("Aviso", "Nenhum valor válido foi lido do arquivo TXT.\nFormato esperado: numero_documento;pis;cofins")
            return

        new_data = []
        d100_ativo = False
        d100_valores = None
        d100_doc_atual = ""
        d100_afetados = set()
        d101_count = 0
        d105_count = 0
        tem_d105_no_d100 = False
        d101_linha_ref = None

        for line in self.data:
            ending = "\n" if line.endswith("\n") else ""
            line_clean = line.strip()
            campos = line_clean.split("|")

            if len(campos) < 2:
                new_data.append(line)
                continue

            reg = campos[1].strip()

            if reg == "D100":
                if d100_ativo and d100_valores and d101_linha_ref and not tem_d105_no_d100:
                    campos_ref = d101_linha_ref.split("|")
                    c2 = campos_ref[2] if len(campos_ref) > 2 else ""
                    c3 = campos_ref[3] if len(campos_ref) > 3 else ""
                    c4 = campos_ref[4] if len(campos_ref) > 4 else ""
                    c5 = campos_ref[5] if len(campos_ref) > 5 else ""
                    c6 = campos_ref[6] if len(campos_ref) > 6 else ""
                    c7 = campos_ref[7] if len(campos_ref) > 7 else ""
                    cta = campos_ref[9] if len(campos_ref) > 9 else ""
                    d105_gerado = f"|D105|{c2}|{c3}|{c4}|{c5}|{c6}|{c7}|{d100_valores['cofins']}|{cta}|\n"
                    new_data.append(d105_gerado)
                    d105_count += 1

                d100_doc_atual = campos[9].strip() if len(campos) > 9 else ""
                d100_doc_norm = d100_doc_atual.lstrip("0")

                if d100_doc_atual in mapa_valores:
                    d100_ativo = True
                    d100_valores = mapa_valores[d100_doc_atual]
                elif d100_doc_norm in mapa_valores:
                    d100_ativo = True
                    d100_valores = mapa_valores[d100_doc_norm]
                else:
                    d100_ativo = False
                    d100_valores = None

                tem_d105_no_d100 = False
                d101_linha_ref = None
                new_data.append(line)

            elif reg in ("D101", "D105") and d100_ativo and d100_valores:
                d100_afetados.add(d100_doc_atual)
                if reg == "D101":
                    while len(campos) <= 8:
                        campos.append("")
                    campos[8] = d100_valores["pis"]
                    d101_count += 1
                    d101_linha_ref = "|".join(campos)
                    new_line = "|".join(campos) + ending
                    new_data.append(new_line)
                elif reg == "D105":
                    tem_d105_no_d100 = True
                    while len(campos) <= 8:
                        campos.append("")
                    campos[8] = d100_valores["cofins"]
                    d105_count += 1
                    new_line = "|".join(campos) + ending
                    new_data.append(new_line)

            else:
                if reg in ("0000", "0150", "0200", "C100", "C500", "D500", "E110", "H010", "1010", "9900"):
                    if d100_ativo and d100_valores and d101_linha_ref and not tem_d105_no_d100:
                        campos_ref = d101_linha_ref.split("|")
                        c2 = campos_ref[2] if len(campos_ref) > 2 else ""
                        c3 = campos_ref[3] if len(campos_ref) > 3 else ""
                        c4 = campos_ref[4] if len(campos_ref) > 4 else ""
                        c5 = campos_ref[5] if len(campos_ref) > 5 else ""
                        c6 = campos_ref[6] if len(campos_ref) > 6 else ""
                        c7 = campos_ref[7] if len(campos_ref) > 7 else ""
                        cta = campos_ref[9] if len(campos_ref) > 9 else ""
                        d105_gerado = f"|D105|{c2}|{c3}|{c4}|{c5}|{c6}|{c7}|{d100_valores['cofins']}|{cta}|\n"
                        new_data.append(d105_gerado)
                        d105_count += 1

                    d100_ativo = False
                    d100_valores = None
                    tem_d105_no_d100 = False
                    d101_linha_ref = None

                new_data.append(line)

        if d100_ativo and d100_valores and d101_linha_ref and not tem_d105_no_d100:
            campos_ref = d101_linha_ref.split("|")
            c2 = campos_ref[2] if len(campos_ref) > 2 else ""
            c3 = campos_ref[3] if len(campos_ref) > 3 else ""
            c4 = campos_ref[4] if len(campos_ref) > 4 else ""
            c5 = campos_ref[5] if len(campos_ref) > 5 else ""
            c6 = campos_ref[6] if len(campos_ref) > 6 else ""
            c7 = campos_ref[7] if len(campos_ref) > 7 else ""
            cta = campos_ref[9] if len(campos_ref) > 9 else ""
            d105_gerado = f"|D105|{c2}|{c3}|{c4}|{c5}|{c6}|{c7}|{d100_valores['cofins']}|{cta}|\n"
            new_data.append(d105_gerado)
            d105_count += 1

        self.data = new_data
        self.display_data()

        messagebox.showinfo(
            "Sucesso",
            f"Processamento concluído com sucesso!\n\n"
            f"Documentos D100 afetados: {len(d100_afetados)}\n"
            f"Registros D101 (PIS) atualizados: {d101_count}\n"
            f"Registros D105 (COFINS) atualizados/gerados: {d105_count}"
        )

    def corrigir_frete_cst56_para_66_fornecedores(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado.")
            return

        fornecedores_alvo = {"FOR000002709", "FOR000002427"}

        msg = (
            "Deseja aplicar a correção de CST 56 -> 66 para os fretes (D100) dos fornecedores:\n"
            "- FOR000002709\n"
            "- FOR000002427\n\n"
            "Regras que serão aplicadas:\n"
            "1. Alterar CST (campo 4) de 56 para 66 em D101 e D105.\n"
            "2. Manter o campo 5 com o valor original.\n"
            "3. Definir o campo 6 (base de cálculo) igual ao campo 3 (VL_ITEM).\n"
            "4. Definir Alíquota PIS em D101 para 1,2375% e Alíquota COFINS em D105 para 5,7%.\n"
            "5. Calcular o Valor do Imposto (campo 8) = campo 3 * (Alíquota / 100)."
        )
        if not messagebox.askyesno("Confirmar Correção", msg):
            return

        new_data = []
        d100_qualificado = False
        d100_cod_part = ""
        d101_alterados = 0
        d105_alterados = 0

        aliq_pis_num = 1.2375 / 100.0
        aliq_cofins_num = 5.7 / 100.0

        def converter_valor(val_str):
            if not val_str:
                return 0.0
            val_clean = val_str.replace("R$", "").replace(" ", "").strip()
            try:
                if "," in val_clean and "." in val_clean:
                    val_clean = val_clean.replace(".", "").replace(",", ".")
                elif "," in val_clean:
                    val_clean = val_clean.replace(",", ".")
                return float(val_clean)
            except ValueError:
                return 0.0

        for line in self.data:
            ending = "\n" if line.endswith("\n") else ""
            line_clean = line.strip()
            campos = line_clean.split("|")

            if len(campos) < 2:
                new_data.append(line)
                continue

            reg = campos[1].strip()

            if reg == "D100":
                d100_cod_part = campos[4].strip() if len(campos) > 4 else ""
                cod_part_upper = d100_cod_part.upper()
                if cod_part_upper in fornecedores_alvo or d100_cod_part in fornecedores_alvo:
                    d100_qualificado = True
                else:
                    d100_qualificado = False
                new_data.append(line)

            elif reg in ("D101", "D105") and d100_qualificado:
                cst = campos[4].strip() if len(campos) > 4 else ""
                if cst == "56":
                    while len(campos) <= 8:
                        campos.append("")

                    campos[4] = "66"

                    val_base_num = converter_valor(campos[3] if len(campos) > 3 else "0")
                    val_base_str = campos[3] if len(campos) > 3 else "0,00"

                    # Campo 6 tem que ser igual ao campo 3 (campo 5 mantido inalterado)
                    campos[6] = val_base_str

                    if reg == "D101":
                        campos[7] = "1,2375"
                        vlr_pis = val_base_num * aliq_pis_num
                        campos[8] = f"{vlr_pis:.2f}".replace(".", ",")
                        d101_alterados += 1

                    elif reg == "D105":
                        campos[7] = "5,7"
                        vlr_cofins = val_base_num * aliq_cofins_num
                        campos[8] = f"{vlr_cofins:.2f}".replace(".", ",")
                        d105_alterados += 1

                    new_line = "|".join(campos) + ending
                    new_data.append(new_line)
                else:
                    new_data.append(line)

            else:
                if reg in ("0000", "0150", "0200", "C100", "C500", "D500", "E110", "H010", "1010", "9900"):
                    d100_qualificado = False
                new_data.append(line)

        self.data = new_data
        self.display_data()

        messagebox.showinfo(
            "Sucesso",
            f"Correção de Frete (CST 56 -> 66) concluída com sucesso!\n\n"
            f"Registros D101 (PIS) alterados: {d101_alterados}\n"
            f"Registros D105 (COFINS) alterados: {d105_alterados}"
        )

    def validar_d100(self):
        """
        Coleta as alterações nos registros D100 com base na regra:
        Se Campo 20 > 0, Campo 19 deve ser igual ao Campo 15.
        A validação é feita para os D100 correspondentes a D190 alterados.
        """
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return []

        # Dicionário para mapear D100 por número de nota fiscal para busca rápida
        d100_map = {}
        for i, line in enumerate(self.data):
            campos = line.strip().split('|')
            if len(campos) > 8 and campos[1] == 'D100':
                # Criando uma chave única (ex: TipoOperacao + ModeloDoc + NumeroDoc)
                chave = f"{campos[2]}|{campos[5]}|{campos[8]}"
                d100_map[chave] = {'index': i, 'fields': campos}
        
        changes = []
        for line_index, line in enumerate(self.data):
            campos = line.strip().split('|')
            # 1. Verificar o D190 primeiro
            if len(campos) > 7 and campos[1] == 'D190':
                try:
                    valor_campo7_d190 = float(campos[7].replace(',', '.'))
                    # 2. Se a condição do D190 for verdadeira
                    if valor_campo7_d190 > 0:
                        # 3. Encontrar o D100 correspondente (pai)
                        # O D190 não tem número de nota, mas o D100 (pai) sim.
                        # Precisamos subir no arquivo para encontrar o D100 mais próximo.
                        # Este é um método simples. Um mais robusto seria mapear C100 -> D100 -> D190
                        # Mas para a sua lógica, podemos seguir a mesma lógica de busca de pai.
                        
                        # A maneira mais segura de fazer isso sem um mapa complexo seria:
                        # Iterar para trás a partir do D190 para encontrar o D100 mais próximo
                        d100_pai = None
                        for i in range(line_index, -1, -1):
                            if self.data[i].strip().split('|')[1] == 'D100':
                                d100_pai = self.data[i].strip().split('|')
                                break
                        
                        if d100_pai and len(d100_pai) > 20: # O campo 20 é o índice 20
                            try:
                                # A validação é: se o campo 20 for > 0
                                valor_campo20_d100 = float(d100_pai[20].replace(',', '.'))
                                
                                # A correção é: campo 19 deve ser igual ao campo 15
                                if valor_campo20_d100 > 0 and d100_pai[19] != d100_pai[15]:
                                    original_d100_line = "|".join(d100_pai)
                                    d100_pai[19] = d100_pai[15]
                                    new_d100_line = "|".join(d100_pai)
                                    
                                    # Adiciona a alteração na lista
                                    changes.append((original_d100_line, new_d100_line))

                            except (ValueError, IndexError):
                                # Ignora D100 com valores ou campos inválidos
                                pass

                except (ValueError, IndexError):
                    # Ignora D190 com valores ou campos inválidos
                    pass
        
        return changes

    def run_validar_d100(self):
        """
        Ponto de entrada para a ação de correção do D100 com pré-visualização.
        """
        changes = self.validar_d100()
        self.show_changes_and_confirm(changes, "Corrigir D100")

    def process_0220_conversion(self):
        unit_map = {}
        products_with_0220 = set()
        new_data = []
        for i, line in enumerate(self.data):
            fields = line.strip().split("|")
            if len(fields) > 6 and fields[1] == "0200":
                product_code = fields[2]
                unit = fields[6]
                unit_map[product_code] = unit
                new_data.append(line)
                for j in range(i + 1, len(self.data)):
                    next_fields = self.data[j].strip().split("|")
                    if len(next_fields) > 6 and next_fields[1] == "C170":
                        used_product = next_fields[3]
                        used_unit = next_fields[6]
                        if used_product == product_code and used_unit != unit:
                            if product_code not in products_with_0220:
                                new_data.append(f"|0220|{used_unit}|1|\n")
                                products_with_0220.add(product_code)
                            break
            else:
                new_data.append(line)
        self.data = new_data
        self.display_data()
        messagebox.showinfo("Sucesso", "Registros 0220 adicionados abaixo dos 0200 necessários.")

    def filter_data(self):
        try:
            column = int(self.entry_filter_column.get().strip()) - 1
            value = self.entry_filter_value.get().strip()
            if not value:
                raise ValueError("O valor para o filtro não pode estar vazio.")
            filtered_data = [
                line for line in self.data
                if len(line.split("|")) > column and line.split("|")[column] == value
            ]
            if not filtered_data:
                messagebox.showinfo("Resultado", "Nenhum registro encontrado para o filtro.")
            else:
                self.text_display.delete("1.0", tk.END)
                for line in filtered_data:
                    self.text_display.insert(tk.END, line)
                self.filtered_data = filtered_data
        except ValueError as e:
            messagebox.showerror("Erro", f"Erro: {e}")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado: {e}")

    def reset_filter(self):
        self.display_data()
        messagebox.showinfo("Filtro", "Os filtros foram removidos.")

    def check_duplicates(self):
        try:
            col2 = "0200"
            col1_input = self.entry_column.get().strip()
            if not col1_input.isdigit():
                raise ValueError("A coluna informada deve ser um número válido.")
            col1 = int(col1_input)
            if col1 < 1 or col1 > len(self.data[0].split("|")):
                raise ValueError(f"A coluna {col1} não existe nos dados.")
            duplicates = []
            seen = set()
            for idx, line in enumerate(self.data, start=1):
                fields = line.split("|")
                if fields[1] == col2:
                    key = fields[2]
                    if key in seen:
                        duplicates.append(idx)
                    else:
                        seen.add(key)
            self.text_display.tag_remove("highlight", "1.0", tk.END)
            for line_num in duplicates:
                start_idx = f"{line_num}.0"
                end_idx = f"{line_num}.end"
                self.text_display.tag_add("highlight", start_idx, end_idx)
                self.text_display.tag_config("highlight", background="yellow", foreground="black")
            if duplicates:
                messagebox.showinfo("Duplicados Encontrados", f"Foram encontrados {len(duplicates)} duplicados.")
                answer = messagebox.askyesno("Excluir Duplicados", "Deseja excluir os duplicados encontrados e manter apenas a primeira ocorrência?")
                if answer:
                    self.data = [line for idx, line in enumerate(self.data, start=1) if idx not in duplicates or idx == duplicates[0]]
                    self.refresh_display()
                    messagebox.showinfo("Duplicados Excluídos", "Os duplicados foram excluídos, mantendo apenas a primeira ocorrência.")
            else:
                messagebox.showinfo("Sem Duplicados", "Nenhum duplicado encontrado.")
        except ValueError as e:
            messagebox.showerror("Erro", f"Erro: {e}")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado: {e}")

    def refresh_display(self):
        self.text_display.delete(1.0, tk.END)
        for line in self.data:
            self.text_display.insert(tk.END, line + '\n')

    def clear_field_6_if_condition(self):
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 5 and fields[1] == "0150" and fields[4] != "1058":
                fields[5] = ""
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Campo 6 modificado onde as condições foram atendidas.")

    def pad_field_8_if_condition(self):
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 8 and fields[1] == "0150" and fields[8][:2] == "31":
                if len(fields[7]) < 13:
                    fields[7] = fields[7].zfill(13)
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Campo 8 atualizado onde as condições foram atendidas.")

    def process_c197_lines(self):
        cfop_set = {
    1101, 1124, 1125, 1201, 1257, 1302, 1401, 1407,
    1551, 1556, 1651, 1653,
    1901, 1902, 1908, 1910, 1914, 1915, 1916,
    1920, 1921, 1949,
    2201, 2202, 2257, 2302,
    2551, 2556,
    2902, 2915,
    3101, 3102,
    3556,
    5101, 5102, 5124,
    5401, 5410, 5413,
    5501,
    5901, 5902, 5903, 5909,
    5910, 5913, 5914, 5915, 5916,
    5920, 5921,
    5949,
    6101, 6102, 6107, 6124,
    6401, 6403,
    6902, 6915, 6916, 6922, 6923,
    7101, 7102,
    7949
}

        
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 8 and fields[1] == "C197":
                try:
                    cfop = int(fields[3]) if fields[3].isdigit() else None
                    if cfop in cfop_set:
                        valor_coluna_7 = fields[7].replace(",", ".")
                        if float(valor_coluna_7) > 0:
                            fields[8] = fields[7]
                            fields[7] = "0"
                except ValueError:
                    pass
            return "|".join(fields)
        self.data = [modify_line(line) + "\n" for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Linhas 'C197' processadas com sucesso!")

    def update_h010_inventory(self):
        mapping = {
            "1": "1.1.10.01.0001", "2": "1.1.10.01.0002", "3": "1.1.10.01.0003",
            "4": "1.1.10.01.0004", "5": "1.1.10.01.0005", "6": "1.1.10.01.0006",
            "7": "1.1.10.01.0006", "8": "1.1.10.01.0008", "9": "1.1.10.01.0009",
            "0": "1.1.10.01.0010",
        }
        def modify_line(line):
            fields = line.strip().split("|")
            if len(fields) > 10 and fields[1] == "H010":
                first_digit = fields[2][0] if fields[2] else "0"
                if first_digit in mapping:
                    fields[10] = mapping[first_digit]
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Registros H010 atualizados com base na regra do INVENTÁRIO.")

    def sincronizar_unidade_h010_com_0200(self):
        """Atualiza o campo 3 do H010 com a unidade do campo 6 do 0200 para o mesmo COD_ITEM."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        mapa_unidades = {}
        for line in self.data:
            fields = line.strip().split("|")
            if len(fields) > 6 and fields[1] == "0200":
                cod_item = fields[2].strip()
                unid_inv = fields[6].strip()
                if cod_item and unid_inv:
                    mapa_unidades[cod_item] = unid_inv

        if not mapa_unidades:
            messagebox.showwarning("Aviso", "Nenhum registro 0200 com unidade encontrada.")
            return

        alterados = 0
        nao_encontrados = 0
        novas_linhas = []

        for line in self.data:
            fields = line.strip().split("|")
            if len(fields) > 3 and fields[1] == "H010":
                cod_item = fields[2].strip()
                unidade_0200 = mapa_unidades.get(cod_item)
                if unidade_0200:
                    if fields[3] != unidade_0200:
                        fields[3] = unidade_0200
                        alterados += 1
                    novas_linhas.append("|".join(fields) + "\n")
                else:
                    nao_encontrados += 1
                    novas_linhas.append("|".join(fields) + "\n")
            else:
                novas_linhas.append(line if line.endswith("\n") else line + "\n")

        self.data = novas_linhas
        self.display_data()
        messagebox.showinfo(
            "Sucesso",
            f"Sincronizacao concluida.\nRegistros H010 alterados: {alterados}\nSem 0200 correspondente: {nao_encontrados}"
        )

    def corrigir_h010_colunas(self):
        """
        Conta as colunas (delimitadores |) do registro H010. 
        Para o registro H010 ser considerado completo (11 campos), ele deve possuir 12 pipes.
        Se tiver menos (como o exemplo do usuário com 11 pipes), adiciona colunas no final.
        """
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        count = 0
        new_data = []
        for line in self.data:
            if not line:
                new_data.append(line)
                continue
            
            stripped_line = line.strip()
            if stripped_line.startswith("|H010|"):
                # Conta os pipes totais na linha
                num_pipes = stripped_line.count("|")
                
                # Se o registro tem 10 campos, ele tem 11 pipes.
                # Se o usuário quer adicionar um campo extra no exemplo enviado,
                # o objetivo é atingir 12 pipes (11 campos de dados).
                if num_pipes < 12:
                    fields = stripped_line.split("|")
                    while num_pipes < 12:
                        fields.insert(-1, "") # Adiciona campo vazio antes do último pipe
                        num_pipes += 1
                    
                    new_line = "|".join(fields) + "\n"
                    new_data.append(new_line)
                    count += 1
                else:
                    new_data.append(line)
            else:
                new_data.append(line)
        
        self.data = new_data
        self.display_data()
        messagebox.showinfo("Sucesso", f"{count} registros H010 processados e corrigidos para 12 colunas (|).")

    def atualizar_encerramento_bloco_0(self):
        """Atualiza o registro 0990 com a quantidade correta de linhas do Bloco 0."""
        if not self.data:
            return
        linhas_bloco0 = 0
        idx_0990 = None
        for idx, line in enumerate(self.data):
            fields = line.strip().split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if reg.startswith("0"):
                linhas_bloco0 += 1
                if reg == "0990":
                    idx_0990 = idx
                    break
        if idx_0990 is not None:
            self.data[idx_0990] = f"|0990|{linhas_bloco0}|\n"

    def verificar_e_importar_0200_faltantes_h010(self):
        """
        Verifica os itens do inventário (H010 campo 2 - COD_ITEM) contra o cadastro de itens (0200 campo 2).
        Se houver itens no H010 sem correspondente no 0200, exibe a lista dos faltantes e permite
        selecionar um arquivo TXT/SPED de referência para localizar, extrair e importar os registros 0200
        completos na sequência correta do Bloco 0 do arquivo aberto.
        """
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado no sistema.")
            return

        # 1. Mapear códigos 0200 existentes no arquivo aberto
        itens_0200_existentes = set()
        ultimo_idx_0200 = None
        idx_insercao_bloco0 = None

        for idx, line in enumerate(self.data):
            fields = line.strip().split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if reg == "0200" and len(fields) > 2:
                cod = fields[2].strip()
                if cod:
                    itens_0200_existentes.add(cod)
                ultimo_idx_0200 = idx
            elif reg in ("0205", "0206", "0210", "0220", "0221") and ultimo_idx_0200 is not None:
                ultimo_idx_0200 = idx
            elif reg in ("0190", "0150", "0100", "0005", "0001", "0000"):
                idx_insercao_bloco0 = idx

        # Ponto de inserção padrão para novos registros 0200
        ponto_insercao = (ultimo_idx_0200 + 1) if ultimo_idx_0200 is not None else ((idx_insercao_bloco0 + 1) if idx_insercao_bloco0 is not None else 0)

        # 2. Mapear itens do inventário H010
        itens_h010 = {}
        for line in self.data:
            fields = line.strip().split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if reg == "H010" and len(fields) > 2:
                cod = fields[2].strip()
                if not cod:
                    continue
                unid = fields[3].strip() if len(fields) > 3 else ""
                qtd_str = fields[4].strip() if len(fields) > 4 else "0"
                vl_item_str = fields[6].strip() if len(fields) > 6 else "0"
                if cod not in itens_h010:
                    itens_h010[cod] = {
                        "ocorrencias": 0,
                        "unidade": unid,
                        "qtd_total": 0.0,
                        "valor_total": 0.0
                    }
                itens_h010[cod]["ocorrencias"] += 1
                itens_h010[cod]["qtd_total"] += converter_valor_decimal(qtd_str)
                itens_h010[cod]["valor_total"] += converter_valor_decimal(vl_item_str)

        if not itens_h010:
            messagebox.showinfo("Inventário H010", "Nenhum registro H010 (Inventário) encontrado no arquivo aberto.")
            return

        # 3. Identificar itens faltantes no 0200
        faltantes = [cod for cod in itens_h010.keys() if cod not in itens_0200_existentes]

        # 4. Se todos os itens constam no 0200
        if not faltantes:
            messagebox.showinfo(
                "Inventário H010 vs 0200",
                f"Validação 100% Correta!\n\n"
                f"Total de itens no H010: {len(itens_h010):,}\n"
                f"Itens cadastrados no 0200: {len(itens_0200_existentes):,}\n\n"
                f"Todos os {len(itens_h010)} itens do inventário (H010) possuem o respectivo registro 0200 no arquivo aberto."
            )
            return

        # 5. Se houver itens faltantes, abrir janela modal interativa
        self._exibir_janela_0200_faltantes_h010(itens_h010, itens_0200_existentes, faltantes, ponto_insercao)

    def _exibir_janela_0200_faltantes_h010(self, itens_h010, itens_0200_existentes, faltantes, ponto_insercao):
        janela = tk.Toplevel(self.root)
        janela.title("Itens do Inventário (H010) Faltantes no Registro 0200")
        janela.geometry("820x560")
        janela.minsize(700, 450)
        janela.transient(self.root)
        janela.grab_set()

        # Centralizar na tela
        janela.update_idletasks()
        x = (janela.winfo_screenwidth() // 2) - (janela.winfo_width() // 2)
        y = (janela.winfo_screenheight() // 2) - (janela.winfo_height() // 2)
        janela.geometry(f"+{x}+{y}")

        # Frame de Cabeçalho / Resumo
        header_frame = tk.Frame(janela, bg="#f1f5f9", padx=16, pady=12, relief="groove", bd=1)
        header_frame.pack(fill="x", padx=12, pady=(12, 6))

        titulo_lbl = tk.Label(
            header_frame,
            text=f"⚠️ Foram encontrados {len(faltantes)} itens do inventário (H010) sem cadastro no 0200",
            font=("Segoe UI", 12, "bold"),
            fg="#b91c1c",
            bg="#f1f5f9"
        )
        titulo_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header_frame,
            text=f"Total no H010: {len(itens_h010):,} itens  |  Já cadastrados no 0200: {len(itens_h010) - len(faltantes):,}  |  Faltantes: {len(faltantes):,}",
            font=("Segoe UI", 10),
            fg="#475569",
            bg="#f1f5f9"
        )
        sub_lbl.pack(anchor="w", pady=(4, 0))

        # Frame de Busca rápida
        busca_frame = tk.Frame(janela, padx=12, pady=4)
        busca_frame.pack(fill="x")
        tk.Label(busca_frame, text="Filtrar Código:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))
        entry_filtro = ttk.Entry(busca_frame, width=25)
        entry_filtro.pack(side="left")

        # Frame da Tabela
        table_frame = tk.Frame(janela, padx=12, pady=6)
        table_frame.pack(fill="both", expand=True)

        colunas = ("item", "ocorrencias", "unidade", "qtd_total", "valor_total")
        tree = ttk.Treeview(table_frame, columns=colunas, show="headings", selectmode="extended")
        tree.heading("item", text="Código do Item (H010 Campo 2)")
        tree.heading("ocorrencias", text="Ocorrências no H010")
        tree.heading("unidade", text="Unidade (H010)")
        tree.heading("qtd_total", text="Qtd Total Inventário")
        tree.heading("valor_total", text="Valor Total (R$)")

        tree.column("item", width=220, anchor="w")
        tree.column("ocorrencias", width=130, anchor="center")
        tree.column("unidade", width=100, anchor="center")
        tree.column("qtd_total", width=150, anchor="e")
        tree.column("valor_total", width=160, anchor="e")

        sb_y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        sb_x = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        tree.grid(row=0, column=0, sticky="nsew")
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        def preencher_tabela(filtro=""):
            for i in tree.get_children():
                tree.delete(i)
            termo = filtro.strip().upper()
            for cod in faltantes:
                if termo and termo not in cod.upper():
                    continue
                info = itens_h010[cod]
                tree.insert(
                    "", "end",
                    values=(
                        cod,
                        f"{info['ocorrencias']}x",
                        info['unidade'],
                        f"{info['qtd_total']:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."),
                        f"R$ {info['valor_total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    )
                )

        preencher_tabela()
        entry_filtro.bind("<KeyRelease>", lambda e: preencher_tabela(entry_filtro.get()))

        # Ações / Funções
        def acao_importar_txt():
            caminho_ref = filedialog.askopenfilename(
                parent=janela,
                title="Selecione o arquivo TXT ou SPED com os registros 0200",
                filetypes=[
                    ("Arquivos Texto / SPED (*.txt;*.sped)", "*.txt;*.TXT;*.sped;*.SPED"),
                    ("Todos os arquivos (*.*)", "*.*")
                ]
            )
            if not caminho_ref:
                return

            try:
                conteudo = ler_arquivo_texto_multiencoding(caminho_ref)
                linhas_ref = conteudo.splitlines()
            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo selecionado:\n{e}", parent=janela)
                return

            set_faltantes = set(faltantes)
            registros_para_importar = {}  # cod_item -> list of lines (0200 and child records)
            cod_atual = None

            for line in linhas_ref:
                line_str = line.strip()
                if not line_str:
                    continue
                fields = line_str.split("|")
                reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")

                if reg == "0200" and len(fields) > 2:
                    cod = fields[2].strip()
                    if cod in set_faltantes and cod not in registros_para_importar:
                        registros_para_importar[cod] = [line_str]
                        cod_atual = cod
                    else:
                        cod_atual = None
                elif reg in ("0205", "0206", "0210", "0220", "0221") and cod_atual:
                    registros_para_importar[cod_atual].append(line_str)
                elif not reg.startswith("02"):
                    cod_atual = None

            if not registros_para_importar:
                messagebox.showwarning(
                    "Nenhum Registro Localizado",
                    f"O arquivo selecionado não contém nenhum dos {len(faltantes)} itens faltantes no registro 0200.",
                    parent=janela
                )
                return

            # Preparar as linhas a serem inseridas no arquivo aberto
            linhas_novas = []
            for cod in sorted(registros_para_importar.keys()):
                for sub_line in registros_para_importar[cod]:
                    linhas_novas.append(sub_line + "\n")

            # Localizar o ponto exato de inserção no Bloco 0
            ultimo_0200_idx = None
            idx_bloco0_fallback = None
            for idx, line in enumerate(self.data):
                fields = line.strip().split("|")
                reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
                if reg.startswith("02"):
                    ultimo_0200_idx = idx
                elif reg in ("0190", "0150", "0100", "0005", "0001", "0000"):
                    idx_bloco0_fallback = idx

            ins_pos = (ultimo_0200_idx + 1) if ultimo_0200_idx is not None else ((idx_bloco0_fallback + 1) if idx_bloco0_fallback is not None else 0)

            # Inserir na sequência do 0200 do arquivo aberto
            self.data = self.data[:ins_pos] + linhas_novas + self.data[ins_pos:]

            # Atualizar 0990 e Bloco 9
            self.atualizar_encerramento_bloco_0()
            self.recalcular_bloco_9()
            self.display_data()

            qtd_importados = len(registros_para_importar)
            ainda_faltantes = [cod for cod in faltantes if cod not in registros_para_importar]

            msg = (
                f"✅ Importação realizada com sucesso!\n\n"
                f"• Registros 0200 importados: {qtd_importados}\n"
                f"• Linhas totais inseridas no Bloco 0: {len(linhas_novas)}\n"
                f"• Itens ainda pendentes: {len(ainda_faltantes)}\n\n"
                f"Os novos registros foram inseridos na sequência correta do Bloco 0 e os totalizadores (0990 e Bloco 9) foram recalculados."
            )
            messagebox.showinfo("Sucesso", msg, parent=janela)

            if not ainda_faltantes:
                janela.destroy()
            else:
                # Atualizar a lista de faltantes na janela
                faltantes.clear()
                faltantes.extend(ainda_faltantes)
                sub_lbl.config(text=f"Total no H010: {len(itens_h010):,} itens  |  Já cadastrados no 0200: {len(itens_h010) - len(faltantes):,}  |  Faltantes: {len(faltantes):,}")
                preencher_tabela(entry_filtro.get())

        def acao_exportar_lista():
            caminho_salvar = filedialog.asksaveasfilename(
                parent=janela,
                title="Salvar Lista de Itens Faltantes",
                defaultextension=".txt",
                filetypes=[("Arquivo de Texto (*.txt)", "*.txt")]
            )
            if not caminho_salvar:
                return
            try:
                with open(caminho_salvar, "w", encoding="utf-8") as f:
                    f.write("COD_ITEM;OCORRENCIAS_H010;UNIDADE;VALOR_TOTAL\n")
                    for cod in faltantes:
                        info = itens_h010[cod]
                        f.write(f"{cod};{info['ocorrencias']};{info['unidade']};{info['valor_total']:.2f}\n")
                messagebox.showinfo("Exportado", f"Lista salva com sucesso em:\n{caminho_salvar}", parent=janela)
            except Exception as e:
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}", parent=janela)

        # Botões de Ação na base da janela
        btn_frame = tk.Frame(janela, padx=12, pady=12, bg="#f8fafc", relief="groove", bd=1)
        btn_frame.pack(fill="x", side="bottom")

        btn_importar = tk.Button(
            btn_frame,
            text="📁 Selecionar TXT/SPED de Referência para Importar 0200...",
            font=("Segoe UI", 10, "bold"),
            bg="#15803d",
            fg="white",
            activebackground="#166534",
            activeforeground="white",
            padx=14,
            pady=6,
            cursor="hand2",
            command=acao_importar_txt
        )
        btn_importar.pack(side="left", padx=(0, 10))

        btn_exportar = tk.Button(
            btn_frame,
            text="💾 Exportar Lista (TXT)",
            font=("Segoe UI", 9),
            padx=10,
            pady=6,
            command=acao_exportar_lista
        )
        btn_exportar.pack(side="left")

        btn_fechar = tk.Button(
            btn_frame,
            text="Fechar",
            font=("Segoe UI", 9),
            padx=12,
            pady=6,
            command=janela.destroy
        )
        btn_fechar.pack(side="right")
    def ajustar_valores_por_nota(self):
        # 1. Solicitar o número da nota
        numero_nota = simpledialog.askstring("Ajuste de Nota", "Digite o número da Nota Fiscal (Campo 08 do C100):")
        if not numero_nota:
            return

        numero_nota = numero_nota.strip()
        found = False
        lines_modified = 0
        
        # Variáveis de controle
        c100_idx = -1
        valor_total_nota = 0.0
        indices_c170 = []
        indices_c190 = []

        # 2. Localizar o C100 e seus filhos
        # Vamos iterar para encontrar o C100 específico
        for i, line in enumerate(self.data):
            campos = line.strip().split('|')
            
            # Se achamos o C100 alvo
            if len(campos) > 8 and campos[1] == 'C100' and campos[8] == numero_nota:
                try:
                    valor_total_nota = float(campos[12].replace(',', '.'))
                    c100_idx = i
                    found = True
                    
                    # Agora varremos as linhas seguintes para pegar os filhos (C170, C190)
                    # Paramos se encontrarmos outro C100 ou fim do arquivo
                    for j in range(i + 1, len(self.data)):
                        sub_campos = self.data[j].strip().split('|')
                        if len(sub_campos) < 2: continue
                        
                        registro = sub_campos[1]
                        
                        if registro == 'C100': # Chegamos na próxima nota, parar busca
                            break
                        
                        if registro == 'C170':
                            indices_c170.append(j)
                        elif registro == 'C190':
                            indices_c190.append(j)
                    
                    # Se achou a nota, paramos o loop principal pois vamos processar essa nota
                    break 
                except ValueError:
                    messagebox.showerror("Erro", f"Valor inválido no C100 da nota {numero_nota}")
                    return

        if not found:
            messagebox.showwarning("Não encontrado", f"Nota fiscal {numero_nota} não encontrada no registro C100.")
            return

        # 3. Função auxiliar para rateio
        def aplicar_rateio(indices, indice_campo_valor, valor_alvo_total):
            """
            Distribui o valor_alvo_total proporcionalmente entre as linhas indicadas.
            """
            if not indices: return 0
            
            # Soma os valores atuais para calcular a proporção
            soma_atual = 0.0
            valores_originais = []
            
            for idx in indices:
                linha = self.data[idx].strip().split('|')
                try:
                    val = float(linha[indice_campo_valor].replace(',', '.'))
                except (ValueError, IndexError):
                    val = 0.0
                valores_originais.append(val)
                soma_atual += val

            # Evita divisão por zero
            if soma_atual == 0:
                soma_atual = 1 

            soma_novos_valores = 0.0
            modificacoes = 0

            # Atualiza todos MENOS o último (para ajustar centavos no final)
            for k, idx in enumerate(indices[:-1]):
                linha = self.data[idx].strip().split('|')
                
                # Regra de 3: (Valor Original / Soma Original) * Novo Total
                peso = valores_originais[k] / soma_atual
                novo_valor = valor_alvo_total * peso
                
                # Formata e Salva
                linha[indice_campo_valor] = f"{novo_valor:.2f}".replace('.', ',')
                self.data[idx] = "|".join(linha) + "\n"
                
                soma_novos_valores += novo_valor
                modificacoes += 1

            # 4. Ajuste final (sobra de centavos) no último item
            ultimo_idx = indices[-1]
            linha_ultima = self.data[ultimo_idx].strip().split('|')
            
            valor_restante = valor_alvo_total - soma_novos_valores
            # Garante que não fique negativo por erro de arredondamento ínfimo, mas matematicamente deve ser exato
            if valor_restante < 0: valor_restante = 0 
            
            linha_ultima[indice_campo_valor] = f"{valor_restante:.2f}".replace('.', ',')
            self.data[ultimo_idx] = "|".join(linha_ultima) + "\n"
            modificacoes += 1
            
            return modificacoes

        # 5. Executar os ajustes
        
        # Rateio no C170 (Campo 7)
        if indices_c170:
            lines_modified += aplicar_rateio(indices_c170, 7, valor_total_nota)
        
        # Rateio no C190 (Campo 5 e Campo 10)
        # Nota: O C190 também deve ser rateado se houver mais de um, 
        # para que a soma dos C190 bata com o C100.
        if indices_c190:
            # Rateia o campo 5 (Valor Operação)
            lines_modified += aplicar_rateio(indices_c190, 5, valor_total_nota)
            # Rateia o campo 10 (conforme seu pedido)
            lines_modified += aplicar_rateio(indices_c190, 10, valor_total_nota)

        # 6. Atualizar a visualização e focar na nota
        self.display_data()
        
        # Dar destaque visual e rolar até a nota
        self.text_display.tag_remove("highlight", "1.0", tk.END)
        
        # Destacar C100
        start = f"{c100_idx + 1}.0"
        end = f"{c100_idx + 1}.end"
        self.text_display.tag_add("highlight", start, end)
        
        # Destacar Filhos
        for idx in indices_c170 + indices_c190:
            s = f"{idx + 1}.0"
            e = f"{idx + 1}.end"
            self.text_display.tag_add("highlight", s, e)
            
        self.text_display.tag_config("highlight", background="lightgreen", foreground="black")
        self.text_display.see(start)

        messagebox.showinfo("Sucesso", 
                            f"Nota {numero_nota} ajustada!\n"
                            f"Valor Alvo (C100): {valor_total_nota:.2f}\n"
                            f"Registros C170 ajustados: {len(indices_c170)}\n"
                            f"Registros C190 ajustados: {len(indices_c190)}")
        
    def totalizar_c190(self):
        c190_totals = {}
        new_data = []
        skip_next = False
        for line in self.data:
            fields = line.strip().split("|")
            if skip_next:
                if fields[1] in ("C195", "C197"):
                    continue
                else:
                    skip_next = False
            if len(fields) > 11 and fields[1] == "C190" and fields[5] in ("0", "0,00"):
                key = fields[3]
                try:
                    value = float(fields[5].replace(",", "."))
                except ValueError:
                    value = 0
                if key in c190_totals:
                    c190_totals[key] += value
                else:
                    c190_totals[key] = value
                new_data.append(line)
                new_data.append(f"|C195|{key}|SP Valor correspondente à coluna Isentas/Não tributadas e Outras|")
                new_data.append(f"|C197|SP90090104|{key}||0|0|0|{c190_totals[key]:.2f}|")
                skip_next = True
            else:
                new_data.append(line)
        self.data = new_data
        self.display_data()
        self.text_display.tag_remove("highlight", "1.0", tk.END)
        for i, line in enumerate(self.data):
            if "|C195|" in line or "|C197|" in line:
                start_idx = f"{i + 1}.0"
                end_idx = f"{i + 1}.end"
                self.text_display.tag_add("highlight", start_idx, end_idx)
                self.text_display.tag_config("highlight", background="yellow", foreground="black")
        messagebox.showinfo("Sucesso", "Totalização do C190 e inclusão dos registros C195 e C197 concluída!")
    def validar_totais_c100_c170(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        erros = []
        
        # Variáveis de controle
        c100_atual = None 
        soma_itens = 0.0
        cfops_encontrados = set() # Conjunto para guardar CFOPs únicos da nota
        
        # --- 1. PROCESSAMENTO DOS DADOS ---
        for i, line in enumerate(self.data):
            campos = line.strip().split('|')
            if len(campos) < 2: continue
            
            registro = campos[1]
            
            # Valida o fechamento do C100 anterior antes de começar um novo
            if registro == 'C100':
                if c100_atual:
                    diferenca = abs(c100_atual['valor'] - soma_itens)
                    if diferenca > 0.01:
                        # Converte o set de CFOPs em string (ex: "5102, 5405")
                        cfops_str = ",".join(sorted(cfops_encontrados))
                        erros.append({
                            'linha': c100_atual['linha_num'],
                            'nota': c100_atual['num_nota'],
                            'cfop': cfops_str,
                            'valor_c100': c100_atual['valor'],
                            'soma_c170': soma_itens,
                            'diff': diferenca
                        })

                # Inicia contexto do novo C100
                try:
                    valor_doc = float(campos[12].replace(',', '.')) if len(campos) > 12 and campos[12] else 0.0
                    num_nota = campos[8] if len(campos) > 8 else "S/N"
                    
                    c100_atual = {
                        'linha_num': i + 1,
                        'num_nota': num_nota,
                        'valor': valor_doc
                    }
                    soma_itens = 0.0
                    cfops_encontrados = set() # Reseta lista de CFOPs
                except ValueError:
                    c100_atual = None
            
            # Acumula valores dos itens (C170)
            elif registro == 'C170' and c100_atual:
                try:
                    # Valor do item (Coluna 7)
                    val_item = float(campos[7].replace(',', '.')) if len(campos) > 7 and campos[7] else 0.0
                    soma_itens += val_item
                    
                    # Captura CFOP (Coluna 11)
                    if len(campos) > 11:
                        cfop = campos[11].strip()
                        if cfop:
                            cfops_encontrados.add(cfop)
                except ValueError:
                    pass

        # Valida o último registro do arquivo
        if c100_atual:
            diferenca = abs(c100_atual['valor'] - soma_itens)
            if diferenca > 0.01:
                cfops_str = ",".join(sorted(cfops_encontrados))
                erros.append({
                    'linha': c100_atual['linha_num'],
                    'nota': c100_atual['num_nota'],
                    'cfop': cfops_str,
                    'valor_c100': c100_atual['valor'],
                    'soma_c170': soma_itens,
                    'diff': diferenca
                })

        # --- 2. EXIBIÇÃO DA TABELA (GUI) ---
        if not erros:
            messagebox.showinfo("Sucesso", "Nenhuma divergência encontrada entre C100 e C170.")
        else:
            # === NOVA LINHA: ORDENAR POR CFOP ===
            # Isso organiza a lista alfabeticamente pelo campo CFOP antes de exibir
            erros.sort(key=lambda x: x['cfop']) 

            relatorio_win = tk.Toplevel(self.root)
            relatorio_win.title(f"Relatório de Divergências ({len(erros)} erros)")
            relatorio_win.geometry("900x500")

            # Label Topo
            tk.Label(relatorio_win, text="Divergências de Totais (Ordenado por CFOP)", font=("Arial", 11, "bold")).pack(pady=5)

            # Frame para a Tabela e Scrollbar
            frame_table = tk.Frame(relatorio_win)
            frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

            # Colunas
            columns = ("linha", "nota", "cfop", "vlr_c100", "soma_c170", "diferenca")
            tree = ttk.Treeview(frame_table, columns=columns, show="headings")
            
            # Cabeçalhos
            tree.heading("linha", text="Linha")
            tree.heading("nota", text="Nota Fiscal")
            tree.heading("cfop", text="CFOP")
            tree.heading("vlr_c100", text="Vlr Nota (C100)")
            tree.heading("soma_c170", text="Soma Itens (C170)")
            tree.heading("diferenca", text="Diferença")

            # Tamanho das Colunas
            tree.column("linha", width=80, anchor="center")
            tree.column("nota", width=100, anchor="center")
            tree.column("cfop", width=100, anchor="center")
            tree.column("vlr_c100", width=120, anchor="e")
            tree.column("soma_c170", width=120, anchor="e")
            tree.column("diferenca", width=100, anchor="e")

            # Scrollbar Vertical
            scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscroll=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Preencher Tabela
            for erro in erros:
                # Formatação BR
                v_c100 = f"{erro['valor_c100']:.2f}".replace('.', ',')
                v_c170 = f"{erro['soma_c170']:.2f}".replace('.', ',')
                v_diff = f"{erro['diff']:.2f}".replace('.', ',')
                
                tree.insert("", tk.END, values=(
                    erro['linha'], 
                    erro['nota'], 
                    erro['cfop'],
                    v_c100, 
                    v_c170, 
                    v_diff
                ))
            def copiar_notas_para_lote():
                # Extrai os números das notas, remove duplicatas com set() e ordena
                notas_unicas = sorted(list(set(str(erro['nota']) for erro in erros)))
                # Junta tudo com ponto e vírgula
                string_notas = ",".join(notas_unicas)
                
                # Limpa a área de transferência e adiciona o novo conteúdo
                self.root.clipboard_clear()
                self.root.clipboard_append(string_notas)
                self.root.update() # Garante que o sistema operacional receba a atualização
                
                messagebox.showinfo("Copiado", f"{len(notas_unicas)} notas copiadas para a área de transferência!\nPronto para colar no Ajuste em Lote.")

            # --- ADIÇÃO DOS BOTÕES NO RODAPÉ ---
            btn_frame_acoes = tk.Frame(relatorio_win)
            btn_frame_acoes.pack(pady=10, fill=tk.X, padx=20)

            # Botão de Copiar (Novo)
            btn_copy = tk.Button(btn_frame_acoes, text="Copiar Notas p/ Lote (,)", 
                                 command=copiar_notas_para_lote, 
                                 bg="#13BFD6", fg="white", font=("Arial", 10, "bold"), height=2)
            btn_copy.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

            # Botão Exportar (Já existente, apenas movido para o frame de botões)
            btn_export = tk.Button(btn_frame_acoes, text="Exportar para XLS", 
                                   command=lambda: exportar_xls(), height=2)
            btn_export.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

            # --- 3. FUNÇÃO DE EXPORTAÇÃO ---
            def exportar_txt():
                save_path = filedialog.asksaveasfilename(
                    title="Exportar Relatório",
                    defaultextension=".txt",
                    filetypes=[("Arquivo Texto", "*.txt")]
                )
                if not save_path:
                    return
                
                try:
                    with open(save_path, "w", encoding="utf-8") as f:
                        # Cabeçalho
                        f.write("| LINHA | NOTA FISCAL | CFOP | VLR NOTA | SOMA ITENS | DIFERENÇA |\n")
                        f.write("|" + "-"*75 + "|\n")
                        
                        # Linhas (já estão ordenadas por CFOP pois a lista 'erros' foi ordenada)
                        for erro in erros:
                            v_c100 = f"{erro['valor_c100']:.2f}".replace('.', ',')
                            v_c170 = f"{erro['soma_c170']:.2f}".replace('.', ',')
                            v_diff = f"{erro['diff']:.2f}".replace('.', ',')
                            
                            line_str = f"| {erro['linha']:<5} | {erro['nota']:<11} | {erro['cfop']:<8} | {v_c100:<10} | {v_c170:<10} | {v_diff:<9} |\n"
                            f.write(line_str)
                            
                    messagebox.showinfo("Sucesso", f"Relatório exportado para:\n{save_path}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao exportar: {e}")

            # Botão Exportar

            def exportar_xls():
                save_path = filedialog.asksaveasfilename(
                    title="Exportar RelatÇürio",
                    defaultextension=".xls",
                    filetypes=[("Arquivo Excel", "*.xls")]
                )
                if not save_path:
                    return

                try:
                    linhas_exportacao = []
                    for erro in erros:
                        linhas_exportacao.append({
                            "LINHA": erro["linha"],
                            "NOTA FISCAL": erro["nota"],
                            "CFOP": erro["cfop"],
                            "VLR NOTA": round(float(erro["valor_c100"]), 2),
                            "SOMA ITENS": round(float(erro["soma_c170"]), 2),
                            "DIFERENCA": round(float(erro["diff"]), 2),
                        })

                    df_exportacao = pd.DataFrame(linhas_exportacao)
                    tabela_html = df_exportacao.to_html(index=False, border=1)
                    conteudo_xls = (
                        "<html><head><meta charset=\"utf-8\"></head><body>"
                        f"{tabela_html}"
                        "</body></html>"
                    )

                    with open(save_path, "w", encoding="utf-8") as f:
                        f.write(conteudo_xls)

                    messagebox.showinfo("Sucesso", f"RelatÇürio exportado para:\n{save_path}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao exportar XLS: {e}")


    def listar_campos_c100_25_26(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        registros_encontrados = []

        indice = 0
        while indice < len(self.data):
            line = self.data[indice]
            campos = line.strip().split('|')
            if len(campos) > 18 and campos[1] == "A100":
                def obter_valor_a100(*indices):
                    primeiro_valor = ""
                    for indice_campo in indices:
                        if len(campos) <= indice_campo:
                            continue

                        valor = campos[indice_campo].strip()
                        if not valor:
                            continue

                        if not primeiro_valor:
                            primeiro_valor = valor

                        try:
                            if float(valor.replace('.', '').replace(',', '.')) != 0:
                                return valor
                        except ValueError:
                            return valor

                    return primeiro_valor

                pis = obter_valor_a100(16, 19)
                cofins = obter_valor_a100(18, 20)
                csts_pis_a100 = []
                csts_cofins_a100 = []
                proximo_indice = indice + 1
                pis_a170 = 0.0
                cofins_a170 = 0.0
                encontrou_a170 = False

                while proximo_indice < len(self.data):
                    campos_filho = self.data[proximo_indice].strip().split('|')
                    if len(campos_filho) < 2:
                        proximo_indice += 1
                        continue

                    registro_filho = campos_filho[1]
                    if registro_filho == "A100" or registro_filho.startswith("A0") or registro_filho.startswith("C") or registro_filho.startswith("D") or registro_filho.startswith("9"):
                        break
                    if registro_filho == "A170":
                        encontrou_a170 = True
                        try:
                            pis_a170 += float((campos_filho[12].strip() if len(campos_filho) > 12 else "0").replace('.', '').replace(',', '.'))
                        except ValueError:
                            pass
                        try:
                            cofins_a170 += float((campos_filho[16].strip() if len(campos_filho) > 16 else "0").replace('.', '').replace(',', '.'))
                        except ValueError:
                            pass
                        cst_p = campos_filho[9].strip() if len(campos_filho) > 9 else ""
                        cst_c = campos_filho[14].strip() if len(campos_filho) > 14 else ""
                        if cst_p and cst_p not in csts_pis_a100:
                            csts_pis_a100.append(cst_p)
                        if cst_c and cst_c not in csts_cofins_a100:
                            csts_cofins_a100.append(cst_c)

                    proximo_indice += 1

                if encontrou_a170:
                    try:
                        valor_pis = float((pis or "0").replace('.', '').replace(',', '.'))
                    except ValueError:
                        valor_pis = 0.0
                    try:
                        valor_cofins = float((cofins or "0").replace('.', '').replace(',', '.'))
                    except ValueError:
                        valor_cofins = 0.0

                    if valor_pis == 0 and pis_a170 > 0:
                        pis = f"{pis_a170:.2f}".replace('.', ',')
                    if valor_cofins == 0 and cofins_a170 > 0:
                        cofins = f"{cofins_a170:.2f}".replace('.', ',')

                registros_encontrados.append({
                    "registro": "A100",
                    "campo_2": campos[2].strip(),
                    "num_doc": campos[8].strip() if len(campos) > 8 else "",
                    "icms": "",
                    "ipi": "",
                    "cst_pis": ", ".join(csts_pis_a100),
                    "pis": pis,
                    "cst_cofins": ", ".join(csts_cofins_a100),
                    "cofins": cofins,
                })
            elif len(campos) > 27 and campos[1] == "C100":
                proximo_indice = indice + 1
                csts_pis_c100 = []
                csts_cofins_c100 = []

                while proximo_indice < len(self.data):
                    campos_filho = self.data[proximo_indice].strip().split('|')
                    if len(campos_filho) < 2:
                        proximo_indice += 1
                        continue

                    registro_filho = campos_filho[1]
                    if registro_filho == "C100" or registro_filho.startswith("C0") or registro_filho.startswith("D") or registro_filho.startswith("9"):
                        break
                    if registro_filho == "C170":
                        cst_p = campos_filho[25].strip() if len(campos_filho) > 25 else ""
                        cst_c = campos_filho[31].strip() if len(campos_filho) > 31 else ""
                        if cst_p and cst_p not in csts_pis_c100:
                            csts_pis_c100.append(cst_p)
                        if cst_c and cst_c not in csts_cofins_c100:
                            csts_cofins_c100.append(cst_c)
                    elif registro_filho == "C175":
                        cst_p = campos_filho[5].strip() if len(campos_filho) > 5 else ""
                        cst_c = campos_filho[9].strip() if len(campos_filho) > 9 else ""
                        if cst_p and cst_p not in csts_pis_c100:
                            csts_pis_c100.append(cst_p)
                        if cst_c and cst_c not in csts_cofins_c100:
                            csts_cofins_c100.append(cst_c)

                    proximo_indice += 1

                registros_encontrados.append({
                    "registro": "C100",
                    "campo_2": campos[2].strip(),
                    "num_doc": campos[8].strip(),
                    "icms": campos[22].strip(),
                    "ipi": campos[25].strip(),
                    "cst_pis": ", ".join(csts_pis_c100),
                    "pis": campos[26].strip(),
                    "cst_cofins": ", ".join(csts_cofins_c100),
                    "cofins": campos[27].strip(),
                })
            elif len(campos) > 14 and campos[1] == "C500":
                proximo_indice = indice + 1
                csts_pis_c500 = []
                csts_cofins_c500 = []

                while proximo_indice < len(self.data):
                    campos_filho = self.data[proximo_indice].strip().split('|')
                    if len(campos_filho) < 2:
                        proximo_indice += 1
                        continue

                    registro_filho = campos_filho[1]
                    if registro_filho == "C500" or registro_filho.startswith("C0") or registro_filho.startswith("D") or registro_filho.startswith("9"):
                        break
                    if registro_filho == "C501" and len(campos_filho) > 2:
                        cst_p = campos_filho[2].strip()
                        if cst_p and cst_p not in csts_pis_c500:
                            csts_pis_c500.append(cst_p)
                    elif registro_filho == "C505" and len(campos_filho) > 2:
                        cst_c = campos_filho[2].strip()
                        if cst_c and cst_c not in csts_cofins_c500:
                            csts_cofins_c500.append(cst_c)

                    proximo_indice += 1

                registros_encontrados.append({
                    "registro": "C500",
                    "campo_2": campos[2].strip(),
                    "num_doc": campos[7].strip(),
                    "icms": campos[11].strip(),
                    "ipi": "",
                    "cst_pis": ", ".join(csts_pis_c500),
                    "pis": campos[13].strip(),
                    "cst_cofins": ", ".join(csts_cofins_c500),
                    "cofins": campos[14].strip(),
                })
            elif len(campos) > 9 and campos[1] == "D100":
                pis = ""
                cofins = ""
                csts_pis_d100 = []
                csts_cofins_d100 = []
                proximo_indice = indice + 1

                while proximo_indice < len(self.data):
                    campos_filho = self.data[proximo_indice].strip().split('|')
                    if len(campos_filho) < 2:
                        proximo_indice += 1
                        continue

                    registro_filho = campos_filho[1]
                    if registro_filho == "D100" or registro_filho.startswith("D0") or registro_filho.startswith("9"):
                        break
                    if registro_filho == "D101" and len(campos_filho) > 8:
                        pis = campos_filho[8].strip()
                        cst_p = campos_filho[4].strip() if len(campos_filho) > 4 else ""
                        if cst_p and cst_p not in csts_pis_d100:
                            csts_pis_d100.append(cst_p)
                        if len(campos_filho) > 13 and campos_filho[13].strip():
                            cofins = campos_filho[13].strip()
                            cst_c = campos_filho[10].strip() if len(campos_filho) > 10 else ""
                            if cst_c and cst_c not in csts_cofins_d100:
                                csts_cofins_d100.append(cst_c)
                    elif registro_filho == "D105" and len(campos_filho) > 8:
                        cofins = campos_filho[8].strip()
                        cst_c = campos_filho[4].strip() if len(campos_filho) > 4 else ""
                        if cst_c and cst_c not in csts_cofins_d100:
                            csts_cofins_d100.append(cst_c)

                    proximo_indice += 1

                registros_encontrados.append({
                    "registro": "D100",
                    "campo_2": campos[2].strip(),
                    "num_doc": campos[9].strip(),
                    "icms": campos[20].strip() if len(campos) > 20 else "",
                    "ipi": "",
                    "cst_pis": ", ".join(csts_pis_d100),
                    "pis": pis,
                    "cst_cofins": ", ".join(csts_cofins_d100),
                    "cofins": cofins,
                })
            elif len(campos) > 14 and campos[1] == "F100":
                registros_encontrados.append({
                    "registro": "F100",
                    "campo_2": campos[2].strip(),
                    "num_doc": "",
                    "icms": "",
                    "ipi": "",
                    "cst_pis": campos[7].strip() if len(campos) > 7 else "",
                    "pis": campos[10].strip(),
                    "cst_cofins": campos[11].strip() if len(campos) > 11 else "",
                    "cofins": campos[14].strip(),
                })
            elif len(campos) > 15 and campos[1] == "F120":
                registros_encontrados.append({
                    "registro": "F120",
                    "campo_2": campos[2].strip(),
                    "num_doc": "",
                    "icms": "",
                    "ipi": "",
                    "cst_pis": campos[7].strip() if len(campos) > 7 else "",
                    "pis": campos[11].strip(),
                    "cst_cofins": campos[11].strip() if len(campos) > 11 else "",
                    "cofins": campos[15].strip(),
                })
            elif len(campos) > 10 and campos[1] == "F600":
                cst_f600 = campos[8].strip() if len(campos) > 8 else ""
                registros_encontrados.append({
                    "registro": "F600",
                    "campo_2": campos[2].strip(),
                    "num_doc": "",
                    "icms": "",
                    "ipi": "",
                    "cst_pis": cst_f600,
                    "pis": campos[9].strip(),
                    "cst_cofins": cst_f600,
                    "cofins": campos[10].strip(),
                })

            indice += 1

        if not registros_encontrados:
            messagebox.showinfo("Resultado", "Nenhum registro A100, C100, C500, D100, F100, F120 ou F600 com os campos solicitados foi encontrado.")
            return

        relatorio_win = tk.Toplevel(self.root)
        relatorio_win.title(f"Lista A100/C100/C500/D100/F100/F120/F600 - Documento, ICMS, IPI, PIS, COFINS e CST ({len(registros_encontrados)} registros)")
        relatorio_win.geometry("1200x560")

        tk.Label(
            relatorio_win,
            text="Campo 2, Número do Documento, ICMS, IPI, CST PIS, PIS, CST COFINS e COFINS dos registros A100, C100, C500, D100, F100, F120 e F600",
            font=("Arial", 11, "bold")
        ).pack(pady=5)

        frame_table = tk.Frame(relatorio_win)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("registro", "campo_2", "num_doc", "icms", "ipi", "cst_pis", "pis", "cst_cofins", "cofins")
        tree = ttk.Treeview(frame_table, columns=columns, show="headings")

        tree.heading("registro", text="Registro")
        tree.heading("campo_2", text="Campo 2")
        tree.heading("num_doc", text="Número do Documento")
        tree.heading("icms", text="ICMS")
        tree.heading("ipi", text="IPI")
        tree.heading("cst_pis", text="CST PIS")
        tree.heading("pis", text="PIS")
        tree.heading("cst_cofins", text="CST COFINS")
        tree.heading("cofins", text="COFINS")

        tree.column("registro", width=80, anchor="center")
        tree.column("campo_2", width=90, anchor="center")
        tree.column("num_doc", width=160, anchor="center")
        tree.column("icms", width=110, anchor="e")
        tree.column("ipi", width=110, anchor="e")
        tree.column("cst_pis", width=90, anchor="center")
        tree.column("pis", width=140, anchor="e")
        tree.column("cst_cofins", width=100, anchor="center")
        tree.column("cofins", width=140, anchor="e")

        scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for registro in registros_encontrados:
            tree.insert(
                "",
                tk.END,
                values=(
                    registro["registro"],
                    registro["campo_2"],
                    registro["num_doc"],
                    registro["icms"],
                    registro["ipi"],
                    registro["cst_pis"],
                    registro["pis"],
                    registro["cst_cofins"],
                    registro["cofins"],
                )
            )

        def exportar_xls():
            save_path = filedialog.asksaveasfilename(
                title="Exportar lista A100/C100/C500/D100/F100/F120/F600",
                defaultextension=".xls",
                filetypes=[("Arquivo Excel", "*.xls")]
            )
            if not save_path:
                return

            try:
                df_exportacao = pd.DataFrame([
                    {
                        "REGISTRO": registro["registro"],
                        "CAMPO 2": registro["campo_2"],
                        "NUMERO DO DOCUMENTO": registro["num_doc"],
                        "ICMS": registro["icms"],
                        "IPI": registro["ipi"],
                        "CST PIS": registro.get("cst_pis", ""),
                        "PIS": registro["pis"],
                        "CST COFINS": registro.get("cst_cofins", ""),
                        "COFINS": registro["cofins"],
                    }
                    for registro in registros_encontrados
                ])

                tabela_html = df_exportacao.to_html(index=False, border=1)
                conteudo_xls = (
                    "<html><head><meta charset=\"utf-8\"></head><body>"
                    f"{tabela_html}"
                    "</body></html>"
                )

                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(conteudo_xls)

                messagebox.showinfo("Sucesso", f"Lista exportada para:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar XLS: {e}")

        btn_frame = tk.Frame(relatorio_win)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        tk.Button(
            btn_frame,
            text="Exportar para XLS",
            command=exportar_xls,
            height=2
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def listar_campos_c170_cfop_pis_cofins(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo carregado.")
            return

        registros_encontrados = []
        c100_ind_oper = ""

        for line in self.data:
            campos = line.strip().split('|')
            if len(campos) > 2 and campos[1] == "C100":
                c100_ind_oper = campos[2].strip()
            elif len(campos) > 36 and campos[1] == "C170":
                tipo_oper = (
                    "0 - Entrada" if c100_ind_oper == "0"
                    else ("1 - Saída" if c100_ind_oper == "1"
                    else c100_ind_oper)
                )
                registros_encontrados.append({
                    "registro": campos[1].strip(),
                    "ind_oper": tipo_oper,
                    "cfop": campos[11].strip(),
                    "pis": campos[30].strip(),
                    "cofins": campos[36].strip(),
                })

        if not registros_encontrados:
            messagebox.showinfo("Resultado", "Nenhum registro C170 com os campos solicitados foi encontrado.")
            return

        relatorio_win = tk.Toplevel(self.root)
        relatorio_win.title(f"Lista C170 - CFOP, PIS e Cofins ({len(registros_encontrados)} registros)")
        relatorio_win.geometry("820x520")

        tk.Label(
            relatorio_win,
            text="Registro, Entrada/Saída, CFOP, valor PIS e valor Cofins dos registros C170",
            font=("Arial", 11, "bold")
        ).pack(pady=5)

        frame_table = tk.Frame(relatorio_win)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("registro", "ind_oper", "cfop", "pis", "cofins")
        tree = ttk.Treeview(frame_table, columns=columns, show="headings")

        tree.heading("registro", text="Registro")
        tree.heading("ind_oper", text="Entrada / Saída")
        tree.heading("cfop", text="CFOP")
        tree.heading("pis", text="Valor PIS")
        tree.heading("cofins", text="Valor Cofins")

        tree.column("registro", width=80, anchor="center")
        tree.column("ind_oper", width=120, anchor="center")
        tree.column("cfop", width=100, anchor="center")
        tree.column("pis", width=200, anchor="e")
        tree.column("cofins", width=200, anchor="e")

        scrollbar = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for registro in registros_encontrados:
            tree.insert(
                "",
                tk.END,
                values=(
                    registro["registro"],
                    registro["ind_oper"],
                    registro["cfop"],
                    registro["pis"],
                    registro["cofins"],
                )
            )

        def exportar_xls():
            save_path = filedialog.asksaveasfilename(
                title="Exportar lista C170 - CFOP, PIS e Cofins",
                defaultextension=".xls",
                filetypes=[("Arquivo Excel", "*.xls")]
            )
            if not save_path:
                return

            try:
                df_exportacao = pd.DataFrame([
                    {
                        "REGISTRO": registro["registro"],
                        "ENTRADA / SAÍDA": registro["ind_oper"],
                        "CFOP": registro["cfop"],
                        "VALOR PIS": registro["pis"],
                        "VALOR COFINS": registro["cofins"],
                    }
                    for registro in registros_encontrados
                ])

                tabela_html = df_exportacao.to_html(index=False, border=1)
                conteudo_xls = (
                    "<html><head><meta charset=\"utf-8\"></head><body>"
                    f"{tabela_html}"
                    "</body></html>"
                )

                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(conteudo_xls)

                messagebox.showinfo("Sucesso", f"Lista exportada para:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar XLS: {e}")

        btn_frame = tk.Frame(relatorio_win)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        tk.Button(
            btn_frame,
            text="Exportar para XLS",
            command=exportar_xls,
            height=2
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def delete_duplicates(self):
        try:
            response = messagebox.askyesno("Excluir Duplicados", "Você deseja excluir as linhas duplicadas?")
            if not response:
                return
            self.data = [line for idx, line in enumerate(self.data, start=1) if f"{idx}" not in self.duplicates]
            self.display_data()
            messagebox.showinfo("Sucesso", "Duplicados excluídos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir duplicados: {e}")

    def delete_line(self):
        try:
            line_num = int(self.entry_line.get())
            if line_num < 1 or line_num > len(self.data):
                raise ValueError("Número de linha inválido.")
            self.deleted_lines.append(self.data[line_num - 1])
            self.data[line_num - 1] = None
            self.display_data()
            messagebox.showinfo("Sucesso", f"Linha {line_num} marcada para exclusão.")
        except ValueError as e:
            messagebox.showerror("Erro", str(e))

    def load_file(self, path=None):
        if not path:
            path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        
        if not path:
            return
            
        if getattr(self, 'is_modified', False):
            ans = messagebox.askyesnocancel("Salvar alterações", "Existem alterações não salvas no arquivo atual.\nDeseja salvar antes de carregar outro arquivo?")
            if ans is True:
                self.save_file()
            elif ans is None:
                return

        self.start_progress(f"Carregando arquivo {os.path.basename(path)}...")

        def _worker():
            try:
                with open(path, "r", encoding="latin-1") as file:
                    lines = file.readlines()
                self.root.after(0, lambda: self._on_file_loaded(path, lines))
            except FileNotFoundError:
                self.root.after(0, lambda: self._on_file_error(path, "Arquivo não encontrado."))
            except Exception as e:
                self.root.after(0, lambda: self._on_file_error(path, str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_file_loaded(self, path, lines):
        self.file_path = path
        self.data = lines
        self.deleted_lines = []
        self.inserted_lines = []
        self.clear_modified()
        self.add_to_recent(path)
        self.display_data()
        self.stop_progress(f"Arquivo carregado ({len(lines)} linhas): {path}")

    def _on_file_error(self, path, err_msg):
        self.stop_progress("Erro ao carregar arquivo")
        messagebox.showerror("Erro ao Carregar", f"Não foi possível abrir o arquivo:\n{err_msg}")
        if path in self.recent_files:
            self.recent_files.remove(path)
            self.save_recent_config()
            self.update_recent_menu()

    def display_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if not self.data:
            return

        filter_reg = getattr(self, "combo_registro_filter", None)
        filtro_selecionado = filter_reg.get().strip().upper() if filter_reg else "TODOS"

        max_cols = 0
        amostra_dados = []
        
        for idx, line in enumerate(self.data):
            if line is None:
                continue
            partes = line.strip().split('|')
            if len(partes) > 0 and partes[0] == '': partes.pop(0)
            if len(partes) > 0 and partes[-1] == '': partes.pop(-1)
            
            if filtro_selecionado != "TODOS":
                if not partes or partes[0].upper() != filtro_selecionado:
                    continue

            if len(partes) > max_cols:
                max_cols = len(partes)
            
            amostra_dados.append((idx + 1, partes))

        cols = ["#Linha"] + [f"C{i+1}" for i in range(max_cols)]
        
        self.tree["columns"] = cols
        self.tree.heading("#Linha", text="Linha")
        self.tree.column("#Linha", width=70, anchor="center", stretch=False)
        
        for c in cols[1:]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w", stretch=False)

        self.tree.tag_configure('odd', background='white')
        self.tree.tag_configure('even', background='#f8fafc')

        count = 0
        for num_linha, colunas_dados in amostra_dados:
            valores = [num_linha] + colunas_dados
            if len(valores) < len(cols):
                valores += [""] * (len(cols) - len(valores))
            
            tag = 'even' if count % 2 == 0 else 'odd'
            self.tree.insert("", "end", iid=str(num_linha), values=valores, tags=(tag,))
            count += 1
            
        self.update_window_title()
    def on_tree_select(self, event):
        """Ao clicar na tabela: Preenche edição E atualiza os títulos das colunas dinamicamente segundo o registro."""
        selected_item = self.tree.selection()
        if not selected_item:
            return
            
        item = self.tree.item(selected_item[0])
        valores = item['values']
        if not valores:
            return
        
        try:
            linha_num = int(valores[0])
        except (ValueError, IndexError):
            return

        self.entry_line.delete(0, tk.END)
        self.entry_line.insert(0, str(linha_num))
        
        # 1. Obtém o registro_tipo bruto de self.data para preservar zeros à esquerda (ex: 0000, 0150, 0200)
        registro_tipo = ""
        if 1 <= linha_num <= len(self.data):
            linha_raw = self.data[linha_num - 1]
            if linha_raw:
                partes = linha_raw.strip().split('|')
                if len(partes) > 0 and partes[0] == '':
                    partes.pop(0)
                if len(partes) > 0:
                    registro_tipo = partes[0].strip().upper()

        # Fallback com zfill caso valores[1] tenha sido convertido para int pelo Tkinter (ex: 0 -> 0000, 150 -> 0150)
        if not registro_tipo and len(valores) > 1:
            registro_tipo = str(valores[1]).strip().upper()
            if registro_tipo.isdigit() and len(registro_tipo) < 4:
                registro_tipo = registro_tipo.zfill(4)

        self.entry_value.delete(0, tk.END)
        if len(valores) > 2:
            self.entry_value.insert(0, f"{valores[1]} | {valores[2]}")
        elif len(valores) > 1:
            self.entry_value.insert(0, str(valores[1]))

        if registro_tipo:
            headers = self.layouts.get(registro_tipo, [])
            colunas_tree = self.tree["columns"]
            
            for i, col_id in enumerate(colunas_tree):
                if i == 0:
                    continue
                
                idx_header = i - 1
                if idx_header < len(headers):
                    novo_nome = headers[idx_header]
                else:
                    novo_nome = f"C{i}"
                
                self.tree.heading(col_id, text=novo_nome)

    def edit_value_tree(self):
        """Nova função de salvar edição compatível com a Tabela."""
        try:
            line_num = int(self.entry_line.get())
            col_num = int(self.entry_column.get()) # O usuário digita qual coluna quer editar (1, 2, 3...)
            new_val = self.entry_value.get()

            if line_num < 1 or line_num > len(self.data):
                messagebox.showerror("Erro", "Linha inválida")
                return

            # Acessa a linha original na memória
            raw_line = self.data[line_num - 1]
            parts = raw_line.strip().split('|')
            
            # Ajuste de índice: O SPED começa com pipe vazio, então a coluna 1 visual é o índice 1 do split
            # Ex: |C100|0|... -> split -> ['', 'C100', '0', ...]
            # Se o usuário pede coluna 1 (C100), é o index 1.
            
            if col_num < 1 or col_num >= len(parts):
                messagebox.showerror("Erro", "Coluna inválida para esta linha.")
                return

            # Atualiza o dado
            parts[col_num] = new_val
            
            # Reconstrói a linha
            self.data[line_num - 1] = "|".join(parts) + "\n"
            
            # Atualiza apenas visualmente ou tudo (tudo é mais seguro para alinhar)
            self.display_data()
            
            # Foca na linha editada
            self.tree.see(str(line_num))
            self.tree.selection_set(str(line_num))
            
            messagebox.showinfo("Sucesso", "Valor alterado.")

        except ValueError:
            messagebox.showerror("Erro", "Verifique se Linha e Coluna são números.")

    def goto_line_tree(self):
        """Vai para a linha digitada."""
        try:
            line_num = self.entry_search_line.get()
            if self.tree.exists(line_num):
                self.tree.see(line_num)
                self.tree.selection_set(line_num)
            else:
                messagebox.showwarning("Aviso", "Linha não encontrada.")
        except Exception:
            pass

    def search_in_tree(self):
        """Busca um valor textual na tabela."""
        term = self.entry_search.get().lower()
        if not term: return
        
        # Remove seleções anteriores
        for item in self.tree.selection():
            self.tree.selection_remove(item)
            
        found_any = False
        for item_id in self.tree.get_children():
            valores = self.tree.item(item_id)['values']
            # Varre todas as colunas da linha
            for v in valores:
                if term in str(v).lower():
                    self.tree.see(item_id)
                    self.tree.selection_add(item_id)
                    found_any = True
                    break # Achou na linha, vai pra próxima
        
        if not found_any:
            messagebox.showinfo("Busca", "Termo não encontrado.")

    def filter_data_tree(self):
        """Filtra a tabela principal visualmente."""
        try:
            col_idx = int(self.entry_filter_column.get())
            val_filter = self.entry_filter_value.get().lower()
            
            # Limpa tabela visual
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            # Reinsere só o que bate
            for idx, line in enumerate(self.data):
                parts = line.strip().split('|')
                # Ajuste de índice do pipe
                if len(parts) > col_idx:
                    if val_filter in parts[col_idx].lower():
                        # Insere
                        valores = [idx + 1] + parts[1:-1] # Ajuste simples para visualização
                        # (Simplificando a logica de inserção do display_data para o filtro rápido)
                        self.tree.insert("", "end", iid=str(idx+1), values=valores)
                        
        except ValueError:
            messagebox.showerror("Erro", "Coluna deve ser número.")
    def load_layouts(self):
        """Define os nomes das colunas para os registros do SPED Fiscal e Contribuições."""
        path_custom = self.get_custom_layout_path()
        if os.path.exists(path_custom):
            try:
                with open(path_custom, "r", encoding="utf-8") as f:
                    self.layouts = json.load(f)
                return
            except Exception as e:
                print(f"Erro ao carregar leiautes customizados: {e}")
        self.layouts = {
            "0000": ["REG", "COD_VER", "COD_FIN", "DT_INI", "DT_FIN", "NOME", "CNPJ", "CPF", "UF", "IE", "COD_MUN", "IM", "SUFRAMA", "IND_PERFIL", "IND_ATIV"],
            "0001": ["REG", "IND_MOV"],
            "0002": ["REG", "CLAS_ESTAB_IND"],
            "0005": ["REG", "FANTASIA", "CEP", "END", "NUM", "COMPL", "BAIRRO", "FONE", "FAX", "EMAIL"],
            "0015": ["REG", "UF_ST", "IE_ST"],
            "0100": ["REG", "NOME", "CPF", "CRC", "CNPJ", "CEP", "END", "NUM", "COMPL", "BAIRRO", "FONE", "FAX", "EMAIL"],
            "0150": ["REG", "COD_PART", "NOME", "COD_PAIS", "CNPJ", "CPF", "IE", "COD_MUN", "SUFRAMA", "END", "NUM", "COMPL", "BAIRRO"],
            "0175": ["REG", "DT_ALT", "NR_CAMPO", "CONT_ANT"],
            "0190": ["REG", "UNID", "DESCR"],
            "0200": ["REG", "COD_ITEM", "DESCR_ITEM", "COD_BARRA", "COD_ANT_ITEM", "UNID_INV", "TIPO_ITEM", "COD_NCM", "EX_IPI", "COD_GEN", "COD_LST", "ALIQ_ICMS", "CEST"],
            "0205": ["REG", "DESCR_ANT_ITEM", "DT_INI", "DT_FIM", "COD_ANT_ITEM"],
            "0206": ["REG", "COD_COMB"],
            "0210": ["REG", "COD_ITEM_COMP", "QTD_COMP", "PERDA"],
            "0220": ["REG", "UNID_CONV", "FAT_CONV", "COD_BARRA"],
            "0221": ["REG"],
            "0300": ["REG", "COD_IND_BEM", "IDENT_MERC", "DESCR_ITEM", "COD_PRNC", "COD_CTA", "NR_PARC"],
            "0305": ["REG", "COD_CCUS", "FUNC", "VIDA_UTIL"],
            "0400": ["REG", "COD_NAT", "DESCR_NAT"],
            "0450": ["REG", "COD_INF", "TXT"],
            "0460": ["REG", "COD_OBS", "TXT"],
            "0500": ["REG", "DT_ALT", "COD_", "IND_CTA", "COD_CTA", "NOME_CTA"],
            "0600": ["REG", "DT_ALT", "COD_CCUS", "CCUS"],
            "0990": ["REG", "QTD_LIN_0"],
            "1001": ["REG", "IND_MOV"],
            "1010": ["REG", "IND_EXP", "IND_CCRF", "IND_COMB", "IND_USINA", "IND_VA", "IND_EE", "IND_CART", "IND_FORM", "IND_AER", "IND_GIAF1", "IND_GIAF3", "IND_GIAF4"],
            "1100": ["REG", "IND_DOC", "NRO_DE", "DT_DE", "NAT_EXP", "NRO_RE", "DT_RE", "CHC_EMB", "DT_CHC", "DT_AVB", "TP_CHC", "PAIS"],
            "1105": ["REG", "COD_MOD", "SERIE", "NUM_DOC", "CHV_NFE", "DT_DOC", "COD_ITEM"],
            "1110": ["REG", "COD_PART", "COD_MOD", "NUM_DOC", "DT_DOC", "CHV_NFE", "NR_", "QTD", "UNID"],
            "1200": ["REG", "COD_AJ_APUR", "SLD_CRED", "CRED_APR", "CRED_RECEB", "CRED_UTIL", "SLD_CRED_FIM"],
            "1210": ["REG", "TIPO_UTIL", "NR_DOC", "VL_CRED_UTIL"],
            "1250": ["REG", "VL_CREDITO_ICMS_OP", "VL_ICMS_ST_REST", "VL_FCP_ST_REST", "VL_ICMS_ST_COMPL", "VL_FCP_ST_COMPL"],
            "1255": ["REG", "COD_MOT_REST_COMPL", "VL_CREDITO_ICMS_OP_MOT", "VL_ICMS_ST_REST_MOT", "VL_FCP_ST_REST_MOT", "VL_ICMS_ST_COMPL_MOT", "VL_FCP_ST_COMPL_MOT"],
            "1300": ["REG", "COD_ITEM", "DT_FECH", "ESTQ_ABERT", "VOL_ENTR", "VOL_DISP", "VOL_SAIDAS", "ESTQ_ESCR", "VAL_AJ_PERDA", "VAL_AJ_GANHO", "FECH_FISICO"],
            "1310": ["REG", "NUM_TANQUE", "ESTQ_ABERT", "VOL_ENTR", "VOL_DISP", "VOL_SAIDAS", "ESTQ_ESCR", "VAL_AJ_PERDA", "VAL_AJ_GANHO", "FECH_FISICO", "CAP_TANQUE"],
            "1320": ["REG", "NUM_BICO", "NR_INTERV", "MOT_INTERV", "NOM_INTERV", "CNPJ_INTERV", "CPF_INTERV", "VAL_FECHA", "VAL_ABERT", "VOL_AFERI", "VOL_VENDAS"],
            "1350": ["REG", "SERIE", "FABRICANTE", "MODELO", "TIPO_MEDICAO"],
            "1360": ["REG", "NUM_LACRE", "DT_APLICACAO"],
            "1370": ["REG", "NUM_BICO", "COD_ITEM", "NUM_TANQUE"],
            "1390": ["REG", "COD_PROD"],
            "1391": ["REG", "DT_REGISTRO", "QTD_MOID", "ESTQ_INI", "QTD_PRODUZ", "ENT_ANID_HID", "OUTR_ENTR", "PERDA", "CONS", "SAI_ANI_HID", "ESTQ_FIN", "ESTQ_INI_MEL", "PROD_DIA_MEL", "UTIL_MEL", "PROD_ALC_MEL", "OBS", "COD_ITEM", "TP_RESIDUO", "QTD_RESIDUO", "QTD_RESIDUO_"],
            "1400": ["REG", "COD_ITEM_IPM", "MUN", "VALOR"],
            "1500": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "SUB", "COD_CONS", "NUM_DOC", "DT_DOC", "DT_E_S", "VL_DOC", "VL_DESC", "VL_FORN", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "COD_INF", "VL_PIS", "VL_COFINS", "TP_LIGACAO", "COD_GRUPO_TENSAO"],
            "1510": ["REG", "NUM_ITEM", "COD_ITEM", "COD_CLASS", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST", "IND_REC", "COD_PART", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "1600": ["REG", "COD_PART", "TOT_CREDITO", "TOT_DEBITO"],
            "1601": ["REG", "COD_PART_IP", "COD_PART_IT", "TOT_VS", "TOT_ISS", "TOT_OUTROS"],
            "1700": ["REG", "COD_DISP", "COD_MOD", "SER", "SUB", "NUM_DOC_INI", "NUM_DOC_FIN", "NUM_AUT"],
            "1710": ["REG", "NUM_DOC_INI", "NUM_DOC_FIN"],
            "1800": ["REG", "VL_CARGA", "VL_PASS", "VL_FAT", "IND_RAT", "VL_ICMS_ANT", "VL_BC_ICMS", "VL_ICMS_APUR", "VL_BC_ICMS_APUR", "VL_DIF"],
            "1900": ["REG", "IND_APUR_ICMS", "DESCR_COMPL_OUT_APUR"],
            "1910": ["REG", "DT_INI", "DT_FIN"],
            "1920": ["REG", "VL_TOT_TRANSF_DEBITOS_OA", "VL_TOT_AJ_DEBITOS_OA", "VL_ESTORNOS_CRED_OA", "VL_TOT_TRANSF_CREDITOS_OA", "VL_TOT_AJ_CREDITOS_OA", "VL_ESTORNOS_DEB_OA", "VL_SLD_CREDOR_ANT_OA", "VL_SLD_APURADO_OA", "VL_TOT_DED", "VL_ICMS_RECOLHER_OA", "VL_SLD_CREDOR_TRANSP_OA", "DEB_ESP_OA"],
            "1921": ["REG", "COD_AJ_APUR", "DESCR_COMPL_AJ", "VL_AJ_APUR"],
            "1922": ["REG", "NUM_DA", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "1923": ["REG", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "COD_ITEM", "VL_AJ_ITEM"],
            "1925": ["REG", "COD_INF_ADIC", "VL_INF_ADIC", "DESCR_COMPL_AJ"],
            "1926": ["REG", "COD_OR", "VL_OR", "DT_VCTO", "COD_REC", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "1960": ["REG", "IND_AP", "G1_01", "G1_02", "G1_03", "G1_04", "G1_05", "G1_06", "G1_07", "G1_08", "G1_09", "G1_10", "G1_11"],
            "1970": ["REG", "IND_AP", "G3_01", "G3_02", "G3_03", "G3_04", "G3_05", "G3_06", "G3_07", "G3_08", "G3_09"],
            "1975": ["REG", "ALIQ_IMP_BASE", "G3_10", "G3_11", "G3_12"],
            "1980": ["REG", "IND_AP", "G4_01", "G4_02", "G4_03", "G4_04", "G4_05", "G4_06", "G4_07", "G4_08", "G4_09", "G4_10", "G4_11", "G4_12"],
            "1990": ["REG", "QTD_LIN_1"],
            "9001": ["REG", "IND_MOV"],
            "9900": ["REG", "REG_BLC", "QTD_REG_BLC"],
            "9990": ["REG", "QTD_LIN_9"],
            "9999": ["REG", "QTD_LIN", "D100", "D700", "C100", "C500"],
            "B001": ["REG", "IND_DAD"],
            "B020": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "NUM_DOC", "CHV_NFE", "DT_DOC", "VL_CONT", "VL_SUB", "VL_ISNT_ISS", "VL_DED_BC", "VL_BC_ISS", "VL_BC_ISS_RT", "VL_ISS_RT", "VL_", "COD_INF_OBS"],
            "B025": ["REG", "VL_CONT_P", "VL_BC_ISS_P", "ALIQ_ISS", "VL_ISS_P", "VL_ISNT_ISS_P", "COD_SERV"],
            "B030": ["REG", "COD_MOD", "SER", "NUM_DOC_INI", "NUM_DOC_FIN", "DT_DOC", "QTD_CANC", "VL_CONT", "VL_ISNT_ISS", "VL_BC_ISS", "VL_", "COD_INF_OBS"],
            "B035": ["REG", "VL_CONT_P", "VL_BC_ISS_P", "ALIQ_ISS", "VL_ISS_P", "VL_ISNT_ISS_P", "COD_SERV"],
            "B350": ["REG", "COD_CTD", "CTA_ISS", "CTA_COSIF", "QTD_OCOR", "COD_SERV", "VL_CONT", "VL_BC_ISS", "ALIQ_ISS", "VL_ISS", "COD_INF_OBS"],
            "B420": ["REG", "VL_CONT", "VL_BC_ISS", "ALIQ_ISS", "VL_ISNT_ISS", "VL_ISS", "COD_SERV"],
            "B440": ["REG", "IND_OPER", "COD_PART", "VL_CONT_RT", "VL_ISS_RT"],
            "B460": ["REG", "IND_DED", "VL_DED", "NUM_PROC", "IND_PROC", "PROC", "COD_INF_OBS", "IND_OBR"],
            "B470": ["REG", "VL_CONT", "VL_MAT_TERC", "VL_MAT_PROP", "VL_SUB", "VL_ISNT", "VL_DED_BC", "VL_BC_ISS", "VL_BC_ISS_RT", "VL_", "VL_ISS_RT", "VL_DED", "VL_ISS_REC_UNI"],
            "B500": ["REG", "VL_REC", "QTD_PROF", "VL_OR"],
            "B510": ["REG", "IND_PROF", "IND_ESC", "IND_SOC", "CPF", "NOME"],
            "B990": ["REG", "QTD_LIN_B"],
            "C001": ["REG", "IND_MOV"],
            "C100": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "NUM_DOC", "CHV_NFE", "DT_DOC", "DT_E_S", "VL_DOC", "IND_PGTO", "VL_DESC", "VL_ABAT_NT", "VL_MERC", "IND_FRT", "VL_FRT", "VL_SEG", "VL_OUT_DA", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_IPI", "VL_PIS", "VL_COFINS", "VL_PIS_ST", "VL_COFINS_ST"],
            "C101": ["REG", "VL_FCP_UF_DEST", "VL_ICMS_UF_DEST", "VL_ICMS_UF_REM"],
            "C105": ["REG", "OPER", "UF"],
            "C110": ["REG", "COD_INF", "TXT_COMPL"],
            "C111": ["REG", "NUM_PROC", "IND_PROC"],
            "C112": ["REG", "COD_DA", "UF", "NUM_DA", "COD_AUT", "VL_DA", "DT_VCTO", "DT_PGTO"],
            "C113": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC"],
            "C114": ["REG", "COD_MOD", "ECF_FAB", "ECF_CX", "NUM_DOC", "DT_DOC"],
            "C115": ["REG", "IND_CARGA", "CNPJ_COL", "IE_COL", "CPF_COL", "COD_MUN_COL", "CNPJ_ENTG", "IE_ENTG", "CPF_ENTG"],
            "C116": ["REG", "COD_MOD", "NR_SAT", "CHV_CFE", "NUM_CFE", "DT_DOC"],
            "C120": ["REG", "COD_DOC_IMP", "NUM_DOC_IMP", "PIS_IMP", "COFINS", "NUM_ACDRAW"],
            "C130": ["REG", "VL_SERV_NT", "VL_BC_ISSQN", "VL_ISSQN", "VL_BC_IRRF", "VL_", "VL_BC_PREV"],
            "C140": ["REG", "IND_EMIT", "IND_TIT", "DESC_TIT", "NUM_TIT", "QTD_PARC", "VL_TIT"],
            "C141": ["REG", "NUM_PARC", "DT_VCTO", "VL_PARC"],
            "C160": ["REG", "COD_PART", "VEIC_ID", "QTD_VOL", "PESO_BRT", "PESO_LIQ", "UF_ID"],
            "C165": ["REG", "VEIC_ID", "COD_AUT", "NR_PASSE", "HORA", "TEMPER", "QTD_VOL", "PESO_BRT", "PESO_LIQ", "NOM_MOT", "CPF", "UF_ID"],
            "C170": ["REG", "NUM_ITEM", "COD_ITEM", "DESCR_COMPL", "QTD", "UNID", "VL_ITEM", "VL_DESC", "IND_MOV", "CST_ICMS", "CFOP", "COD_NAT", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "ALIQ_ST", "VL_ICMS_ST", "IND_APUR", "CST_IPI", "COD_ENQ", "VL_BC_IPI", "ALIQ_IPI", "VL_IPI", "CST_PIS", "VL_BC_PIS", "ALIQ_PIS", "QUANT_BC_PIS", "VL_PIS", "CST_COFINS", "VL_BC_COFINS", "ALIQ_COFINS", "VL_COFINS", "COD_CTA", "VL_ABAT_NT"],
            "C171": ["REG", "NUM_TANQUE", "QTDE"],
            "C172": ["REG", "VL_BC_ISSQN", "ALIQ_ISSQN", "VL_ISSQN"],
            "C173": ["REG", "LOTE_MED", "QTD_ITEM", "DT_FAB", "DT_VAL", "IND_MED", "TP_PROD", "VL_TAB_MAX"],
            "C174": ["REG", "IND_ARM", "NUM_ARM", "DESCR_COMPL"],
            "C175": ["REG", "IND_VEIC_OPER", "CNPJ", "UF", "CHASSI_VEIC"],
            "C176": ["REG", "COD_MOD_ULT_E", "NUM_DOC_ULT_E", "SER_ULT_E", "DT_ULT_E", "COD_PART_ULT_E", "QUANT_ULT_E", "VL_UNIT_ULT_E", "VL_UNIT_BC_ST", "NUM_ITEM_ULT_E", "ALIQ_ICMS_ULT_E", "ALIQ_ST_ULT_E", "VL_UNIT_RES", "COD_RESP_RET", "COD_MOT_RES", "CHAVE_NFE_RET", "SER_NFE_RET", "NUM_NFE_RET", "ITEM_NFE_RET", "COD_DA", "NUM_DA"],
            "C177": ["REG", "COD_SELO_IPI", "QT_SELO_IPI"],
            "C178": ["REG", "CL_ENQ", "VL_UNID", "QUANT_PAD"],
            "C179": ["REG", "BC_ST_ORIG_DEST", "ICMS_ST_REP", "ICMS_ST_COMPL", "BC_RET", "ICMS_RET"],
            "C180": ["REG", "COD_RESP_RET", "QUANT_CONV", "UNID", "VL_UNIT_CONV", "VL_UNIT_ICMS_OP_CONV", "VL_UNIT_BC_ICMS_ST", "VL_UNIT_ICMS_ST_CONV", "VL_UNIT_FCP_ST_CONV", "COD_DA", "NUM_DA"],
            "C181": ["REG", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "COD_MOD_SAIDA", "SERIE_SAIDA", "ECF_FAB_SAIDA", "NUM_DOC_SAIDA", "CHV_DFE_SAIDA", "DT_DOC_SAIDA", "NUM_ITEM_SAIDA", "VL_UNIT_CONV_SAIDA"],
            "C185": ["REG", "NUM_ITEM", "COD_ITEM", "CST_ICMS", "CFOP", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "VL_UNIT_CONV", "VL_UNIT_ICMS_OP_CONV"],
            "C186": ["REG", "NUM_ITEM", "COD_ITEM", "CST_ICMS", "CFOP", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "COD_MOD_ENTRADA", "SERIE_ENTRADA", "NUM_DOC_ENTRADA", "CHV_DFE_ENTRADA", "DT_DOC_ENTRADA", "NUM_ITEM_ENTRADA", "VL_UNIT_CONV_ENTRADA", "VL_UNIT_BC_ICMS_ST"],
            "C190": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_ICMS_ST", "VL_RED_BC", "VL_IPI", "COD_OBS"],
            "C191": ["REG"],
            "C195": ["REG", "COD_OBS", "TXT_COMPL"],
            "C197": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "C300": ["REG", "COD_MOD", "SER", "SUB", "NUM_DOC_INI", "NUM_DOC_FIN", "DT_DOC", "VL_DOC", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "C310": ["REG", "NUM_DOC_CANC"],
            "C320": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "C321": ["REG", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "VL_DESC", "VL_BC_ICMS", "VL_ICMS", "VL_PIS", "VL_COFINS"],
            "C330": ["REG", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "VL_UNIT_CONV"],
            "C350": ["REG", "SER", "SUB_SER", "NUM_DOC", "DT_DOC", "CNPJ_CPF", "VL_MERC", "VL_DOC", "VL_DESC", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "C370": ["REG", "NUM_ITEM", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "VL_DESC"],
            "C380": ["REG", "QUANT_CONV", "UNID", "VL_UNIT_CONV", "CST_ICMS", "CFOP"],
            "C390": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "C400": ["REG", "COD_MOD", "ECF_MOD", "ECF_FAB", "ECF_CX"],
            "C405": ["REG", "DT_DOC", "CRO", "CRZ", "GT_FIN", "VL_BRT"],
            "C410": ["REG", "VL_PIS", "VL_COFINS"],
            "C420": ["REG", "COD_TOT_PAR", "VLR_ACUM_TOT", "NR_TOT", "DESCR_NR_TOT"],
            "C425": ["REG", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "VL_PIS", "VL_COFINS"],
            "C430": ["REG", "QUANT_CONV", "UNID", "VL_UNIT_CONV", "CST_ICMS", "CFOP"],
            "C460": ["REG", "COD_MOD", "COD_SIT", "NUM_DOC", "DT_DOC", "VL_DOC", "VL_PIS", "VL_COFINS", "CPF_CNPJ", "NOM_ADQ"],
            "C465": ["REG", "CHV_CFE", "NUM_CCF"],
            "C470": ["REG", "COD_ITEM", "QTD", "QTD_CANC", "UNID", "VL_ITEM", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_PIS", "VL_COFINS"],
            "C480": ["REG", "QUANT_CONV", "UNID", "VL_UNIT_CONV", "CST_ICMS", "CFOP"],
            "C490": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "COD_OBS"],
            "C495": ["REG", "ALIQ_ICMS", "COD_ITEM", "QTD", "QTD_CANC", "UNID", "VL_ITEM", "VL_DESC", "VL_CANC", "VL_ACMO", "VL_BC_ICMS", "VL_ICMS", "VL_ISEN", "VL_NT", "VL_ICMS_ST"],
            "C500": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "SUB", "COD_CONS", "NUM_DOC", "DT_DOC", "DT_E_S", "VL_DOC", "VL_DESC", "VL_FORN", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "COD_INF", "VL_PIS", "VL_COFINS", "TP", "COD_GRUPO_TENSAO", "COD_MOD_DOC_REF", "HASH_DOC_REF", "SER_DOC_REF", "NUM_DOC_REF", "MES_DOC_REF", "ENER_INJET", "OUTRAS_DED"],
            "C510": ["REG", "NUM_ITEM", "COD_ITEM", "COD_CLASS", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "ALIQ_ST", "VL_ICMS_ST", "IND_REC", "COD_PART", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "C590": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_RED_BC", "COD_OBS"],
            "C591": ["REG", "VL_FCP_OP", "VL_FCP_ST"],
            "C595": ["REG", "COD_OBS", "TXT_COMPL"],
            "C597": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "C600": ["REG", "COD_MOD", "COD_MUN", "SER", "SUB", "COD_CONS", "QTD_CONS", "QTD_CANC", "DT_DOC", "VL_DOC", "VL_DESC", "CONS", "VL_FORN", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_PIS", "VL_COFINS"],
            "C601": ["REG", "NUM_DOC_CANC"],
            "C610": ["REG", "COD_CLASS", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "C690": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "VL_BC_ICMS_ST", "VL_ICMS_ST", "COD_OBS"],
            "C700": ["REG", "COD_MOD", "SER", "NRO_ORD_INI", "NRO_ORD_FIN", "DT_DOC_INI", "DT_DOC_FIN", "NOM_MEST", "CHV_COD_DIG"],
            "C790": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_RED_BC", "COD_OBS"],
            "C791": ["REG", "UF", "VL_BC_ICMS_ST", "VL_ICMS_ST"],
            "C800": ["REG", "COD_MOD", "COD_SIT", "NUM_CFE", "DT_DOC", "VL_CFE", "VL_PIS", "VL_COFINS", "CNPJ_CPF", "NR_SAT", "CHV_CFE", "VL_DESC", "VL_MERC", "VL_OUT_DA", "VL_ICMS", "VL_PIS_ST", "VL_COFINS_ST"],
            "C810": ["REG", "NUM_ITEM", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "CST_ICMS", "CFOP"],
            "C815": ["REG", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "VL_UNIT_CONV"],
            "C850": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "COD_OBS"],
            "C855": ["REG", "COD_OBS", "TXT_COMPL"],
            "C857": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "C860": ["REG", "COD_MOD", "NR_SAT", "DT_DOC", "DOC_INI", "DOC_FIM"],
            "C870": ["REG", "COD_ITEM", "QTD", "UNID", "CST_ICMS", "CFOP"],
            "C880": ["REG", "COD_MOT_REST_COMPL", "QUANT_CONV", "UNID", "VL_UNIT_CONV"],
            "C890": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "COD_OBS"],
            "C895": ["REG", "COD_OBS", "TXT_COMPL"],
            "C897": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "C990": ["REG", "QTD_LIN_C"],
            "D001": ["REG", "IND_MOV"],
            "D100": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "SUB", "NUM_DOC", "CHV_CTE", "DT_DOC", "DT_A_P", "CHV_CTE_REF", "VL_DOC", "VL_DESC", "IND_FRT", "VL_SERV", "VL_BC_ICMS", "VL_ICMS", "VL_NT", "COD_INF", "COD_CTA", "COD_MUN_ORIG", "COD_MUN_DEST"],
            "D101": ["REG", "VL_FCP_UF_DEST", "VL_ICMS_UF_REM"],
            "D110": ["REG", "NUM_ITEM", "COD_ITEM", "VL_SERV", "VL_OUT"],
            "D120": ["REG", "COD_MUN_ORIG", "COD_MUN_DEST", "VEIC_ID", "UF_ID"],
            "D130": ["REG", "COD_PART_CONSG", "COD_PART_RED", "IND_FRT_RED", "COD_MUN_ORIG", "COD_MUN_DEST", "VEIC_ID", "VL_LIQ_FRT", "VL_SEC_CAT", "VL_DESP", "VL_PEDG", "VL_OUT", "VL_FRT", "UF_ID"],
            "D140": ["REG", "COD_PART_CONSG", "COD_MUN_ORIG", "COD_MUN_DEST", "IND_VEIC", "VEIC_ID", "IND_NAV", "VIAGEM", "VL_FRT_LIQ", "VL_DESP_PORT", "VL_OUT", "VL_FRT_BRT", "VL_FRT_MM"],
            "D150": ["REG", "COD_MUN_ORIG", "COD_MUN_DEST", "VEIC_ID", "VIAGEM", "IND_TFA", "VL_PESO_TX", "VL_TX_TERR", "VL_TX_RED", "VL_OUT", "VL_TX_ADV"],
            "D160": ["REG", "DESPACHO", "CNPJ_CPF_REM", "IE_REM", "COD_MUN_ORI", "CNPJ_CPF_DEST", "IE_DEST", "COD_MUN_DEST"],
            "D161": ["REG", "IND_CARGA", "CNPJ_CPF_COL", "IE_COL", "COD_MUN_COL", "CNPJ_CPF_ENTG", "IE_ENTG", "COD_MUN_ENTG"],
            "D162": ["REG", "COD_MOD", "SER", "NUM_DOC", "DT_DOC", "VL_DOC", "VL_MERC", "QTD_VOL", "PESO_BRT", "PESO_LIQ"],
            "D170": ["REG", "COD_PART_CONSG", "COD_PART_RED", "COD_MUN_ORIG", "COD_MUN_DEST", "OTM", "IND_NAT_FRT", "VL_LIQ_FRT", "VL_GRIS", "VL_PDG", "VL_OUT", "VL_FRT", "VEIC_ID", "UF_ID"],
            "D180": ["REG", "NUM_SEQ", "IND_EMIT", "CNPJ_CPF_EMIT", "UF_EMIT", "IE_EMIT", "COD_MUN_ORIG", "CNPJ_CPF_TOM", "UF_TOM", "IE_TOM", "COD_MUN_DEST", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "VL_DOC"],
            "D190": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "D195": ["REG", "COD_OBS", "TXT_COMPL"],
            "D197": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "D300": ["REG", "COD_MOD", "SER", "SUB", "NUM_DOC_INI", "CST_ICMS", "CFOP", "ALIQ_ICMS", "DT_DOC", "VL_OPR", "VL_DESC", "VL_SERV", "VL_SEG", "VL_OUT", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS", "COD_CTA"],
            "D301": ["REG", "NUM_DOC_CANC"],
            "D310": ["REG", "COD_MUN_ORIG", "VL_SERV", "VL_BC_ICMS", "VL_ICMS"],
            "D350": ["REG", "COD_MOD", "ECF_MOD", "ECF_FAB", "ECF_CX"],
            "D355": ["REG", "DT_DOC", "CRO", "CRZ", "NUM_COO_FIN", "GT_FIN", "VL_BRT"],
            "D360": ["REG", "VL_PIS", "VL_COFINS"],
            "D365": ["REG", "COD_TOT_PAR", "VLR_ACUM_TOT", "NR_TOT", "DESCR_NR_TOT"],
            "D370": ["REG", "COD_MUN_ORIG", "VL_SERV", "QTD_BILH", "VL_BC_ICMS", "VL_ICMS"],
            "D390": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ISSQN", "ALIQ_ISSQN", "VL_ISSQN", "VL_BC_ICMS", "VL_ICMS", "COD_OBS"],
            "D400": ["REG", "COD_PART", "COD_MOD", "COD_SIT", "SER", "SUB", "NUM_DOC", "DT_DOC", "VL_DOC", "VL_DESC", "VL_SERV", "VL_ICMS", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "D410": ["REG", "COD_MOD", "SER", "SUB", "NUM_DOC_INI", "NUM_DOC_FIN", "DT_DOC", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_DESC", "VL_SERV", "VL_BC_ICMS", "VL_ICMS"],
            "D411": ["REG", "NUM_DOC_CANC"],
            "D420": ["REG", "COD_MUN_ORIG", "VL_SERV", "VL_BC_ICMS", "VL_ICMS"],
            "D500": ["REG", "IND_OPER", "IND_EMIT", "COD_PART", "COD_MOD", "COD_SIT", "SER", "SUB", "NUM_DOC", "DT_DOC", "DT_A_P", "VL_DOC", "VL_DESC", "VL_SERV", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "COD_INF", "VL_PIS", "VL_COFINS", "COD_CTA", "TP_ASSINANTE"],
            "D510": ["REG", "NUM_ITEM", "COD_ITEM", "COD_CLASS", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_BC_ICMS_UF", "VL_ICMS_UF", "IND_REC", "COD_PART", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "D530": ["REG", "IND_SERV", "DT_INI_SERV", "DT_FIN_SERV", "PER_FISCAL", "COD_AREA", "TERMINAL"],
            "D590": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_UF", "VL_ICMS_UF", "VL_RED_BC", "COD_OBS"],
            "D600": ["REG", "COD_MOD", "COD_MUN", "SER", "SUB", "COD_CONS", "QTD_CONS", "DT_DOC", "VL_DOC", "VL_DESC", "VL_SERV", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "VL_PIS", "VL_COFINS"],
            "D610": ["REG", "COD_CLASS", "COD_ITEM", "QTD", "UNID", "VL_ITEM", "VL_DESC", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_BC_ICMS", "VL_ICMS", "VL_BC_ICMS_UF", "VL_ICMS_UF", "VL_RED_BC", "VL_PIS", "VL_COFINS", "COD_CTA"],
            "D690": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "D695": ["REG", "COD_MOD", "SER", "NRO_ORD_INI", "NRO_ORD_FIN", "DT_DOC_INI", "DT_DOC_FIN", "NOM_MEST", "CHV_COD_DIG"],
            "D696": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_BC_ICMS", "VL_ICMS", "VL_ICMS_UF", "VL_RED_BC", "COD_OBS"],
            "D697": ["REG", "UF", "VL_BC_ICMS", "VL_ICMS"],
            "D700": ["REG", "DT_E_S", "VL_DOC", "VL_DESC", "VL_SERV", "VL_SERV_NT", "VL_TERC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "COD_INF", "VL_PIS", "VL_COFINS", "TIP_FAT", "COD_MOD_DOC_REF", "HASH_DOC_REF", "SER_DOC_REF", "NUM_DOC_REF", "MES_DOC_REF", "COD_MUN_DEST", "DED"],
            "D730": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "D731": ["REG", "VL_FCP_OP"],
            "D735": ["REG", "COD_OBS", "TXT_COMPL"],
            "D737": ["REG", "COD_AJ", "DESCR_COMPL_AJ", "COD_ITEM", "VL_BC_ICMS", "ALIQ_ICMS", "VL_ICMS", "VL_OUTROS"],
            "D750": ["REG", "COD_MOD", "SER", "DT_DOC", "QTD_CONS", "IND_PREPAGO", "VL_DOC", "VL_SERV", "VL_SERV_NT", "VL_TERC", "VL_DESC", "VL_DA", "VL_BC_ICMS", "VL_ICMS", "VL_PIS", "VL_COFINS", "DED"],
            "D760": ["REG", "CST_ICMS", "CFOP", "ALIQ_ICMS", "VL_OPR", "VL_BC_ICMS", "VL_ICMS", "VL_RED_BC", "COD_OBS"],
            "D761": ["REG", "VL_FCP_OP"],
            "D990": ["REG", "QTD_LIN_D"],
            "E001": ["REG", "IND_MOV"],
            "E100": ["REG", "DT_INI", "DT_FIN"],
            "E110": ["REG", "VL_TOT_DEBITOS", "VL_AJ_DEBITOS", "VL_TOT_AJ_DEBITOS", "VL_ESTORNOS_CRED", "VL_TOT_CREDITOS", "VL_AJ_CREDITOS", "VL_TOT_AJ_CREDITOS", "VL_ESTORNOS_DEB", "VL_SLD_CREDOR_ANT", "VL_SLD_APURADO", "VL_TOT_DED", "VL_ICMS_RECOLHER", "DEB_ESP"],
            "E111": ["REG", "COD_AJ_APUR", "DESCR_COMPL_AJ", "VL_AJ_APUR"],
            "E112": ["REG", "NUM_DA", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E113": ["REG", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "COD_ITEM", "VL_AJ_ITEM"],
            "E115": ["REG", "COD_INF_ADIC", "VL_INF_ADIC", "DESCR_COMPL_AJ"],
            "E116": ["REG", "COD_OR", "VL_OR", "DT_VCTO", "COD_REC", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E200": ["REG", "UF", "DT_INI", "DT_FIN"],
            "E210": ["REG", "IND_MOV_ST", "VL_SLD_CRED_ANT_ST", "VL_DEVOL_ST", "VL_RESSARC_ST", "VL_OUT_CRED_ST", "VL_AJ_CREDITOS_ST", "VL_OUT_DEB_ST", "VL_AJ_DEBITOS_ST", "VL_SLD_DEV_ANT_ST", "VL_ICMS_RECOL_ST", "DEB_ESP_ST"],
            "E220": ["REG", "COD_AJ_APUR", "DESCR_COMPL_AJ", "VL_AJ_APUR"],
            "E230": ["REG", "NUM_DA", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E240": ["REG", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "COD_ITEM", "VL_AJ_ITEM"],
            "E250": ["REG", "COD_OR", "VL_OR", "DT_VCTO", "COD_REC", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E300": ["REG", "UF", "DT_INI", "DT_FIN"],
            "E310": ["REG", "IND_MOV_DIFAL", "VL_SLD_CRED_ANT_DIFAL", "VL_TOT_DEBITOS_DIFAL", "VL_OUT_DEB_DIFAL", "VL_TOT_DEB_FCP", "VL_TOT_CREDITOS_DIFAL", "VL_TOT_CRED_FCP", "VL_OUT_CRED_DIFAL", "VL_SLD_DEV_ANT_DIFAL", "VL_RECOL", "DEB_ESP_DIFAL", "IND_MOV_FCP_DIFAL", "VL_RECOL_DIFAL", "VL_SLD_CRED_ANT_FCP", "VL_OUT_DEB_FCP", "VL_OUT_CRED_FCP", "VL_SLD_DEV_ANT_FCP", "VL_RECOL_FCP", "DEB_ESP_FCP"],
            "E311": ["REG", "COD_AJ_APUR", "DESCR_COMPL_AJ", "VL_AJ_APUR"],
            "E312": ["REG", "NUM_DA", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E313": ["REG", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "COD_ITEM", "VL_AJ_ITEM"],
            "E316": ["REG", "COD_OR", "VL_OR", "DT_VCTO", "COD_REC", "NUM_PROC", "IND_PROC", "PROC", "TXT_COMPL"],
            "E500": ["REG", "IND_APUR", "DT_INI", "DT_FIN"],
            "E510": ["REG", "CFOP", "CST_IPI", "VL_CONT_IPI", "VL_BC_IPI", "VL_IPI"],
            "E520": ["REG", "VL_SD_ANT_IPI", "VL_DEB_IPI", "VL_CRED_IPI", "VL_OD_IPI", "VL_OC_IPI", "VL_SC_IPI", "VL_SD_IPI"],
            "E530": ["REG", "IND_AJ", "VL_AJ", "COD_AJ", "IND_DOC", "NUM_DOC", "DESCR_AJ"],
            "E531": ["REG", "COD_PART", "COD_MOD", "SER", "SUB", "NUM_DOC", "DT_DOC", "COD_ITEM", "VL_AJ_ITEM", "CHV_NFE"],
            "E990": ["REG", "QTD_LIN_E"],
            "G001": ["REG", "IND_MOV"],
            "G110": ["REG", "DT_INI", "DT_FIN", "SALDO_IN_ICMS", "SOM_PARC", "VL_TRIB_EXP", "VL_TOTAL", "IND_PER_SAI", "ICMS_APROP", "SOM_ICMS_OC"],
            "G125": ["REG", "COD_IND_BEM", "DT_MOV", "TIPO_MOV", "NUM_PARC", "VL_PARC_PASS"],
            "G126": ["REG", "DT_INI", "DT_FIM", "NUM_PARC", "VL_PARC_PASS", "VL_TRIB_OC", "VL_TOTAL", "IND_PER_SAI", "VL_PARC_APROP"],
            "G130": ["REG", "IND_EMIT", "COD_PART", "COD_MOD", "SERIE", "NUM_DOC", "CHV_NFE_CTE", "DT_DOC", "NUM_DA"],
            "G140": ["REG", "NUM_ITEM", "COD_ITEM", "QTDE", "UNID", "VL_ICMS_OP_APLICADO", "VL_ICMS_ST_APLICADO", "VL_ICMS_FRT_APLICADO", "VL_ICMS_DIF_APLICADO"],
            "G990": ["REG", "QTD_LIN_G"],
            "H001": ["REG", "IND_MOV"],
            "H005": ["REG", "DT_INV", "VL_INV", "MOT_INV"],
            "H010": ["REG", "COD_ITEM", "UNID", "QTD", "VL_UNIT", "VL_ITEM", "IND_PROP", "COD_PART", "TXT_COMPL", "COD_CTA", "VL_ITEM_IR"],
            "H020": ["REG", "CST_ICMS", "BC_ICMS", "VL_ICMS"],
            "H030": ["REG", "VL_ICMS_OP", "VL_BC_ICMS_ST", "VL_ICMS_ST", "VL_FCP"],
            "H990": ["REG", "QTD_LIN_H"],
            "K001": ["REG", "IND_MOV"],
            "K010": ["REG", "IND_TP_LEIAUTE"],
            "K100": ["REG", "DT_INI", "DT_FIN"],
            "K200": ["REG", "DT_EST", "COD_ITEM", "QTD", "IND_EST", "COD_PART"],
            "K210": ["REG", "DT_INI_OS", "DT_FIN_OS", "COD_DOC_OS", "COD_ITEM_ORI", "QTD_ORI"],
            "K215": ["REG", "COD_ITEM_DES", "QTD_DES"],
            "K220": ["REG", "DT_MOV", "COD_ITEM_ORI", "COD_ITEM_DEST", "QTD_ORI", "QTD_DEST"],
            "K230": ["REG", "DT_INI_OP", "DT_FIN_OP", "COD_DOC_OP", "COD_ITEM", "QTD_ENC"],
            "K235": ["REG", "COD_ITEM", "QTD"],
            "K250": ["REG", "DT_PROD", "COD_ITEM", "QTD"],
            "K255": ["REG", "DT_CONS", "COD_ITEM", "QTD"],
            "K260": ["REG", "COD_OP_OS", "COD_ITEM", "DT_RET", "QTD_RET"],
            "K265": ["REG", "COD_ITEM", "QTD_CONS", "QTD_RET"],
            "K270": ["REG", "DT_INI_AP", "DT_FIN_AP", "COD_OP_OS", "COD_ITEM", "QTD_COR_POS", "QTD_COR_NEG", "ORIGEM"],
            "K275": ["REG", "COD_ITEM", "COD_INS_SUBST"],
            "K280": ["REG", "COD_ITEM", "QTD_COR_POS", "QTD_COR_NEG", "IND_EST", "COD_PART"],
            "K290": ["REG", "DT_INI_OP", "DT_FIN_OP", "COD_DOC_OP"],
            "K291": ["REG", "COD_ITEM", "QTD"],
            "K292": ["REG", "COD_ITEM", "QTD"],
            "K300": ["REG", "DT_PROD"],
            "K301": ["REG", "COD_ITEM", "QTD"],
            "K302": ["REG", "COD_ITEM", "QTD"],
            "K990": ["REG", "QTD_LIN_K"],
        }


    def get_custom_layout_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "sped_layouts_custom.json")

    def save_custom_layouts(self):
        """Salva a estrutura atual de leiautes em arquivo JSON local."""
        try:
            path = self.get_custom_layout_path()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.layouts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Erro ao salvar leiautes customizados: {e}")

    def abrir_gerenciador_layouts(self):
        """Abre a janela de Gerenciamento de Leiautes dos Registros SPED."""
        win = tk.Toplevel(self.root)
        win.title("Gerenciador de Leiautes de Registros SPED")
        win.geometry("950x600")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=10)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Gerenciador de Leiautes de Registros SPED", font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        frame_search = tk.Frame(win, pady=8, padx=10)
        frame_search.pack(fill=tk.X)

        tk.Label(frame_search, text="Pesquisar Registro (ex: C100, 0200):", font=("Calibri", 10, "bold")).pack(side=tk.LEFT, padx=5)
        entry_search_reg = tk.Entry(frame_search, width=15, font=("Consolas", 10))
        entry_search_reg.pack(side=tk.LEFT, padx=5)

        frame_table = tk.Frame(win)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("reg", "qtd", "campos")
        tree_layouts = ttk.Treeview(frame_table, columns=cols, show="headings", selectmode="browse")
        tree_layouts.heading("reg", text="Registro")
        tree_layouts.heading("qtd", text="Qtd Campos")
        tree_layouts.heading("campos", text="Campos (Nomenclatura do Cabeçalho)")

        tree_layouts.column("reg", width=90, anchor="center")
        tree_layouts.column("qtd", width=100, anchor="center")
        tree_layouts.column("campos", width=700, anchor="w")

        vsb = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree_layouts.yview)
        tree_layouts.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree_layouts.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def preencher_tabela_layouts(filtro=""):
            for item in tree_layouts.get_children():
                tree_layouts.delete(item)
            
            filtro_upper = filtro.strip().upper()
            count = 0
            for reg, campos in sorted(self.layouts.items()):
                if filtro_upper and filtro_upper not in reg.upper():
                    continue
                tag = 'even' if count % 2 == 0 else 'odd'
                campos_str = " | ".join(campos)
                tree_layouts.insert("", tk.END, iid=reg, values=(reg, len(campos), campos_str), tags=(tag,))
                count += 1

        tree_layouts.tag_configure('odd', background='white')
        tree_layouts.tag_configure('even', background='#f8fafc')

        preencher_tabela_layouts()

        entry_search_reg.bind("<KeyRelease>", lambda e: preencher_tabela_layouts(entry_search_reg.get()))

        def acao_editar_ou_cadastrar():
            sel = tree_layouts.selection()
            reg_inicial = sel[0] if sel else ""
            campos_iniciais = " | ".join(self.layouts.get(reg_inicial, ["REG"])) if reg_inicial else "REG, Campo2, Campo3"

            dlg = tk.Toplevel(win)
            dlg.title(f"Cadastrar / Editar Registro SPED" if not reg_inicial else f"Editar Registro: {reg_inicial}")
            dlg.geometry("550x300")
            dlg.transient(win)
            dlg.grab_set()

            tk.Label(dlg, text="Código do Registro (4 caracteres):", font=("Calibri", 10, "bold")).pack(anchor="w", padx=15, pady=(15, 2))
            ent_reg = tk.Entry(dlg, width=15, font=("Consolas", 11))
            ent_reg.pack(anchor="w", padx=15, pady=2)
            if reg_inicial:
                ent_reg.insert(0, reg_inicial)
                ent_reg.config(state="disabled")

            tk.Label(dlg, text="Campos do Registro (separados por vírgula ou |):", font=("Calibri", 10, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
            txt_campos = scrolledtext.ScrolledText(dlg, width=60, height=6, font=("Consolas", 9))
            txt_campos.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
            txt_campos.insert(tk.END, campos_iniciais)

            def salvar_registro_dlg():
                cod_reg = ent_reg.get().strip().upper()
                if not cod_reg or len(cod_reg) != 4:
                    messagebox.showerror("Erro", "O código do registro deve ter exatamente 4 caracteres (ex: C100, 0200).")
                    return
                raw_campos = txt_campos.get("1.0", tk.END).strip()
                if "|" in raw_campos:
                    lista_campos = [c.strip() for c in raw_campos.split("|") if c.strip()]
                else:
                    lista_campos = [c.strip() for c in raw_campos.split(",") if c.strip()]

                if not lista_campos:
                    messagebox.showerror("Erro", "Informe ao menos um campo para o registro.")
                    return
                
                if lista_campos[0].upper() != "REG":
                    lista_campos.insert(0, "REG")

                self.layouts[cod_reg] = lista_campos
                self.save_custom_layouts()
                preencher_tabela_layouts(entry_search_reg.get())
                dlg.destroy()
                messagebox.showinfo("Sucesso", f"Registro {cod_reg} salvo com sucesso ({len(lista_campos)} campos)!")

            tk.Button(dlg, text="Salvar Registro", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), pady=5, command=salvar_registro_dlg).pack(pady=10)

        frame_bottom_btns = tk.Frame(win, pady=10)
        frame_bottom_btns.pack(fill=tk.X)

        tk.Button(frame_bottom_btns, text="Cadastrar / Editar Registro", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=acao_editar_ou_cadastrar, padx=10, pady=5).pack(side=tk.LEFT, padx=15)

        def restaurar_padroes():
            if messagebox.askyesno("Restaurar Padrões", "Deseja restaurar os 270 leiautes padrões do Guia Prático EFD v3.2.2?"):
                path_custom = self.get_custom_layout_path()
                if os.path.exists(path_custom):
                    try:
                        os.remove(path_custom)
                    except Exception:
                        pass
                # Recarrega leiautes padrão
                self.load_layouts()
                preencher_tabela_layouts()
                messagebox.showinfo("Sucesso", "Leiautes restaurados para os padrões oficiais do Guia Prático!")

        tk.Button(frame_bottom_btns, text="Restaurar Padrões (Guia Prático v3.2.2)", bg="#d32f2f", fg="white", font=("Calibri", 10, "bold"), command=restaurar_padroes, padx=10, pady=5).pack(side=tk.RIGHT, padx=15)

        tree_layouts.bind("<Double-1>", lambda e: acao_editar_ou_cadastrar())

    def push_undo_state(self):
        """Salva uma cópia do estado atual do arquivo para desfazer (Ctrl+Z)."""
        if not hasattr(self, 'undo_stack'):
            self.undo_stack = []
        if len(self.undo_stack) >= 25:
            self.undo_stack.pop(0)
        self.undo_stack.append(list(self.data))

    def undo_last_action(self):
        """Reverte a última ação de edição efetuada (Ctrl+Z)."""
        if hasattr(self, 'undo_stack') and self.undo_stack:
            self.data = self.undo_stack.pop()
            self.display_data()
            self.is_modified = True
            self.update_window_title()
            messagebox.showinfo("Desfazer", f"Última alteração desfeita com sucesso. ({len(self.undo_stack)} estados restantes)")
        else:
            messagebox.showinfo("Desfazer", "Nenhuma alteração anterior para desfazer.")

    def recalcular_bloco_9(self):
        """Recalcula e regenera automaticamente a estrutura exata do Bloco 9 (9001, 9900, 9990, 9999)."""
        if not self.data:
            return

        dados_sem_bloco9 = []
        for line in self.data:
            line_str = line.strip()
            if not line_str:
                continue
            fields = line_str.split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if not reg.startswith("9"):
                dados_sem_bloco9.append(line_str)

        contagem = defaultdict(int)
        for line_str in dados_sem_bloco9:
            fields = line_str.split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if reg:
                contagem[reg] += 1

        contagem["9001"] = 1
        contagem["9990"] = 1
        contagem["9999"] = 1

        regs_ordenados = sorted(contagem.keys())
        contagem["9900"] = len(regs_ordenados) + 1

        if "9900" not in regs_ordenados:
            regs_ordenados.append("9900")
            regs_ordenados.sort()

        linhas_bloco9 = []
        linhas_bloco9.append("|9001|0|")

        for reg in regs_ordenados:
            qtd = contagem[reg]
            linhas_bloco9.append(f"|9900|{reg}|{qtd}|")

        qtd_bloco9 = len(linhas_bloco9) + 1
        linhas_bloco9.append(f"|9990|{qtd_bloco9}|")

        qtd_total_arquivo = len(dados_sem_bloco9) + len(linhas_bloco9) + 1
        linhas_bloco9.append(f"|9999|{qtd_total_arquivo}|")

        self.data = [l + "\n" for l in (dados_sem_bloco9 + linhas_bloco9)]

    @staticmethod
    def validar_dv_nfe(chave):
        """Valida o dígito verificador Módulo 11 de uma chave de NFe/CTe com 44 dígitos."""
        chave_clean = re.sub(r"\D", "", str(chave))
        if len(chave_clean) != 44:
            return False
        corpo = chave_clean[:43]
        dv_informado = int(chave_clean[43])
        pesos = [2, 3, 4, 5, 6, 7, 8, 9] * 6
        pesos = pesos[:43][::-1]
        soma = sum(int(digit) * peso for digit, peso in zip(corpo, pesos))
        resto = soma % 11
        dv_calculado = 0 if resto in (0, 1) else (11 - resto)
        return dv_informado == dv_calculado

    @staticmethod
    def validar_cnpj(cnpj):
        """Valida os dígitos verificadores de um CNPJ de 14 dígitos."""
        clean = re.sub(r"\D", "", str(cnpj))
        if len(clean) != 14 or len(set(clean)) == 1:
            return False
        def calc_dv(digits, weights):
            s = sum(int(d) * w for d, w in zip(digits, weights))
            r = s % 11
            return '0' if r < 2 else str(11 - r)
        w1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        w2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        dv1 = calc_dv(clean[:12], w1)
        dv2 = calc_dv(clean[:12] + dv1, w2)
        return clean[-2:] == (dv1 + dv2)

    @staticmethod
    def validar_cpf(cpf):
        """Valida os dígitos verificadores de um CPF de 11 dígitos."""
        clean = re.sub(r"\D", "", str(cpf))
        if len(clean) != 11 or len(set(clean)) == 1:
            return False
        def calc_dv(digits, weight_start):
            s = sum(int(d) * w for d, w in zip(digits, range(weight_start, 1, -1)))
            r = (s * 10) % 11
            return '0' if r in (10, 11) else str(r)
        dv1 = calc_dv(clean[:9], 10)
        dv2 = calc_dv(clean[:9] + dv1, 11)
        return clean[-2:] == (dv1 + dv2)

    def auditar_pre_validacao(self):
        """Audita e exibe erros de validação prévia antes de enviar para o PVA."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para auditar.")
            return

        erros = []

        for idx, line in enumerate(self.data, start=1):
            line_str = line.strip()
            if not line_str:
                continue
            fields = line_str.split("|")
            idx_off = 1 if fields[0] == "" else 0
            reg = fields[idx_off] if len(fields) > idx_off else ""

            # 1. Checagem de CNPJ / CPF no 0000 e 0150
            if reg in ("0000", "0150"):
                idx_cnpj = idx_off + 6 if reg == "0000" else idx_off + 4
                idx_cpf = idx_off + 7 if reg == "0000" else idx_off + 5
                
                cnpj_val = fields[idx_cnpj] if len(fields) > idx_cnpj else ""
                cpf_val = fields[idx_cpf] if len(fields) > idx_cpf else ""

                if cnpj_val and not self.validar_cnpj(cnpj_val):
                    erros.append((idx, reg, f"CNPJ inválido no registro {reg}: {cnpj_val}"))
                if cpf_val and not self.validar_cpf(cpf_val):
                    erros.append((idx, reg, f"CPF inválido no registro {reg}: {cpf_val}"))

            # 2. Checagem de Chave NFe no C100 / D100 / C800
            if reg in ("C100", "D100", "C800"):
                idx_chv = idx_off + 8 if reg == "C100" else idx_off + 9
                chv = fields[idx_chv] if len(fields) > idx_chv else ""
                if chv and not self.validar_dv_nfe(chv):
                    erros.append((idx, reg, f"Chave de Acesso {reg} com Dígito Verificador inválido: {chv}"))

            # 3. Checagem de Campos Obrigatórios em Branco (ex: 0200 Campo 3, 0150 Campo 3)
            regras_ob = {
                "0150": [(3, "NOME (Nome do Cliente/Fornecedor)")],
                "0200": [(3, "DESCR_ITEM (Descrição do Produto/Serviço)"), (6, "UNID_INV")]
            }
            reg_c = reg.strip().upper()
            if reg_c in regras_ob:
                for f_i, f_n in regras_ob[reg_c]:
                    tc = idx_off + (f_i - 1)
                    if len(fields) <= tc or not fields[tc].strip():
                        erros.append((idx, reg_c, f"Campo {f_i} ({f_n}) em BRANCO!"))

        # Exibir resultado em janela gráfica
        win = tk.Toplevel(self.root)
        win.title("Resultado da Pré-Validação Fiscal SPED")
        win.geometry("850x500")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=10)
        frame_top.pack(fill=tk.X)

        lbl_tit = f"Auditoria Concluída: {len(erros)} inconsistência(s) encontrada(s)"
        tk.Label(frame_top, text=lbl_tit, font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        frame_t = tk.Frame(win, padx=10, pady=10)
        frame_t.pack(fill=tk.BOTH, expand=True)

        cols = ("linha", "reg", "erro")
        tree_err = ttk.Treeview(frame_t, columns=cols, show="headings", selectmode="browse")
        tree_err.heading("linha", text="Linha")
        tree_err.heading("reg", text="Registro")
        tree_err.heading("erro", text="Descrição do Erro")

        tree_err.column("linha", width=80, anchor="center")
        tree_err.column("reg", width=90, anchor="center")
        tree_err.column("erro", width=620, anchor="w")

        vsb = ttk.Scrollbar(frame_t, orient=tk.VERTICAL, command=tree_err.yview)
        tree_err.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree_err.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, (lin, r, err) in enumerate(erros):
            tag = 'even' if i % 2 == 0 else 'odd'
            tree_err.insert("", tk.END, iid=str(lin), values=(lin, r, err), tags=(tag,))

        tree_err.tag_configure('odd', background='white')
        tree_err.tag_configure('even', background='#fff5f5')

        def ir_para_linha_erro():
            sel = tree_err.selection()
            if sel:
                l_num = sel[0]
                if hasattr(self, 'tree'):
                    if self.tree.exists(l_num):
                        self.tree.selection_set(l_num)
                        self.tree.see(l_num)
                        self.on_tree_select(None)
                        self.abrir_editor_linha(None)

        tree_err.bind("<Double-1>", lambda e: ir_para_linha_erro())
        tk.Button(win, text="Abrir Informações do Campo (Duplo-Clique)", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=ir_para_linha_erro, pady=6).pack(pady=10)

    def exportar_registro_excel(self):
        """Exporta os dados de um registro SPED selecionado para uma planilha Excel (.xlsx) com cabeçalhos oficiais."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para exportar.")
            return

        # Obter lista de registros disponíveis no arquivo
        regs_existentes = set()
        for line in self.data:
            fields = line.strip().split("|")
            reg = fields[1] if (len(fields) > 1 and fields[0] == "") else (fields[0] if len(fields) > 0 else "")
            if reg:
                regs_existentes.add(reg)

        if not regs_existentes:
            messagebox.showerror("Erro", "Nenhum registro SPED válido encontrado.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Exportar Registro para Excel")
        dlg.geometry("400x200")
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="Selecione o Registro para Exportar:", font=("Calibri", 11, "bold")).pack(pady=(20, 5))
        cbo_regs = ttk.Combobox(dlg, values=sorted(list(regs_existentes)), state="readonly", font=("Consolas", 11))
        cbo_regs.pack(pady=5)
        cbo_regs.set(sorted(list(regs_existentes))[0])

        def acao_exportar():
            target_reg = cbo_regs.get()
            dlg.destroy()

            linhas_reg = []
            for idx, line in enumerate(self.data, start=1):
                fields = line.strip().split("|")
                idx_off = 1 if fields[0] == "" else 0
                reg = fields[idx_off] if len(fields) > idx_off else ""
                if reg == target_reg:
                    # Remove campos vazios iniciais/finais do split de pipe
                    vals = fields[idx_off:]
                    if vals and vals[-1] == "":
                        vals.pop(-1)
                    linhas_reg.append([idx] + vals)

            if not linhas_reg:
                messagebox.showwarning("Aviso", f"Nenhum registro {target_reg} encontrado.")
                return

            # Obter cabeçalhos oficiais
            headers_oficiais = ["Linha"] + self.layouts.get(target_reg, [f"C{i+1}" for i in range(50)])

            # Ajustar tamanho das colunas do DataFrame
            max_cols = max(len(row) for row in linhas_reg)
            while len(headers_oficiais) < max_cols:
                headers_oficiais.append(f"C{len(headers_oficiais)}")
            headers_oficiais = headers_oficiais[:max_cols]

            df = pd.DataFrame(linhas_reg, columns=headers_oficiais)

            file_dest = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"SPED_Registro_{target_reg}.xlsx", filetypes=[("Planilha Excel (*.xlsx)", "*.xlsx")])
            if file_dest:
                try:
                    df.to_excel(file_dest, index=False)
                    ans = messagebox.askyesno("Sucesso", f"Registro {target_reg} ({len(linhas_reg)} linhas) exportado com sucesso para:\n{file_dest}\n\nDeseja abrir a planilha agora?")
                    if ans:
                        os.startfile(file_dest)
                except Exception as e:
                    messagebox.showerror("Erro ao Exportar", f"Erro ao salvar arquivo Excel: {e}")

        tk.Button(dlg, text="Exportar Planilha Excel", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=acao_exportar, pady=6, padx=10).pack(pady=20)

    def abrir_dashboard_apuracao(self):
        """Abre o Dashboard Executivo de Apuração Fiscal com indicadores visuais."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para gerar o Dashboard.")
            return

        tot_saidas_val = 0.0
        tot_saidas_icms = 0.0
        tot_entradas_val = 0.0
        tot_entradas_icms = 0.0

        tot_pis = 0.0
        tot_cofins = 0.0
        tot_ipi = 0.0

        e110_deb = 0.0
        e110_cred = 0.0
        e110_recolher = 0.0
        e110_saldo_cred = 0.0

        bloco_counts = defaultdict(int)

        def to_f(v):
            try:
                return float(str(v).replace(".", "").replace(",", "."))
            except Exception:
                return 0.0

        for line in self.data:
            line_str = line.strip()
            if not line_str:
                continue
            fields = line_str.split("|")
            idx_off = 1 if fields[0] == "" else 0
            reg = fields[idx_off] if len(fields) > idx_off else ""
            if reg:
                bloco_counts[reg[0]] += 1

            if reg == "C100":
                ind_oper = fields[idx_off + 1] if len(fields) > (idx_off + 1) else ""
                val_doc = to_f(fields[idx_off + 11]) if len(fields) > (idx_off + 11) else 0.0
                val_icms = to_f(fields[idx_off + 21]) if len(fields) > (idx_off + 21) else 0.0
                val_pis = to_f(fields[idx_off + 25]) if len(fields) > (idx_off + 25) else 0.0
                val_cofins = to_f(fields[idx_off + 26]) if len(fields) > (idx_off + 26) else 0.0
                val_ipi = to_f(fields[idx_off + 23]) if len(fields) > (idx_off + 23) else 0.0

                tot_pis += val_pis
                tot_cofins += val_cofins
                tot_ipi += val_ipi

                if ind_oper == "1":
                    tot_saidas_val += val_doc
                    tot_saidas_icms += val_icms
                elif ind_oper == "0":
                    tot_entradas_val += val_doc
                    tot_entradas_icms += val_icms

            elif reg == "E110":
                e110_deb = to_f(fields[idx_off + 1]) if len(fields) > (idx_off + 1) else 0.0
                e110_cred = to_f(fields[idx_off + 5]) if len(fields) > (idx_off + 5) else 0.0
                e110_recolher = to_f(fields[idx_off + 10]) if len(fields) > (idx_off + 10) else 0.0
                e110_saldo_cred = to_f(fields[idx_off + 13]) if len(fields) > (idx_off + 13) else 0.0

        win = tk.Toplevel(self.root)
        win.title("Dashboard de Apuração e Resumo Fiscal SPED")
        win.geometry("980x640")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=10, padx=15)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Dashboard Executivo de Apuração Fiscal", font=("Calibri", 14, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Aba 1: Resumo de Faturamento e Impostos
        aba1 = ttk.Frame(notebook)
        notebook.add(aba1, text="Resumo Geral & Apuração")

        frame_cards = tk.Frame(aba1, pady=15, padx=15)
        frame_cards.pack(fill=tk.X)

        def card(parent, title, val1_lbl, val1, val2_lbl=None, val2=None, bg="#f8fafc", fg_val="#17324d"):
            c = tk.Frame(parent, bg=bg, relief=tk.RIDGE, bd=2, padx=12, pady=10)
            tk.Label(c, text=title, font=("Calibri", 11, "bold"), bg=bg, fg="#334155").pack(anchor="w")
            tk.Label(c, text=f"{val1_lbl}: R$ {val1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 11, "bold"), bg=bg, fg=fg_val).pack(anchor="w", pady=(5, 2))
            if val2_lbl:
                tk.Label(c, text=f"{val2_lbl}: R$ {val2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 10), bg=bg, fg="#475569").pack(anchor="w")
            return c

        card(frame_cards, "Faturamento (Saídas)", "Total Vendas", tot_saidas_val, "ICMS Destacado", tot_saidas_icms, bg="#e0f2fe", fg_val="#0369a1").pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        card(frame_cards, "Compras (Entradas)", "Total Compras", tot_entradas_val, "ICMS Creditado", tot_entradas_icms, bg="#f0fdf4", fg_val="#15803d").pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        card(frame_cards, "Outros Impostos (Notas)", "PIS Total", tot_pis, "COFINS Total", tot_cofins, bg="#fef3c7", fg_val="#b45309").pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # Apuração E110
        frame_e110 = tk.LabelFrame(aba1, text="Apuração de ICMS (Bloco E110)", font=("Calibri", 11, "bold"), padx=15, pady=12)
        frame_e110.pack(fill=tk.X, padx=15, pady=10)

        f_e1 = tk.Frame(frame_e110)
        f_e1.pack(fill=tk.X)

        tk.Label(f_e1, text=f"Total Débitos: R$ {e110_deb:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 11), fg="#dc2626").pack(side=tk.LEFT, padx=15)
        tk.Label(f_e1, text=f"Total Créditos: R$ {e110_cred:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 11), fg="#16a34a").pack(side=tk.LEFT, padx=15)
        tk.Label(f_e1, text=f"ICMS a Recolher: R$ {e110_recolher:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 11, "bold"), fg="#b91c1c").pack(side=tk.LEFT, padx=15)
        tk.Label(f_e1, text=f"Saldo Credor Seg.: R$ {e110_saldo_cred:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), font=("Consolas", 11, "bold"), fg="#0d9488").pack(side=tk.LEFT, padx=15)

        # Aba 2: Distribuição por Bloco
        aba2 = ttk.Frame(notebook)
        notebook.add(aba2, text="Contagem por Bloco SPED")

        frame_bl = tk.Frame(aba2, padx=15, pady=15)
        frame_bl.pack(fill=tk.BOTH, expand=True)

        cols_b = ("bloco", "qtd")
        tree_b = ttk.Treeview(frame_bl, columns=cols_b, show="headings")
        tree_b.heading("bloco", text="Bloco SPED")
        tree_b.heading("qtd", text="Quantidade de Linhas")
        tree_b.column("bloco", width=200, anchor="center")
        tree_b.column("qtd", width=300, anchor="center")

        n_blocos = {"0": "Bloco 0 - Abertura e Cadastros", "A": "Bloco A - Serviços (PIS/COFINS)", "B": "Bloco B - ISSQN", "C": "Bloco C - Mercadorias (NFe/CTe)", "D": "Bloco D - Transportes", "E": "Bloco E - Apuração ICMS/IPI", "F": "Bloco F - Demais Documentos", "G": "Bloco G - CIAP", "H": "Bloco H - Inventário", "K": "Bloco K - Produção/Estoque", "M": "Bloco M - Apuração PIS/COFINS", "P": "Bloco P - CPRB", "1": "Bloco 1 - Outras Informações", "9": "Bloco 9 - Encerramento"}

        for i, (b, q) in enumerate(sorted(bloco_counts.items())):
            desc_b = n_blocos.get(b, f"Bloco {b}")
            tag = 'even' if i % 2 == 0 else 'odd'
            tree_b.insert("", tk.END, values=(desc_b, q), tags=(tag,))

        tree_b.tag_configure('odd', background='white')
        tree_b.tag_configure('even', background='#f8fafc')
        tree_b.pack(fill=tk.BOTH, expand=True)

    def abrir_comparador_speds(self):
        """Compara o arquivo SPED aberto com um segundo arquivo SPED no disco."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto como referência principal.")
            return

        file2_path = filedialog.askopenfilename(title="Selecione o Segundo Arquivo SPED para Comparar", filetypes=[("Arquivos SPED (*.txt;*.ret)", "*.txt;*.ret"), ("Todos os Arquivos", "*.*")])
        if not file2_path:
            return

        try:
            with open(file2_path, "r", encoding="utf-8", errors="ignore") as f:
                data2 = f.readlines()
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Não foi possível ler o segundo arquivo:\n{e}")
            return

        # Mapeia C100 do Arquivo 1
        nfs1 = {}
        for line in self.data:
            fields = line.strip().split("|")
            idx_off = 1 if fields[0] == "" else 0
            if len(fields) > (idx_off + 11) and fields[idx_off] == "C100":
                num_doc = fields[idx_off + 8]
                chave = fields[idx_off + 9]
                val = fields[idx_off + 11]
                key = chave if len(chave) == 44 else num_doc
                nfs1[key] = (num_doc, val, line.strip())

        # Mapeia C100 do Arquivo 2
        nfs2 = {}
        for line in data2:
            fields = line.strip().split("|")
            idx_off = 1 if fields[0] == "" else 0
            if len(fields) > (idx_off + 11) and fields[idx_off] == "C100":
                num_doc = fields[idx_off + 8]
                chave = fields[idx_off + 9]
                val = fields[idx_off + 11]
                key = chave if len(chave) == 44 else num_doc
                nfs2[key] = (num_doc, val, line.strip())

        so_no_arq1 = set(nfs1.keys()) - set(nfs2.keys())
        so_no_arq2 = set(nfs2.keys()) - set(nfs1.keys())
        divergencias_val = []

        for k in set(nfs1.keys()).intersection(set(nfs2.keys())):
            v1 = nfs1[k][1]
            v2 = nfs2[k][1]
            if v1 != v2:
                divergencias_val.append((nfs1[k][0], v1, v2))

        win = tk.Toplevel(self.root)
        win.title("Resultado da Comparação de Arquivos SPED")
        win.geometry("900x550")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=12)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Comparação entre Arquivos SPED", font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Tab 1: Faltantes no Arq 2
        aba1 = ttk.Frame(notebook)
        notebook.add(aba1, text=f"Ausentes no Arq 2 ({len(so_no_arq1)})")

        cols = ("num", "val", "key")
        tree1 = ttk.Treeview(aba1, columns=cols, show="headings")
        tree1.heading("num", text="Nº Doc")
        tree1.heading("val", text="Valor Total")
        tree1.heading("key", text="Chave / Identificador")
        tree1.column("num", width=120, anchor="center")
        tree1.column("val", width=120, anchor="center")
        tree1.column("key", width=550, anchor="w")
        tree1.pack(fill=tk.BOTH, expand=True)

        for k in so_no_arq1:
            num, val, _ = nfs1[k]
            tree1.insert("", tk.END, values=(num, val, k))

        # Tab 2: Faltantes no Arq 1
        aba2 = ttk.Frame(notebook)
        notebook.add(aba2, text=f"Ausentes no Arq 1 ({len(so_no_arq2)})")

        tree2 = ttk.Treeview(aba2, columns=cols, show="headings")
        tree2.heading("num", text="Nº Doc")
        tree2.heading("val", text="Valor Total")
        tree2.heading("key", text="Chave / Identificador")
        tree2.column("num", width=120, anchor="center")
        tree2.column("val", width=120, anchor="center")
        tree2.column("key", width=550, anchor="w")
        tree2.pack(fill=tk.BOTH, expand=True)

        for k in so_no_arq2:
            num, val, _ = nfs2[k]
            tree2.insert("", tk.END, values=(num, val, k))

        # Tab 3: Divergência de Valores
        aba3 = ttk.Frame(notebook)
        notebook.add(aba3, text=f"Divergência de Valores ({len(divergencias_val)})")

        cols_d = ("num", "v1", "v2")
        tree3 = ttk.Treeview(aba3, columns=cols_d, show="headings")
        tree3.heading("num", text="Nº Doc")
        tree3.heading("v1", text="Valor Arq 1")
        tree3.heading("v2", text="Valor Arq 2")
        tree3.column("num", width=150, anchor="center")
        tree3.column("v1", width=250, anchor="center")
        tree3.column("v2", width=250, anchor="center")
        tree3.pack(fill=tk.BOTH, expand=True)

        for num, v1, v2 in divergencias_val:
            tree3.insert("", tk.END, values=(num, v1, v2))

    def auditar_matriz_tributaria(self):
        """Audita a coerência de CFOP x CST nos itens C170 e C190 com opção de correção em 1-clique."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para auditar.")
            return

        inconsistencias = []

        for idx, line in enumerate(self.data, start=1):
            fields = line.strip().split("|")
            idx_off = 1 if fields[0] == "" else 0
            reg = fields[idx_off] if len(fields) > idx_off else ""

            if reg == "C170":
                cfop = fields[idx_off + 10] if len(fields) > (idx_off + 10) else ""
                cst_icms = fields[idx_off + 9] if len(fields) > (idx_off + 9) else ""
                
                # Regra 1: CST Tributado (000, 010) com CFOP de ST/Isento (5405, 5403, 1405)
                if cst_icms in ("000", "010", "020") and cfop in ("5405", "5403", "1405", "2405", "6405"):
                    inconsistencias.append((idx, reg, f"Item com CST Tributado ({cst_icms}) mas CFOP de Substituição Tributária ({cfop}). Sugestão: Alterar CST para 060."))
                
                # Regra 2: CFOP de Entrada em documento de Emissão Própria de Saída
                if cfop.startswith(("1", "2")) and len(fields) > (idx_off + 10):
                    pass

            elif reg == "C190":
                cst_icms = fields[idx_off + 1] if len(fields) > (idx_off + 1) else ""
                cfop = fields[idx_off + 2] if len(fields) > (idx_off + 2) else ""
                if cst_icms in ("000", "010") and cfop in ("5405", "5403", "1405", "2405", "6405"):
                    inconsistencias.append((idx, reg, f"Registro C190 com CST Tributado ({cst_icms}) e CFOP de ST ({cfop}). Sugestão: Alterar CST para 060."))

        win = tk.Toplevel(self.root)
        win.title("Auditoria de Matriz Tributária (CFOP x CST)")
        win.geometry("900x500")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=12)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text=f"Auditoria Concluída: {len(inconsistencias)} inconsistência(s) de Matriz Tributária", font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        frame_t = tk.Frame(win, padx=10, pady=10)
        frame_t.pack(fill=tk.BOTH, expand=True)

        cols = ("lin", "reg", "det")
        tree = ttk.Treeview(frame_t, columns=cols, show="headings")
        tree.heading("lin", text="Linha")
        tree.heading("reg", text="Registro")
        tree.heading("det", text="Inconsistência Encontrada & Sugestão de Correção")

        tree.column("lin", width=80, anchor="center")
        tree.column("reg", width=90, anchor="center")
        tree.column("det", width=670, anchor="w")

        vsb = ttk.Scrollbar(frame_t, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, (lin, r, det) in enumerate(inconsistencias):
            tag = 'even' if i % 2 == 0 else 'odd'
            tree.insert("", tk.END, iid=str(lin), values=(lin, r, det), tags=(tag,))

        tree.tag_configure('odd', background='white')
        tree.tag_configure('even', background='#fff5f5')

        def ir_para_linha_matriz():
            sel = tree.selection()
            if sel:
                l_num = sel[0]
                if hasattr(self, 'tree'):
                    if self.tree.exists(l_num):
                        self.tree.selection_set(l_num)
                        self.tree.see(l_num)
                        self.on_tree_select(None)
                        self.abrir_editor_linha(None)

        tree.bind("<Double-1>", lambda e: ir_para_linha_matriz())

        def acao_corrigir_lote():
            self.push_undo_state()
            corrigidos = 0
            for lin, r, det in inconsistencias:
                idx_arr = lin - 1
                if 0 <= idx_arr < len(self.data):
                    fields = self.data[idx_arr].strip().split("|")
                    idx_off = 1 if fields[0] == "" else 0
                    if r == "C170" and len(fields) > (idx_off + 9):
                        fields[idx_off + 9] = "060"
                        self.data[idx_arr] = "|".join(fields) + "\n"
                        corrigidos += 1
                    elif r == "C190" and len(fields) > (idx_off + 1):
                        fields[idx_off + 1] = "060"
                        self.data[idx_arr] = "|".join(fields) + "\n"
                        corrigidos += 1
            self.display_data()
            self.is_modified = True
            self.update_window_title()
            win.destroy()
            messagebox.showinfo("Sucesso", f"Total de {corrigidos} registro(s) corrigido(s) automaticamente para CST 060!")

        tk.Button(win, text="Corrigir Inconsistências em Lote (1-Clique)", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=acao_corrigir_lote, pady=6, padx=10).pack(pady=10)

    def get_custom_macros_path(self):
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "sped_rules_custom.json")

    def abrir_gerenciador_macros(self):
        """Abre a janela de Gerenciamento de Regras e Macros Automáticas de Edição."""
        path_json = self.get_custom_macros_path()
        macros = []
        if os.path.exists(path_json):
            try:
                with open(path_json, "r", encoding="utf-8") as f:
                    macros = json.load(f)
            except Exception:
                macros = []

        win = tk.Toplevel(self.root)
        win.title("Gerenciador de Regras e Macros Personalizadas SPED")
        win.geometry("850x520")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=12)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="Gerenciador de Macros de Edição Recorrente", font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        frame_t = tk.Frame(win, padx=10, pady=10)
        frame_t.pack(fill=tk.BOTH, expand=True)

        cols = ("nome", "reg", "cond", "acao")
        tree_m = ttk.Treeview(frame_t, columns=cols, show="headings", selectmode="browse")
        tree_m.heading("nome", text="Nome da Macro")
        tree_m.heading("reg", text="Registro")
        tree_m.heading("cond", text="Condição (Campo = Valor)")
        tree_m.heading("acao", text="Ação (Definir Campo = Novo Valor)")

        tree_m.column("nome", width=220, anchor="w")
        tree_m.column("reg", width=80, anchor="center")
        tree_m.column("cond", width=240, anchor="w")
        tree_m.column("acao", width=260, anchor="w")

        tree_m.pack(fill=tk.BOTH, expand=True)

        def preencher_macros():
            for item in tree_m.get_children():
                tree_m.delete(item)
            for i, m in enumerate(macros):
                cond_str = f"Campo {m.get('cond_col')} == '{m.get('cond_val')}'"
                act_str = f"Campo {m.get('act_col')} = '{m.get('act_val')}'"
                tree_m.insert("", tk.END, iid=str(i), values=(m.get("nome"), m.get("reg"), cond_str, act_str))

        preencher_macros()

        def criar_nova_macro():
            dlg = tk.Toplevel(win)
            dlg.title("Criar Nova Macro de Edição Recorrente")
            dlg.geometry("450x380")
            dlg.transient(win)
            dlg.grab_set()

            tk.Label(dlg, text="Nome da Macro:", font=("Calibri", 10, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
            ent_nome = tk.Entry(dlg, width=45)
            ent_nome.pack(anchor="w", padx=15)

            tk.Label(dlg, text="Registro Alvo (ex: C170, 0150, C100):", font=("Calibri", 10, "bold")).pack(anchor="w", padx=15, pady=(10, 2))
            ent_reg = tk.Entry(dlg, width=15)
            ent_reg.pack(anchor="w", padx=15)

            f_c = tk.Frame(dlg)
            f_c.pack(fill=tk.X, padx=15, pady=5)
            tk.Label(f_c, text="Condição (Nº Campo = Valor):", font=("Calibri", 10, "bold")).pack(anchor="w")
            ent_ccol = tk.Entry(f_c, width=6)
            ent_ccol.pack(side=tk.LEFT, padx=(0, 5))
            tk.Label(f_c, text="=").pack(side=tk.LEFT)
            ent_cval = tk.Entry(f_c, width=20)
            ent_cval.pack(side=tk.LEFT, padx=5)

            f_a = tk.Frame(dlg)
            f_a.pack(fill=tk.X, padx=15, pady=5)
            tk.Label(f_a, text="Ação (Nº Campo = Novo Valor):", font=("Calibri", 10, "bold")).pack(anchor="w")
            ent_acol = tk.Entry(f_a, width=6)
            ent_acol.pack(side=tk.LEFT, padx=(0, 5))
            tk.Label(f_a, text="=").pack(side=tk.LEFT)
            ent_aval = tk.Entry(f_a, width=20)
            ent_aval.pack(side=tk.LEFT, padx=5)

            def salvar_macro():
                nome = ent_nome.get().strip()
                reg = ent_reg.get().strip().upper()
                try:
                    ccol = int(ent_ccol.get().strip())
                    cval = ent_cval.get().strip()
                    acol = int(ent_acol.get().strip())
                    aval = ent_aval.get().strip()
                except ValueError:
                    messagebox.showerror("Erro", "O número das colunas/campos deve ser um inteiro válido.")
                    return

                if not nome or not reg:
                    messagebox.showerror("Erro", "Informe o Nome e o Registro Alvo.")
                    return

                nova_m = {"nome": nome, "reg": reg, "cond_col": ccol, "cond_val": cval, "act_col": acol, "act_val": aval}
                macros.append(nova_m)
                try:
                    with open(path_json, "w", encoding="utf-8") as f:
                        json.dump(macros, f, indent=2, ensure_ascii=False)
                except Exception as e:
                    print(f"Erro ao salvar macro em JSON: {e}")

                preencher_macros()
                dlg.destroy()
                messagebox.showinfo("Sucesso", "Macro criada e salva com sucesso!")

            tk.Button(dlg, text="Salvar Macro", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=salvar_macro, pady=6).pack(pady=15)

        def executar_macro_selecionada():
            sel = tree_m.selection()
            if not sel:
                messagebox.showwarning("Aviso", "Selecione uma macro na tabela para executar.")
                return

            if not self.data:
                messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto.")
                return

            idx_m = int(sel[0])
            m = macros[idx_m]

            self.push_undo_state()

            target_reg = m.get("reg")
            ccol = m.get("cond_col")
            cval = m.get("cond_val")
            acol = m.get("act_col")
            aval = m.get("act_val")

            executados = 0

            for i, line in enumerate(self.data):
                fields = line.strip().split("|")
                idx_off = 1 if fields[0] == "" else 0
                reg = fields[idx_off] if len(fields) > idx_off else ""

                if reg == target_reg:
                    idx_ccond = idx_off + (ccol - 1)
                    idx_cact = idx_off + (acol - 1)

                    if len(fields) > idx_ccond and fields[idx_ccond] == cval:
                        while len(fields) <= idx_cact:
                            fields.append("")
                        fields[idx_cact] = aval
                        self.data[i] = "|".join(fields) + "\n"
                        executados += 1

            self.display_data()
            self.is_modified = True
            self.update_window_title()
            messagebox.showinfo("Macro Executada", f"Macro '{m.get('nome')}' executada com sucesso! Total de {executados} registro(s) alterado(s).")

        frame_btns = tk.Frame(win, pady=10)
        frame_btns.pack(fill=tk.X)

        tk.Button(frame_btns, text="Executar Macro no Arquivo Aberto", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=executar_macro_selecionada, padx=10, pady=5).pack(side=tk.LEFT, padx=15)
        tk.Button(frame_btns, text="Criar Nova Macro", bg="#15803d", fg="white", font=("Calibri", 10, "bold"), command=criar_nova_macro, padx=10, pady=5).pack(side=tk.LEFT, padx=5)

    def auditar_campos_obrigatorios(self):
        """Audita o arquivo SPED em busca de campos obrigatórios em branco (ex: Descrição do 0200, Nome do 0150)."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para auditar.")
            return

        regras_obrigatorias = {
            "0000": [(6, "NOME (Nome da Empresa)"), (7, "CNPJ/CPF")],
            "0150": [(2, "COD_PART (Código do Participante)"), (3, "NOME (Nome do Cliente/Fornecedor)"), (4, "COD_PAIS (Código do País)")],
            "0200": [(2, "COD_ITEM (Código do Item)"), (3, "DESCR_ITEM (Descrição do Produto/Serviço)"), (6, "UNID_INV (Unidade de Medida)"), (7, "TIPO_ITEM (Tipo do Item)")],
            "C100": [(2, "IND_OPER"), (3, "IND_EMIT"), (4, "COD_PART"), (5, "COD_MOD"), (8, "NUM_DOC"), (11, "DT_DOC")],
            "C170": [(2, "NUM_ITEM"), (3, "COD_ITEM"), (5, "QTD"), (6, "UNID"), (7, "VL_ITEM"), (11, "CFOP"), (12, "CST_ICMS")],
            "D100": [(2, "IND_OPER"), (3, "IND_EMIT"), (4, "COD_PART"), (5, "COD_MOD"), (9, "NUM_DOC"), (11, "DT_AQS")],
            "H010": [(2, "COD_ITEM"), (3, "UNID"), (4, "QTD"), (5, "VL_UNIT"), (6, "VL_ITEM")]
        }

        inconsistencias = []

        for idx, line in enumerate(self.data, start=1):
            line_str = line.strip()
            if not line_str:
                continue
            fields = line_str.split("|")
            idx_off = 1 if fields[0] == "" else 0
            reg = fields[idx_off] if len(fields) > idx_off else ""

            # Garantir formato de registro de 4 caracteres com zfill se necessário
            reg_clean = reg.strip().upper()
            if reg_clean.isdigit() and len(reg_clean) < 4:
                reg_clean = reg_clean.zfill(4)

            if reg_clean in regras_obrigatorias:
                for f_idx, f_name in regras_obrigatorias[reg_clean]:
                    target_col = idx_off + (f_idx - 1)
                    val = fields[target_col].strip() if len(fields) > target_col else ""
                    if not val:
                        inconsistencias.append((idx, reg_clean, f"Campo {f_idx} ({f_name}) está EM BRANCO!"))

        win = tk.Toplevel(self.root)
        win.title("Auditoria de Campos Obrigatórios em Branco")
        win.geometry("880x520")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=8, padx=12)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text=f"Resultado: {len(inconsistencias)} campo(s) obrigatório(s) em branco encontrado(s)", font=("Calibri", 12, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        frame_t = tk.Frame(win, padx=10, pady=10)
        frame_t.pack(fill=tk.BOTH, expand=True)

        cols = ("lin", "reg", "det")
        tree = ttk.Treeview(frame_t, columns=cols, show="headings", selectmode="browse")
        tree.heading("lin", text="Linha")
        tree.heading("reg", text="Registro")
        tree.heading("det", text="Campo Obrigatório Ausente / Em Branco")

        tree.column("lin", width=80, anchor="center")
        tree.column("reg", width=90, anchor="center")
        tree.column("det", width=650, anchor="w")

        vsb = ttk.Scrollbar(frame_t, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for i, (lin, r, det) in enumerate(inconsistencias):
            tag = 'even' if i % 2 == 0 else 'odd'
            tree.insert("", tk.END, iid=str(lin), values=(lin, r, det), tags=(tag,))

        tree.tag_configure('odd', background='white')
        tree.tag_configure('even', background='#fff5f5')

        def ir_para_linha_matriz():
            sel = tree.selection()
            if sel:
                l_num = sel[0]
                if hasattr(self, 'tree'):
                    if self.tree.exists(l_num):
                        self.tree.selection_set(l_num)
                        self.tree.see(l_num)
                        self.on_tree_select(None)
                        self.abrir_editor_linha(None)

        tree.bind("<Double-1>", lambda e: ir_para_linha_matriz())

        def ir_para_linha_campo():
            sel = tree.selection()
            if sel:
                l_num = sel[0]
                if hasattr(self, 'tree'):
                    if self.tree.exists(l_num):
                        self.tree.selection_set(l_num)
                        self.tree.see(l_num)
                        self.on_tree_select(None)
                        self.abrir_editor_linha(None)

        tree.bind("<Double-1>", lambda e: ir_para_linha_campo())
        tk.Button(win, text="Abrir Informações do Campo (Duplo-Clique)", bg="#17324d", fg="white", font=("Calibri", 10, "bold"), command=ir_para_linha_campo, pady=6, padx=10).pack(pady=10)

    def abrir_manual_uso(self):
        """Abre a janela interativa com o Manual Completo de Uso do Editor SPED."""
        win = tk.Toplevel(self.root)
        win.title("Manual Completo de Uso - Editor SPED Fiscal & Contribuições")
        win.geometry("980x680")
        win.transient(self.root)

        frame_top = tk.Frame(win, bg="#17324d", pady=10, padx=15)
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="📖 Manual Completo de Uso - Editor SPED", font=("Calibri", 14, "bold"), fg="white", bg="#17324d").pack(side=tk.LEFT)

        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        def add_aba_manual(titulo, texto_md):
            frame_aba = ttk.Frame(notebook)
            notebook.add(frame_aba, text=titulo)

            txt = scrolledtext.ScrolledText(frame_aba, wrap=tk.WORD, font=("Calibri", 11), padx=12, pady=12)
            txt.pack(fill=tk.BOTH, expand=True)
            txt.insert(tk.END, texto_md.strip())
            txt.config(state=tk.DISABLED)

        # ----------------------------------------------------------------------
        # Capítulo 1: Visão Geral e Navegação
        # ----------------------------------------------------------------------
        c1 = """
================================================================================
CAPÍTULO 1: VISÃO GERAL E NAVEGAÇÃO NA TABELA
================================================================================

O Editor de Arquivos SPED foi desenvolvido para proporcionar máxima eficiência, segurança e precisão na edição, auditoria e tratamento de arquivos textuais do SPED Fiscal (EFD ICMS/IPI) e SPED Contribuições (EFD PIS/COFINS).

1. CARREGAMENTO ASSÍNCRONO DE ARQUIVOS
   - Ao abrir um arquivo (Ctrl+O ou menu Arquivo -> Abrir Arquivo SPED), a leitura é processada em segundo plano (Threading).
   - Uma barra de progresso na parte inferior indica a evolução do carregamento sem travar a interface.

2. NAVEGAÇÃO E MAPEAMENTO DINÂMICO DE CABEÇALHOS
   - A tabela exibe os registros delimitados por pipes (|).
   - Ao clicar em qualquer linha (ex: 0000, 0150, 0200, C100, C170, D100, E110, H010, K200, M200), o sistema identifica o código do registro (suportando zeros à esquerda) e altera INSTANTANEAMENTE os títulos das colunas (C1, C2, C3...) para os NOMES OFICIAIS DOS CAMPOS extraídos do Guia Prático da EFD v3.2.2.

3. CONTROLE DE ZOOM DA TABELA
   - Ajuste o tamanho da fonte e a altura proporcional das linhas em tempo real.
   - Atalhos rápidos:
     • Ctrl + Plus (ou Ctrl + =): Aumentar fonte
     • Ctrl + Minus (ou Ctrl + -): Diminuir fonte
     • Ctrl + 0: Redefinir para o padrão (10pt)

4. FILTROS DE PESQUISA COMBINADA
   - Filtro de Registro: Selecione no painel superior o registro desejado (ex: C100, C170, 0200) para isolar as linhas.
   - Filtro Texto: Digite qualquer texto (ex: NCM "38151210" ou CFOP "5102") para realizar busca textual combinada simultânea.
"""
        add_aba_manual("1. Visão Geral & Zoom", c1)

        # ----------------------------------------------------------------------
        # Capítulo 2: Módulos de Auditoria
        # ----------------------------------------------------------------------
        c2 = """
================================================================================
CAPÍTULO 2: MÓDULOS DE AUDITORIA E PRÉ-VALIDAÇÃO FISCAL
================================================================================

O sistema conta com 4 módulos avançados de auditoria prévia antes de submeter o arquivo ao PVA da Receita Federal:

1. AUDITORIA DE CAMPOS OBRIGATÓRIOS EM BRANCO
   - Menu: Ferramentas -> Auditoria - Campos Obrigatórios em Branco
   - Identifica descrições nulas ou códigos ausentes em registros críticos:
     • Registro 0200: Campo 3 (Descrição do Produto), Campo 2 (Código), Campo 6 (Unidade), Campo 7 (Tipo).
     • Registro 0150: Campo 3 (Nome do Cliente/Fornecedor), Campo 2 (Código Participante), Campo 4 (País).
     • Registro 0000: Campo 6 (Razão Social) e CNPJ/CPF.
     • Registros C100 / C170: Campos de identificação, CFOP ou CST ausentes.
   - Duplo-Clique: Ao dar duplo-clique em qualquer erro no relatório, o Editor de Campos da linha abre diretamente SEM FECHAR a janela da auditoria, permitindo corrigir múltiplos itens em sequência.

2. MÓDULO DE PRÉ-VALIDAÇÃO FISCAL (PVA)
   - Menu: Ferramentas -> Auditoria - Módulo de Pré-Validação Fiscal (PVA)
   - Valida com algoritmo Módulo 11 o Dígito Verificador (DV) de Chaves de Acesso de NFe/CTe (44 dígitos) nos registros C100, D100 e C800.
   - Valida os dígitos verificadores de CNPJ (14 dígitos) e CPF (11 dígitos) nos registros 0000 e 0150.

3. AUDITORIA DE MATRIZ TRIBUTÁRIA (CFOP x CST)
   - Menu: Ferramentas -> Auditoria - Matriz Tributária (CFOP x CST)
   - Identifica itens C170 e C190 tributados (CST 000, 010) associados a CFOPs de Isenção ou Substituição Tributária (5405, 5403, 1405).
   - Botão "Corrigir em Lote (1-Clique)": Altera automaticamente todos os CSTs inconsistentes para 060 com um único clique.

4. AUDITORIA DE REGISTROS DUPLICADOS
   - Menu: Ferramentas -> Auditoria - Verificar / Remover Duplicados
   - Localiza e remove ocorrências duplicadas (ex: registros 0600 ou itens repetidos) preservando a integridade das linhas únicas.
"""
        add_aba_manual("2. Auditoria & Validação", c2)

        # ----------------------------------------------------------------------
        # Capítulo 3: Tratamento de Blocos
        # ----------------------------------------------------------------------
        c3 = """
================================================================================
CAPÍTULO 3: OPERAÇÕES E TRATAMENTO POR BLOCO SPED
================================================================================

Ferramentas especializadas para ajustes e correções em lote nos Blocos do SPED:

1. BLOCO 0 (CADASTROS E ABERTURA)
   - 0200 - Preencher / Atualizar CEST: Atualiza o código CEST (campo 13) nos produtos 0200 utilizando o mapeamento oficial da planilha CEST12.xlsx (742 pares NCM x CEST).
   - 0200 - Atualizar Alíquota / Campo 7: Aplica regras automáticas de classificação no Tipo do Item.
   - 0220 - Conversão de Unidades: Processa fatores de conversão de medida.
   - 0150 - Município / IE MG: Formata Inscrições Estaduais de MG para 13 dígitos e preenche o código de município IBGE.

2. BLOCO C (DOCUMENTOS FISCAIS DE MERCADORIAS)
   - C100 - Preencher Série Padrão (000): Insere a série padrão "000" nos documentos C100 com série nula.
   - C170 - Vincular Plano de Contas (Campo 38): Associa automaticamente o código da conta contábil segundo o grupo do item.
   - C170 - Alterar CST PIS/COFINS em Lote: Atualiza os CSTs de PIS/COFINS para toda a escrituração.
   - C190 - Totalização por CFOP e Alíquota: Audita a soma dos valores do C190.
   - C190 - Gerar Ajustes C195 / C197: Cria linhas de ajuste fiscal automaticamente abaixo de cada C190.
   - C112 - Gerar Guia de Arrecadação: Insere informações de guias de arrecadação via arquivo externo.
   - Rateio de Notas: Ajusta valores de itens e notas proporcionalmente por rateio.

3. BLOCO D (TRANSPORTES E COMUNICAÇÃO)
   - D100 - Configurar Conta Contábil de Frete: Vincula contas contábeis nos conhecimentos de transporte.
   - D101 / D105 - Recalcular Impostos: Processa e gera registros D105 recalculando PIS e COFINS.

4. BLOCO H (INVENTÁRIO FÍSICO)
   - H010 - Sincronizar Unidades de Medida com 0200: Garante que as unidades cadastradas no Inventário H010 sejam idênticas às cadastradas no cadastro de produtos 0200.

5. BLOCO 9 (ENCERRAMENTO E TOTALIZAÇÃO DO ARQUIVO)
   - Recálculo Automático ao Salvar: Sempre que você salva o arquivo (Ctrl+S), o sistema limpa os blocos 9 antigos, recalcula a contagem exata de cada registro e gera as linhas 9900 (contadores por registro), 9990 (total do bloco 9) e 9999 (total geral de linhas do arquivo SPED).
"""
        add_aba_manual("3. Tratamento de Blocos", c3)

        # ----------------------------------------------------------------------
        # Capítulo 4: Recursos Avançados e Macros
        # ----------------------------------------------------------------------
        c4 = """
================================================================================
CAPÍTULO 4: INTELIGÊNCIA FISCAL, RELATÓRIOS E MACROS
================================================================================

1. DASHBOARD EXECUTIVO DE APURAÇÃO FISCAL
   - Menu: Ferramentas -> Inteligência - Dashboard Executivo de Apuração
   - Painel com cards de estatísticas visuais:
     • Total de Faturamento (Vendas / Saídas) vs Compras (Entradas).
     • Impostos Destacados (ICMS, ICMS ST, IPI, PIS e COFINS).
     • Apuração de ICMS (E110): Débitos, Créditos, ICMS a Recolher e Saldo Credor para o mês seguinte.
     • Apuração PIS/COFINS (M200/M600).
     • Tabela com a contagem de linhas de cada bloco.

2. COMPARADOR DE DOIS ARQUIVOS SPED
   - Menu: Ferramentas -> Auditoria - Comparador de Dois Arquivos SPED
   - Selecione um segundo arquivo SPED no disco (ex: SPED gerado pelo ERP vs SPED Transmitido) para cruzar dados:
     • Identifica notas fiscais (C100) ausentes no Arquivo 2 ou ausentes no Arquivo 1.
     • Destaca divergências nos valores totais entre documentos de mesmo número/chave.

3. EXPORTADOR DE REGISTROS PARA EXCEL (.XLSX)
   - Menu: Ferramentas -> Exportar Registro Selecionado para Excel (.xlsx)
   - Exporta qualquer registro específico (C100, C170, 0200, H010, E110, M200, etc.) para uma planilha Excel limpa contendo os NOMES OFICIAIS DOS CAMPOS NO CABEÇALHO DA PLANILHA.

4. GERENCIADOR DE LEIAUTES SPED
   - Menu: Cadastros -> Gerenciar Leiautes de Registros SPED
   - Permite visualizar, pesquisar e cadastrar novos registros SPED (ex: C198) ou editar nomes de campos. Salva em arquivo local (sped_layouts_custom.json).

5. GERENCIADOR DE MACROS E REGRAS SALVAS
   - Menu: Ferramentas -> Inteligência - Gerenciador de Macros Personalizadas
   - Permite salvar rotinas de edição frequentes (ex: "Definir Conta 41330 se CFOP for 5102") em arquivo JSON e executá-las com 1 clique sobre qualquer arquivo SPED.

6. HISTÓRICO DE DESFAZER (CTRL+Z)
   - Atalho Ctrl+Z reverte as últimas alterações efetuadas na sessão de trabalho.
"""
        add_aba_manual("4. Inteligência & Macros", c4)

        # ----------------------------------------------------------------------
        # Capítulo 5: Tabela de Atalhos
        # ----------------------------------------------------------------------
        c5 = """
================================================================================
CAPÍTULO 5: TABELA COMPLETA DE ATALHOS DE TECLADO
================================================================================

Aproveite os atalhos de teclado para acelerar o fluxo de trabalho:

┌─────────────────┬────────────────────────────────────────────────────────────┐
│ Atalho          │ Ação Executada                                             │
├─────────────────┼────────────────────────────────────────────────────────────┤
│ F1              │ Abre o Manual Completo de Uso                              │
│ Ctrl + O        │ Abrir Arquivo SPED...                                      │
│ Ctrl + S        │ Salvar Alterações (Recalcula o Bloco 9)                   │
│ Ctrl + Shift + S│ Salvar Como...                                             │
│ Ctrl + Z        │ Desfazer última alteração de edição                        │
│ Ctrl + F        │ Focar no campo de Busca de Texto                           │
│ F3              │ Localizar Próxima ocorrência de texto                      │
│ Ctrl + Plus / = │ Aumentar Tamanho da Fonte e Altura das Linhas da Tabela    │
│ Ctrl + Minus / -│ Diminuir Tamanho da Fonte e Altura das Linhas da Tabela    │
│ Ctrl + 0        │ Redefinir Tamanho da Fonte para o Padrão (10pt)           │
│ Duplo-Clique    │ Na Tabela: Abre Editor de Linha Completo                   │
│                 │ Nas Auditorias: Abre Editor de Campos sem fechar o relatório│
└─────────────────┴────────────────────────────────────────────────────────────┘
"""
        add_aba_manual("5. Atalhos de Teclado", c5)

    def abrir_atalhos_teclado(self):
        """Exibe uma caixa de diálogo rápida com os atalhos de teclado principais."""
        texto = """ATALHOS DE TECLADO DO EDITOR SPED:

• F1: Manual Completo de Uso
• Ctrl + O: Abrir Arquivo SPED
• Ctrl + S: Salvar Arquivo (Recalcula Bloco 9)
• Ctrl + Shift + S: Salvar Como...
• Ctrl + Z: Desfazer última alteração
• Ctrl + F: Focar no campo de busca de texto
• F3: Localizar próxima ocorrência
• Ctrl + Plus (+ / =): Aumentar Fonte e Altura da Tabela
• Ctrl + Minus (-): Diminuir Fonte e Altura da Tabela
• Ctrl + 0: Redefinir Tamanho da Fonte (10pt)
• Duplo-Clique: Na Tabela/Auditoria abre o Editor de Campos da Linha"""

        messagebox.showinfo("Atalhos de Teclado", texto)

    def search_document(self):
        search_text = self.entry_search.get().strip()
        if not search_text:
            messagebox.showwarning("Aviso", "Digite um texto para pesquisar.")
            return
        self.text_display.tag_remove("highlight", "1.0", tk.END)
        pos = self.text_display.search(search_text, self.current_search_index, stopindex=tk.END, nocase=True)
        if pos:
            end_pos = f"{pos}+{len(search_text)}c"
            self.text_display.tag_add("highlight", pos, end_pos)
            self.text_display.tag_config("highlight", background="yellow", foreground="black")
            self.text_display.see(pos)
            self.current_search_index = end_pos
            self.last_search_text = search_text
        else:
            messagebox.showinfo("Pesquisa", "Nenhuma ocorrência encontrada.")
            self.current_search_index = "1.0"

    def next_occurrence(self):
        if not self.last_search_text:
            messagebox.showwarning("Aviso", "Nenhuma pesquisa realizada ainda.")
            return
        pos = self.text_display.search(self.last_search_text, self.current_search_index, stopindex=tk.END, nocase=True)
        if pos:
            end_pos = f"{pos}+{len(self.last_search_text)}c"
            self.text_display.tag_add("highlight", pos, end_pos)
            self.text_display.tag_config("highlight", background="yellow", foreground="black")
            self.text_display.see(pos)
            self.current_search_index = end_pos
        else:
            messagebox.showinfo("Pesquisa", "Nenhuma ocorrência encontrada.")
            self.current_search_index = "1.0"

    def edit_value(self):
        try:
            line_num = int(self.entry_line.get())
            column_num = int(self.entry_column.get())
            if line_num < 1 or line_num > len(self.data):
                raise ValueError("Número de linha inválido.")
            line = self.data[line_num - 1]
            fields = line.split("|")
            if column_num < 1 or column_num > len(fields):
                raise ValueError("Número de coluna inválido.")
            current_value = fields[column_num - 1]
            response = messagebox.askyesno(
                "Editar Valor",
                f"O valor atual da linha {line_num}, coluna {column_num} é:\n\n{current_value}\n\nDeseja continuar com a edição?"
            )
            if not response:
                return
            new_value = self.entry_value.get()
            fields[column_num - 1] = new_value
            self.data[line_num - 1] = "|".join(fields)
            self.display_data()
            messagebox.showinfo("Sucesso", f"Linha {line_num}, coluna {column_num} atualizada.")
        except ValueError as e:
            messagebox.showerror("Erro", str(e))

    def process_c190(self):
        new_lines = []
        for line in self.data:
            fields = line.strip().split("|")
            new_lines.append(line)
            if len(fields) > 4 and fields[1] == "C190" and fields[4] in ("0", "0,00"):
                new_lines.append(f"|C195|{fields[3]}||\n")
                new_lines.append(f"|C197|SP90090104|{fields[3]}|0|0|0|0|{fields[5]}|\n")
        self.data = new_lines
        self.display_data()
        messagebox.showinfo("Sucesso", "Registros C195 e C197 adicionados abaixo dos registros C190.")

    def update_0200_cest(self):
        """Atualiza o código CEST (campo 13) nos registros 0200 com base na tabela NCM x CEST oficial (CEST12.xlsx)."""
        # Mapeamento NCM (8 dígitos) -> CEST (7 dígitos) importado de D:\Downloads\CEST12.xlsx
        cest_mapping = {
        "38151210": "0100100",
        "38151290": "0100100",
        "39170000": "0100200",
        "39181000": "0100300",
        "39233000": "0100400",
        "39263000": "0100500",
        "40103000": "0100600",
        "59100000": "0100600",
        "40169300": "0100700",
        "48239090": "0100700",
        "77313800": "0100800",
        "40169990": "0100900",
        "57050000": "0100900",
        "59039000": "0101000",
        "59090000": "0101100",
        "63061000": "0101200",
        "65061000": "0101300",
        "68130000": "0101400",
        "70071100": "0101500",
        "70072100": "0101500",
        "70091000": "0101600",
        "70140000": "0101700",
        "73110000": "0101800",
        "73200000": "0102000",
        "73250000": "0102100",
        "78060000": "0102200",
        "80070090": "0102300",
        "83012000": "0102400",
        "83016000": "0102400",
        "83017000": "0102500",
        "83021000": "0102600",
        "83023000": "0102600",
        "83100000": "0102700",
        "84073000": "0102800",
        "84082000": "0102900",
        "84099000": "0103000",
        "84122000": "0103100",
        "84133000": "0103200",
        "84141000": "0103300",
        "84148010": "0103400",
        "84148020": "0103400",
        "84139190": "0103500",
        "84149010": "0103500",
        "84149030": "0103500",
        "84149039": "0103500",
        "84152000": "0103600",
        "84212300": "0103700",
        "84212990": "0103800",
        "84219000": "0103900",
        "84241000": "0104000",
        "84213100": "0104100",
        "84213920": "0104200",
        "84254200": "0104300",
        "23856830": "0104400",
        "84314920": "0104500",
        "84339090": "0104500",
        "84811000": "0104600",
        "84812000": "0104700",
        "84818092": "0104800",
        "84820000": "0104900",
        "84830000": "0105000",
        "84840000": "0105100",
        "85052000": "0105200",
        "85071000": "0105300",
        "85110000": "0105400",
        "85122000": "0105500",
        "85124000": "0105500",
        "85129000": "0105500",
        "24171580": "0105600",
        "85180000": "0105700",
        "85185000": "0105800",
        "85198100": "0105900",
        "85255010": "0106000",
        "85256010": "0106000",
        "85272000": "0106100",
        "85272190": "0106200",
        "85219090": "0106200",
        "85291090": "0106300",
        "85340000": "0106400",
        "85353000": "0106500",
        "85365000": "0106500",
        "85361000": "0106600",
        "85362000": "0106700",
        "85364000": "0106800",
        "85380000": "0106900",
        "85365090": "0107000",
        "85391000": "0107100",
        "85392000": "0107200",
        "85442000": "0107300",
        "85443000": "0107400",
        "87070000": "0107500",
        "87080000": "0107600",
        "87141000": "0107700",
        "87169090": "0107800",
        "90261000": "0107900",
        "90262000": "0108000",
        "90290000": "0108100",
        "90303321": "0108200",
        "90318040": "0108300",
        "90328920": "0108400",
        "91040000": "0108500",
        "94012000": "0108600",
        "94019090": "0108600",
        "96138000": "0108700",
        "40090000": "0108800",
        "45049000": "0108900",
        "68129910": "0108900",
        "48234000": "0109000",
        "39191000": "0109100",
        "39199000": "0109100",
        "87082999": "0109100",
        "84123110": "0109200",
        "84131900": "0109300",
        "84135090": "0109300",
        "84138100": "0109300",
        "84136019": "0109400",
        "84137010": "0109400",
        "84145910": "0109500",
        "84145990": "0109500",
        "84213990": "0109600",
        "24112590": "0109700",
        "85013110": "0109800",
        "85045000": "0109900",
        "85072000": "0110000",
        "85073000": "0110000",
        "85123000": "0110100",
        "90328980": "0110200",
        "90328990": "0110200",
        "90271000": "0110300",
        "40081100": "0110400",
        "56012219": "0110500",
        "57032000": "0110600",
        "57033000": "0110700",
        "59119000": "0110800",
        "69039099": "0110900",
        "70072900": "0111000",
        "73145000": "0111100",
        "73151100": "0111200",
        "19781330": "0111300",
        "84189900": "0111400",
        "84195000": "0111500",
        "84249090": "0111600",
        "84254910": "0111700",
        "84314100": "0111800",
        "85016100": "0111900",
        "85311090": "0112000",
        "90141000": "0112100",
        "90251990": "0112200",
        "90259010": "0112300",
        "90269000": "0112400",
        "26051940": "0112500",
        "90321090": "0112600",
        "90322000": "0112700",
        "87169000": "0112800",
        "22050000": "0200100",
        "22089000": "0200100",
        "22072000": "0200400",
        "22084000": "0200400",
        "22060090": "0200500",
        "22082000": "0200600",
        "22085000": "0200800",
        "22087000": "0201000",
        "22083000": "0201600",
        "22086000": "0201800",
        "22060010": "0202200",
        "22040000": "0202400",
        "22060000": "0202500",
        "22070000": "0202500",
        "22080000": "0202500",
        "22010000": "0300100",
        "22020000": "0300600",
        "21069010": "0300900",
        "22029000": "0301000",
        "21069000": "0301100",
        "22030000": "0301200",
        "24020000": "0400100",
        "24031000": "0400200",
        "25230000": "0500100",
        "22071000": "0600100",
        "29618600": "0600200",
        "27101910": "0600300",
        "27101920": "0600400",
        "27101930": "0600500",
        "27101990": "0600600",
        "27109000": "0600700",
        "27110000": "0600800",
        "27130000": "0600900",
        "38260000": "0601000",
        "34030000": "0601100",
        "27102000": "0601200",
        "27160000": "0700100",
        "44170010": "0800200",
        "44170090": "0800200",
        "68040000": "0800300",
        "82010000": "0800400",
        "82022000": "0800500",
        "82029100": "0800600",
        "82020000": "0800700",
        "82030000": "0800800",
        "82040000": "0800900",
        "82050000": "0801000",
        "82060000": "0801100",
        "82074000": "0801200",
        "82076000": "0801200",
        "82077000": "0801200",
        "82070000": "0801300",
        "82080000": "0801400",
        "82090011": "0801500",
        "82090000": "0801600",
        "82110000": "0801700",
        "82130000": "0801800",
        "84670000": "0801900",
        "90150000": "0802000",
        "90172000": "0802100",
        "90173000": "0802100",
        "90178000": "0802100",
        "90179090": "0802100",
        "90251190": "0802200",
        "90251900": "0802300",
        "90259090": "0802300",
        "85390000": "0900100",
        "85400000": "0900200",
        "85041000": "0900300",
        "85437099": "0900500",
        "25220000": "1000100",
        "38160010": "1000200",
        "38245000": "1000200",
        "32149000": "1000300",
        "39100000": "1000400",
        "39160000": "1000500",
        "39180000": "1000700",
        "39190000": "1000800",
        "39200000": "1000900",
        "39210000": "1000900",
        "39220000": "1001300",
        "39240000": "1001400",
        "39251000": "1001500",
        "39259000": "1001600",
        "39252000": "1001800",
        "39253000": "1001900",
        "39269000": "1002000",
        "48140000": "1002100",
        "68101900": "1002200",
        "68110000": "1002300",
        "69010000": "1002500",
        "69020000": "1002600",
        "69040000": "1002700",
        "69050000": "1002800",
        "69060000": "1002900",
        "69070000": "1003000",
        "69080000": "1003000",
        "69090000": "1003001",
        "69100000": "1003100",
        "69120000": "1003200",
        "70030000": "1003300",
        "70040000": "1003400",
        "70050000": "1003500",
        "70071900": "1003600",
        "70080000": "1003800",
        "70160000": "1003900",
        "72142000": "1004000",
        "73089010": "1004100",
        "72130000": "1004300",
        "72171090": "1004400",
        "73120000": "1004400",
        "72172000": "1004500",
        "73070000": "1004600",
        "73083000": "1004700",
        "73084000": "1004800",
        "73089000": "1004800",
        "73089090": "1005000",
        "73100000": "1005100",
        "73130000": "1005200",
        "73140000": "1005300",
        "73151290": "1005500",
        "73158200": "1005600",
        "73170000": "1005700",
        "73180000": "1005800",
        "73230000": "1005900",
        "73240000": "1006000",
        "73260000": "1006200",
        "74070000": "1006300",
        "20131350": "1006400",
        "74120000": "1006500",
        "74150000": "1006600",
        "74182000": "1006700",
        "76071990": "1006800",
        "76080000": "1006900",
        "76090000": "1007000",
        "76100000": "1007100",
        "76152000": "1007200",
        "76160000": "1007300",
        "83024100": "1007400",
        "83010000": "1007500",
        "83070000": "1007700",
        "83110000": "1007800",
        "84810000": "1007900",
        "28289011": "1100100",
        "28289019": "1100100",
        "32064100": "1100100",
        "38089419": "1100100",
        "34012090": "1100200",
        "34022000": "1100400",
        "34020000": "1100700",
        "38099190": "1100800",
        "39241000": "1100900",
        "39249000": "1100900",
        "68053010": "1100900",
        "68053090": "1100900",
        "73231000": "1101100",
        "85040000": "1200100",
        "85160000": "1200200",
        "85350000": "1200300",
        "85360000": "1200400",
        "74130000": "1200600",
        "85440000": "1200700",
        "76050000": "1200700",
        "76140000": "1200700",
        "85460000": "1200800",
        "85470000": "1200900",
        "30030000": "1300100",
        "30040000": "1300100",
        "30050000": "1300101",
        "30060000": "1300102",
        "30070000": "1300200",
        "30080000": "1300201",
        "30090000": "1300202",
        "30100000": "1300300",
        "30110000": "1300301",
        "30120000": "1300302",
        "30130000": "1300400",
        "30140000": "1300401",
        "30150000": "1300402",
        "30066000": "1300500",
        "29360000": "1300600",
        "30063000": "1300700",
        "30020000": "1300800",
        "40151100": "1301000",
        "40151900": "1301000",
        "40141000": "1301100",
        "90183100": "1301200",
        "90183210": "1301300",
        "39269090": "1301400",
        "90189099": "1301400",
        "48232090": "1400100",
        "48236000": "1400200",
        "48131000": "1400300",
        "40110000": "1600100",
        "40115000": "1600500",
        "40121000": "1600600",
        "40129000": "1600700",
        "40130000": "1600800",
        "40132000": "1600900",
        "17049010": "1700100",
        "18063110": "1700200",
        "18063120": "1700200",
        "18063210": "1700300",
        "18063220": "1700300",
        "18069000": "1700400",
        "17049090": "1700700",
        "21012000": "1700900",
        "22021000": "1701000",
        "20090000": "1701200",
        "20098000": "1701300",
        "04021000": "1701700",
        "04022000": "1701700",
        "04029000": "1701700",
        "06590000": "1701800",
        "06490000": "1701900",
        "19011090": "1702000",
        "19011030": "1702000",
        "04011010": "1702100",
        "04012010": "1702100",
        "04014010": "1702200",
        "04015010": "1702200",
        "04011090": "1702300",
        "04012090": "1702300",
        "04014020": "1702400",
        "04022130": "1702400",
        "04022930": "1702400",
        "04011000": "1702402",
        "04012000": "1702402",
        "04015000": "1702402",
        "04022920": "1702402",
        "40290000": "1702500",
        "04030000": "1702600",
        "40390000": "1702700",
        "04060000": "1702800",
        "40510000": "1703000",
        "15171000": "1703100",
        "15179000": "1703202",
        "15162000": "1703300",
        "19019020": "1703400",
        "19041000": "1703500",
        "19049000": "1703500",
        "19059090": "1703600",
        "20052000": "1703700",
        "20059000": "1703700",
        "20081000": "1703800",
        "21032010": "1703900",
        "21039021": "1704000",
        "21039091": "1704000",
        "74428000": "1704100",
        "21033010": "1704200",
        "21033021": "1704300",
        "21039011": "1704400",
        "20020000": "1704500",
        "19042000": "1704700",
        "11010010": "1704900",
        "11010020": "1705000",
        "19012000": "1705100",
        "19019090": "1705100",
        "19023000": "1705200",
        "19020000": "1705300",
        "19024000": "1705301",
        "19021000": "1705400",
        "19052000": "1705500",
        "19052090": "1705600",
        "19052010": "1705700",
        "19053100": "1705800",
        "19059020": "1706100",
        "19053200": "1706200",
        "19054000": "1706400",
        "19059010": "1706500",
        "19051000": "1706800",
        "19059000": "1707000",
        "15079011": "1707100",
        "15080000": "1707200",
        "15090000": "1707300",
        "15100000": "1707400",
        "15121911": "1707500",
        "15122910": "1707500",
        "15141000": "1707600",
        "15151900": "1707700",
        "15152910": "1707800",
        "15122990": "1707900",
        "15179010": "1708000",
        "15110000": "1708100",
        "15130000": "1708100",
        "15140000": "1708100",
        "15150000": "1708100",
        "15160000": "1708100",
        "15180000": "1708100",
        "16010000": "1708200",
        "16020000": "1708500",
        "16040000": "1708600",
        "16050000": "1708800",
        "02010000": "1708900",
        "02020000": "1708900",
        "02040000": "1708900",
        "02060000": "1708900",
        "02102000": "1708900",
        "02109900": "1708900",
        "15020000": "1708900",
        "02030000": "1709100",
        "02070000": "1709100",
        "02090000": "1709100",
        "02101000": "1709100",
        "15010000": "1709100",
        "07100000": "1709200",
        "08110000": "1709300",
        "20010000": "1709400",
        "20040000": "1709500",
        "20050000": "1709600",
        "20060000": "1709700",
        "20070000": "1709800",
        "20080000": "1709900",
        "09010000": "1710000",
        "09020000": "1710100",
        "12119090": "1710100",
        "21069090": "1710100",
        "90300000": "1710200",
        "17011000": "1710300",
        "17019900": "1710300",
        "17019100": "1710400",
        "17020000": "1710900",
        "20081900": "1711000",
        "21011000": "1711100",
        "21011190": "1711300",
        "21011200": "1711300",
        "18305140": "1800100",
        "69111090": "1800200",
        "32131000": "1900100",
        "39162000": "1900200",
        "39261000": "1900300",
        "42021000": "1900400",
        "42029000": "1900400",
        "48022090": "1900600",
        "48119090": "1900600",
        "48025490": "1900700",
        "48025499": "1900800",
        "48025799": "1900800",
        "48162000": "1900800",
        "48025690": "1900900",
        "48025790": "1900900",
        "48025890": "1900900",
        "37031010": "1901000",
        "37031029": "1901000",
        "37032000": "1901000",
        "37039010": "1901000",
        "37040000": "1901000",
        "48022000": "1901000",
        "48101390": "1901100",
        "48161000": "1901200",
        "39202019": "1901300",
        "48062000": "1901400",
        "48081000": "1901500",
        "48102290": "1901600",
        "48090000": "1901700",
        "48160000": "1901700",
        "48170000": "1901800",
        "48201000": "1901900",
        "48202000": "1902000",
        "48203000": "1902100",
        "48204000": "1902200",
        "48205000": "1902300",
        "48209000": "1902400",
        "49090000": "1902500",
        "96081000": "1902600",
        "96082000": "1902700",
        "96083000": "1902800",
        "96080000": "1902900",
        "48025600": "1903000",
        "27121000": "2000200",
        "28142000": "2000300",
        "28470000": "2000400",
        "30067000": "2000500",
        "33010000": "2000600",
        "33030010": "2000700",
        "33030020": "2000800",
        "33041000": "2000900",
        "33042010": "2001000",
        "33042090": "2001100",
        "33043000": "2001200",
        "33049100": "2001300",
        "33049910": "2001400",
        "33049990": "2001500",
        "33051000": "2001700",
        "33052000": "2001800",
        "33053000": "2001900",
        "33059000": "2002000",
        "33061000": "2002300",
        "33062000": "2002400",
        "33069000": "2002500",
        "33071000": "2002600",
        "33072010": "2002700",
        "33072090": "2002900",
        "33073000": "2003100",
        "33079000": "2003200",
        "34011190": "2003400",
        "34011900": "2003500",
        "34012010": "2003600",
        "34013000": "2003700",
        "40149010": "2003800",
        "40149090": "2003900",
        "39269040": "2003901",
        "48181000": "2004100",
        "48182000": "2004300",
        "48183000": "2004500",
        "48189090": "2004600",
        "96190000": "2004700",
        "56012190": "2005000",
        "56039290": "2005100",
        "82032090": "2005200",
        "82141000": "2005300",
        "82142000": "2005400",
        "90251110": "2005500",
        "96032000": "2005600",
        "96032100": "2005700",
        "96033000": "2005800",
        "96050000": "2005900",
        "96150000": "2006000",
        "96162000": "2006100",
        "70102000": "2006200",
        "82121020": "2006300",
        "82122010": "2006300",
        "73211100": "2100100",
        "73218100": "2100100",
        "73219000": "2100100",
        "84181000": "2100200",
        "84182100": "2100300",
        "84182900": "2100400",
        "84183000": "2100500",
        "84184000": "2100600",
        "84185000": "2100700",
        "84186990": "2100800",
        "84186999": "2100900",
        "84211200": "2101100",
        "84211990": "2101200",
        "84186931": "2101300",
        "84221100": "2101500",
        "84229010": "2101500",
        "84433100": "2101600",
        "84433200": "2101700",
        "84439900": "2101800",
        "84501100": "2101900",
        "84501200": "2102000",
        "84501900": "2102100",
        "84502000": "2102200",
        "84509000": "2102300",
        "84512100": "2102400",
        "84512990": "2102500",
        "84519000": "2102600",
        "84521000": "2102700",
        "84713000": "2102800",
        "84714000": "2102900",
        "84715010": "2103000",
        "84716050": "2103100",
        "84716090": "2103200",
        "84717000": "2103300",
        "84719000": "2103400",
        "84733000": "2103500",
        "85043000": "2103600",
        "85044010": "2103700",
        "85044040": "2103800",
        "85078000": "2103900",
        "85080000": "2104000",
        "85090000": "2104100",
        "85098010": "2104200",
        "85161000": "2104300",
        "85164000": "2104400",
        "85165000": "2104500",
        "85166000": "2104600",
        "85167100": "2104800",
        "85167200": "2104900",
        "85167900": "2105000",
        "85169000": "2105100",
        "85171100": "2105200",
        "85171200": "2105300",
        "85171890": "2105400",
        "85176250": "2105500",
        "85190000": "2105700",
        "85220000": "2105700",
        "85271000": "2105700",
        "85198190": "2105800",
        "85219010": "2105900",
        "85235110": "2106100",
        "85235200": "2106200",
        "85258020": "2106300",
        "85279000": "2106400",
        "85284929": "2106500",
        "85285920": "2106500",
        "85286900": "2106500",
        "85285120": "2106600",
        "85287000": "2106700",
        "90061000": "2107200",
        "90064000": "2107300",
        "90189050": "2107400",
        "90191000": "2107500",
        "90328911": "2107600",
        "95045000": "2107700",
        "85176210": "2107800",
        "85176222": "2107900",
        "85176239": "2108000",
        "85176240": "2108100",
        "85176262": "2108200",
        "85176290": "2108300",
        "85177021": "2108400",
        "82149000": "2108500",
        "85100000": "2108500",
        "84145000": "2108600",
        "84146000": "2108800",
        "84149020": "2108900",
        "84151000": "2109000",
        "84158000": "2109000",
        "23798400": "2109100",
        "23798480": "2109200",
        "84151090": "2109300",
        "84159010": "2109400",
        "84159020": "2109500",
        "84212100": "2109600",
        "84243010": "2109700",
        "84243090": "2109700",
        "84672100": "2109800",
        "85162000": "2109900",
        "85163100": "2110000",
        "85163200": "2110100",
        "85270000": "2110300",
        "84796000": "2110500",
        "84159090": "2110600",
        "85258019": "2110700",
        "84231000": "2110800",
        "85170000": "2111000",
        "85290000": "2111200",
        "85310000": "2111300",
        "85311000": "2111400",
        "85318000": "2111500",
        "85414011": "2111700",
        "85414021": "2111700",
        "85414022": "2111700",
        "85437092": "2111800",
        "90303000": "2111900",
        "90308900": "2112000",
        "91070000": "2112100",
        "94050000": "2112200",
        "23090000": "2200100",
        "21050000": "2300100",
        "18060000": "2300200",
        "19010000": "2300200",
        "21060000": "2300200",
        "32080000": "2400100",
        "32090000": "2400100",
        "32100000": "2400100",
        "28210000": "2400200",
        "32041700": "2400200",
        "32060000": "2400200",
        "87021000": "2500100",
        "87029090": "2500200",
        "87032100": "2500300",
        "87032210": "2500400",
        "87032290": "2500500",
        "87032310": "2500600",
        "87032390": "2500700",
        "87032410": "2500800",
        "87032490": "2500900",
        "87033210": "2501000",
        "87033290": "2501100",
        "87033310": "2501200",
        "87033390": "2501300",
        "87042110": "2501400",
        "87042120": "2501500",
        "87042130": "2501600",
        "87042190": "2501700",
        "87043110": "2501800",
        "87043120": "2501900",
        "87043130": "2502000",
        "87043190": "2502100",
        "87110000": "2600100",
        "70090000": "2700100",
        "70130000": "2700200",
        "70133700": "2700300",
        "70134290": "2700400",
        "96032900": "2802700",
        "96161000": "2802900",
        "33340000": "2803500",
        "44646582": "2803600",
        "39424871": "2803700",
        "61626400": "2803800",
        "42525558": "2803900",
        "39405663": "2804000",
        "13152300": "2804100",
        "22272829": "2804300"
}

        # Tentar carregar arquivo customizado se existir ou permitir importação via Excel
        custom_cest_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cest_mapping_custom.json")
        if os.path.exists(custom_cest_json):
            try:
                with open(custom_cest_json, "r", encoding="utf-8") as f:
                    cest_mapping.update(json.load(f))
            except Exception as e:
                print(f"Aviso ao carregar cest_mapping_custom.json: {e}")

        registros_atualizados = 0

        for i, line in enumerate(self.data):
            fields = line.strip().split("|")
            # Valida se é um registro 0200
            if len(fields) > 1 and (fields[1] == "0200" or fields[0] == "0200"):
                idx_ncm = 8
                idx_cest = 13

                if len(fields) > idx_ncm:
                    raw_ncm = re.sub(r"\D", "", fields[idx_ncm])
                    if raw_ncm:
                        # Busca por NCM exato (8 dígitos) ou prefixo de 6/4 dígitos
                        cest_encontrado = cest_mapping.get(raw_ncm)
                        if not cest_encontrado and len(raw_ncm) >= 6:
                            cest_encontrado = cest_mapping.get(raw_ncm[:6] + "00")
                        if not cest_encontrado and len(raw_ncm) >= 4:
                            cest_encontrado = cest_mapping.get(raw_ncm[:4] + "0000")

                        if cest_encontrado:
                            while len(fields) <= idx_cest:
                                fields.append("")
                            if fields[idx_cest] != cest_encontrado:
                                fields[idx_cest] = cest_encontrado
                                self.data[i] = "|".join(fields) + "\n"
                                registros_atualizados += 1

        self.display_data()
        self.is_modified = True
        self.update_window_title()
        messagebox.showinfo("Atualização CEST Concluída", f"Total de {registros_atualizados} registro(s) 0200 atualizado(s) com o código CEST segundo a tabela CEST12.xlsx.")

    def corrigir_0200_ncm_vazio_ou_zerado(self):
        """Localiza registros 0200 onde o Campo 8 (NCM) esteja vazio ou com '00' e abre janela interativa para correção."""
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED aberto para verificar.")
            return

        inconsistencias = []
        for idx, line in enumerate(self.data):
            line_str = line.strip()
            if not line_str:
                continue
            fields = line_str.split("|")
            idx_off = 1 if fields[0] == "" else 0
            reg = fields[idx_off] if len(fields) > idx_off else ""
            reg_clean = reg.strip().upper()
            if reg_clean.isdigit() and len(reg_clean) < 4:
                reg_clean = reg_clean.zfill(4)

            if reg_clean == "0200":
                # Layout do registro 0200:
                # Campo 1 (idx_off + 0): REG (0200)
                # Campo 2 (idx_off + 1): COD_ITEM
                # Campo 3 (idx_off + 2): DESCR_ITEM
                # Campo 4 (idx_off + 3): COD_BARRA
                # Campo 5 (idx_off + 4): COD_ANT_ITEM
                # Campo 6 (idx_off + 5): UNID_INV
                # Campo 7 (idx_off + 6): TIPO_ITEM
                # Campo 8 (idx_off + 7): COD_NCM
                cod_item = fields[idx_off + 1].strip() if len(fields) > (idx_off + 1) else ""
                descr_item = fields[idx_off + 2].strip() if len(fields) > (idx_off + 2) else ""
                unid_inv = fields[idx_off + 5].strip() if len(fields) > (idx_off + 5) else ""
                tipo_item = fields[idx_off + 6].strip() if len(fields) > (idx_off + 6) else ""
                idx_ncm = idx_off + 7
                cod_ncm = fields[idx_ncm].strip() if len(fields) > idx_ncm else ""

                ncm_digits = re.sub(r"\D", "", cod_ncm)

                # É considerado inválido se:
                # 1. Campo vazio ou apenas espaços
                # 2. Igual a '00', '0', '00000000', ou formado apenas por zeros
                # 3. Valor não alfanumérico ou com menos de 2 dígitos
                eh_invalido = False
                if not cod_ncm or cod_ncm == "":
                    eh_invalido = True
                elif cod_ncm in ("00", "0", "00000000", "0000", "00.00.00.00", "0000.00.00", "00.00", "000"):
                    eh_invalido = True
                elif ncm_digits in ("00", "0", "00000000", "0000", "000") or (len(ncm_digits) > 0 and all(c == '0' for c in ncm_digits)):
                    eh_invalido = True
                elif len(ncm_digits) < 2 and not cod_ncm.isalnum():
                    eh_invalido = True

                if eh_invalido:
                    inconsistencias.append({
                        "id": len(inconsistencias),
                        "linha_idx": idx,
                        "linha_num": idx + 1,
                        "cod_item": cod_item,
                        "descr_item": descr_item,
                        "unid_inv": unid_inv,
                        "tipo_item": tipo_item,
                        "ncm_atual": cod_ncm if cod_ncm else "(Vazio)",
                        "novo_ncm": "",
                        "idx_ncm": idx_ncm
                    })

        if not inconsistencias:
            messagebox.showinfo(
                "Registro 0200 - NCM (Campo 8)",
                "Nenhum registro 0200 com o Campo 8 (NCM) vazio ou zerado ('00') foi encontrado!\n\nTodos os itens cadastrados no arquivo possuem NCM preenchido."
            )
            return

        self._exibir_janela_correcao_0200_ncm(inconsistencias)

    def _exibir_janela_correcao_0200_ncm(self, inconsistencias):
        """Abre janela interativa para exibir os registros 0200 com NCM vazio/00 e permitir a definição do valor correto."""
        win = tk.Toplevel(self.root)
        win.title("0200 - Correção de NCM Vazio ou Zerado (Campo 8)")
        win.geometry("1020x640")
        win.minsize(850, 500)
        win.transient(self.root)
        win.grab_set()

        # Centralizar na tela
        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
        y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{x}+{y}")

        # Frame de Cabeçalho
        frame_header = tk.Frame(win, bg="#17324d", padx=16, pady=12, relief="groove", bd=1)
        frame_header.pack(fill=tk.X)

        lbl_titulo = tk.Label(
            frame_header,
            text=f"⚠️ {len(inconsistencias)} Registro(s) 0200 com Campo 8 (NCM) Vazio ou Zerado ('00')",
            font=("Segoe UI", 12, "bold"),
            fg="#fecaca",
            bg="#17324d"
        )
        lbl_titulo.pack(anchor="w")

        lbl_sub = tk.Label(
            frame_header,
            text="Selecione um item para informar o NCM correto, utilize o assistente passo a passo ou preencha em lote. Após definir, clique em 'Salvar e Aplicar no SPED'.",
            font=("Segoe UI", 9),
            fg="#cbd5e1",
            bg="#17324d"
        )
        lbl_sub.pack(anchor="w", pady=(3, 0))

        # Barra de Filtro e Contadores
        frame_toolbar = tk.Frame(win, padx=12, pady=6, bg="#f8fafc", relief="groove", bd=1)
        frame_toolbar.pack(fill=tk.X)

        tk.Label(frame_toolbar, text="🔍 Filtrar Código / Descrição:", font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#334155").pack(side=tk.LEFT, padx=(0, 6))
        entry_filtro = ttk.Entry(frame_toolbar, width=28)
        entry_filtro.pack(side=tk.LEFT, padx=(0, 15))

        lbl_status_contagem = tk.Label(
            frame_toolbar,
            text=f"Total: {len(inconsistencias)} | 🔴 Pendentes: {len(inconsistencias)} | 🟢 Definidos: 0",
            font=("Segoe UI", 9, "bold"),
            bg="#f8fafc",
            fg="#0f172a"
        )
        lbl_status_contagem.pack(side=tk.RIGHT, padx=5)

        # Tabela Treeview
        frame_table = tk.Frame(win, padx=12, pady=6)
        frame_table.pack(fill=tk.BOTH, expand=True)

        cols = ("id", "linha", "cod_item", "descr_item", "unid", "tipo", "ncm_atual", "novo_ncm", "status")
        tree = ttk.Treeview(frame_table, columns=cols, show="headings", selectmode="extended")
        
        tree.heading("id", text="#")
        tree.heading("linha", text="Linha SPED")
        tree.heading("cod_item", text="Código do Item (Campo 2)")
        tree.heading("descr_item", text="Descrição do Produto / Serviço (Campo 3)")
        tree.heading("unid", text="Unid.")
        tree.heading("tipo", text="Tipo")
        tree.heading("ncm_atual", text="NCM Atual")
        tree.heading("novo_ncm", text="Novo NCM (Correto)")
        tree.heading("status", text="Status")

        tree.column("id", width=35, anchor="center")
        tree.column("linha", width=75, anchor="center")
        tree.column("cod_item", width=140, anchor="w")
        tree.column("descr_item", width=320, anchor="w")
        tree.column("unid", width=50, anchor="center")
        tree.column("tipo", width=50, anchor="center")
        tree.column("ncm_atual", width=90, anchor="center")
        tree.column("novo_ncm", width=110, anchor="center")
        tree.column("status", width=95, anchor="center")

        tree.tag_configure("pendente", background="#fff5f5", foreground="#991b1b")
        tree.tag_configure("definido", background="#f0fdf4", foreground="#166534")

        vsb = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(frame_table, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame_table.grid_rowconfigure(0, weight=1)
        frame_table.grid_columnconfigure(0, weight=1)

        def atualizar_contadores():
            definidos = sum(1 for item in inconsistencias if item.get("novo_ncm", "").strip())
            pendentes = len(inconsistencias) - definidos
            lbl_status_contagem.config(
                text=f"Total: {len(inconsistencias)} | 🔴 Pendentes: {pendentes} | 🟢 Definidos: {definidos}"
            )

        def preencher_tabela(filtro=""):
            for item_id in tree.get_children():
                tree.delete(item_id)
            termo = filtro.strip().upper()
            for item in inconsistencias:
                if termo:
                    c_item = item["cod_item"].upper()
                    d_item = item["descr_item"].upper()
                    if termo not in c_item and termo not in d_item:
                        continue
                novo = item.get("novo_ncm", "").strip()
                status_txt = "🟢 Definido" if novo else "🔴 Pendente"
                tag = "definido" if novo else "pendente"
                tree.insert(
                    "",
                    tk.END,
                    iid=str(item["id"]),
                    values=(
                        item["id"] + 1,
                        item["linha_num"],
                        item["cod_item"],
                        item["descr_item"],
                        item["unid_inv"],
                        item["tipo_item"],
                        item["ncm_atual"],
                        novo if novo else "-",
                        status_txt
                    ),
                    tags=(tag,)
                )
            atualizar_contadores()

        entry_filtro.bind("<KeyRelease>", lambda e: preencher_tabela(entry_filtro.get()))

        def abrir_dialogo_edicao_item(item_dict):
            """Abre janela para definir o NCM correto de um item específico com validação TIPI."""
            dlg = tk.Toplevel(win)
            dlg.title("Definir Valor Correto do Campo 8 (NCM)")
            dlg.geometry("560x360")
            dlg.resizable(False, False)
            dlg.transient(win)
            dlg.grab_set()

            # Centralizar
            dlg.update_idletasks()
            dx = (win.winfo_x() + (win.winfo_width() // 2)) - 280
            dy = (win.winfo_y() + (win.winfo_height() // 2)) - 180
            dlg.geometry(f"+{dx}+{dy}")

            f_top = tk.Frame(dlg, bg="#1e293b", padx=14, pady=10)
            f_top.pack(fill=tk.X)
            tk.Label(
                f_top,
                text="Qual o valor correto do Campo 8 (NCM)?",
                font=("Segoe UI", 11, "bold"),
                fg="white",
                bg="#1e293b"
            ).pack(anchor="w")

            f_body = tk.Frame(dlg, padx=16, pady=12)
            f_body.pack(fill=tk.BOTH, expand=True)

            # Informações do Item
            f_info = tk.LabelFrame(f_body, text=" Dados do Item (Registro 0200) ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
            f_info.pack(fill=tk.X, pady=(0, 10))

            tk.Label(f_info, text=f"Linha no SPED: {item_dict['linha_num']}", font=("Segoe UI", 8), fg="#64748b").pack(anchor="w")
            tk.Label(f_info, text=f"Código do Item: {item_dict['cod_item']}", font=("Segoe UI", 9, "bold"), fg="#0f172a").pack(anchor="w")
            tk.Label(f_info, text=f"Descrição: {item_dict['descr_item']}", font=("Segoe UI", 9), fg="#334155", wraplength=480, justify="left").pack(anchor="w", pady=(2, 0))
            tk.Label(f_info, text=f"Unid: {item_dict['unid_inv']}  |  Tipo: {item_dict['tipo_item']}  |  NCM Atual: {item_dict['ncm_atual']}", font=("Segoe UI", 8), fg="#64748b").pack(anchor="w", pady=(2, 0))

            # Entrada do Novo NCM
            f_input = tk.Frame(f_body)
            f_input.pack(fill=tk.X, pady=(4, 6))

            tk.Label(f_input, text="Novo NCM (8 dígitos):", font=("Segoe UI", 10, "bold"), fg="#0f172a").pack(side=tk.LEFT, padx=(0, 8))
            entry_ncm = ttk.Entry(f_input, font=("Segoe UI", 11, "bold"), width=16)
            entry_ncm.pack(side=tk.LEFT, padx=(0, 10))
            if item_dict.get("novo_ncm"):
                entry_ncm.insert(0, item_dict["novo_ncm"])
            entry_ncm.focus_set()

            lbl_tipi_desc = tk.Label(
                f_body,
                text="Digite o código NCM acima (ex: 84818099) para consultar a TIPI.",
                font=("Segoe UI", 8, "italic"),
                fg="#64748b",
                wraplength=490,
                justify="left"
            )
            lbl_tipi_desc.pack(anchor="w", pady=(0, 10))

            def validar_tipi_live(e=None):
                val = re.sub(r"\D", "", entry_ncm.get().strip())
                if len(val) >= 4:
                    res = consultar_ncm_tipi(val)
                    if res.get("encontrado") and res.get("registro"):
                        reg = res["registro"]
                        lbl_tipi_desc.config(
                            text=f"✅ TIPI: {reg['DESCRICAO']} (Alíq. IPI: {reg['ALIQUOTA']}%)",
                            fg="#15803d"
                        )
                    else:
                        lbl_tipi_desc.config(
                            text=f"ℹ️ NCM {val} não encontrada na tabela TIPI local.",
                            fg="#b45309"
                        )
                else:
                    lbl_tipi_desc.config(
                        text="Digite ao menos 4 a 8 dígitos do NCM para consulta.",
                        fg="#64748b"
                    )

            entry_ncm.bind("<KeyRelease>", validar_tipi_live)
            if item_dict.get("novo_ncm"):
                validar_tipi_live()

            resultado_salvo = [False]

            def confirmar():
                valor = re.sub(r"\D", "", entry_ncm.get().strip())
                if not valor:
                    messagebox.showwarning("Aviso", "Por favor, digite o código NCM desejado.", parent=dlg)
                    return
                item_dict["novo_ncm"] = valor
                resultado_salvo[0] = True
                dlg.destroy()

            f_btns = tk.Frame(dlg, bg="#f1f5f9", padx=12, pady=10)
            f_btns.pack(fill=tk.X, side=tk.BOTTOM)

            tk.Button(
                f_btns,
                text="✔️ Confirmar NCM",
                font=("Segoe UI", 9, "bold"),
                bg="#15803d",
                fg="white",
                padx=12,
                pady=4,
                command=confirmar
            ).pack(side=tk.RIGHT, padx=5)

            tk.Button(
                f_btns,
                text="Cancelar",
                font=("Segoe UI", 9),
                bg="#64748b",
                fg="white",
                padx=10,
                pady=4,
                command=dlg.destroy
            ).pack(side=tk.RIGHT, padx=5)

            dlg.bind("<Return>", lambda e: confirmar())
            dlg.bind("<Escape>", lambda e: dlg.destroy())

            win.wait_window(dlg)
            return resultado_salvo[0]

        def editar_selecionado():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Aviso", "Selecione um item na lista para editar.", parent=win)
                return
            item_id = int(sel[0])
            item = inconsistencias[item_id]
            if abrir_dialogo_edicao_item(item):
                preencher_tabela(entry_filtro.get())
                if tree.exists(str(item_id)):
                    tree.selection_set(str(item_id))
                    tree.see(str(item_id))

        tree.bind("<Double-1>", lambda e: editar_selecionado())

        def assistente_passo_a_passo():
            """Executa assistente iterativo perguntando um a um o valor correto para todos os itens pendentes."""
            pendentes = [it for it in inconsistencias if not it.get("novo_ncm", "").strip()]
            if not pendentes:
                resp = messagebox.askyesno(
                    "Assistente",
                    "Todos os itens já possuem NCM definido. Deseja revisar todos os itens novamente?",
                    parent=win
                )
                if not resp:
                    return
                pendentes = inconsistencias

            total = len(pendentes)
            alterados = 0

            for i, item in enumerate(pendentes):
                # Seleciona visualmente na tabela
                if tree.exists(str(item["id"])):
                    tree.selection_set(str(item["id"]))
                    tree.see(str(item["id"]))

                # Modal com opção de Pular, Salvar ou Cancelar assistente
                dlg = tk.Toplevel(win)
                dlg.title(f"Assistente Passo a Passo ({i+1} de {total})")
                dlg.geometry("580x380")
                dlg.resizable(False, False)
                dlg.transient(win)
                dlg.grab_set()

                # Centralizar
                dlg.update_idletasks()
                dx = (win.winfo_x() + (win.winfo_width() // 2)) - 290
                dy = (win.winfo_y() + (win.winfo_height() // 2)) - 190
                dlg.geometry(f"+{dx}+{dy}")

                f_top = tk.Frame(dlg, bg="#6d28d9", padx=14, pady=10)
                f_top.pack(fill=tk.X)
                tk.Label(
                    f_top,
                    text=f"Item {i+1} de {total}: Qual o valor correto do Campo 8 (NCM)?",
                    font=("Segoe UI", 11, "bold"),
                    fg="white",
                    bg="#6d28d9"
                ).pack(anchor="w")

                f_body = tk.Frame(dlg, padx=16, pady=12)
                f_body.pack(fill=tk.BOTH, expand=True)

                f_info = tk.LabelFrame(f_body, text=" Registro 0200 ", font=("Segoe UI", 9, "bold"), padx=10, pady=8)
                f_info.pack(fill=tk.X, pady=(0, 10))

                tk.Label(f_info, text=f"Linha SPED: {item['linha_num']}", font=("Segoe UI", 8), fg="#64748b").pack(anchor="w")
                tk.Label(f_info, text=f"Código do Item: {item['cod_item']}", font=("Segoe UI", 10, "bold"), fg="#0f172a").pack(anchor="w")
                tk.Label(f_info, text=f"Descrição: {item['descr_item']}", font=("Segoe UI", 9), fg="#334155", wraplength=500, justify="left").pack(anchor="w", pady=(2, 0))
                tk.Label(f_info, text=f"Unid: {item['unid_inv']}  |  Tipo: {item['tipo_item']}  |  NCM Atual: {item['ncm_atual']}", font=("Segoe UI", 8), fg="#64748b").pack(anchor="w", pady=(2, 0))

                f_input = tk.Frame(f_body)
                f_input.pack(fill=tk.X, pady=(4, 6))

                tk.Label(f_input, text="NCM Correto (8 dígitos):", font=("Segoe UI", 10, "bold"), fg="#0f172a").pack(side=tk.LEFT, padx=(0, 8))
                entry_ncm = ttk.Entry(f_input, font=("Segoe UI", 11, "bold"), width=16)
                entry_ncm.pack(side=tk.LEFT, padx=(0, 10))
                if item.get("novo_ncm"):
                    entry_ncm.insert(0, item["novo_ncm"])
                entry_ncm.focus_set()

                lbl_tipi_desc = tk.Label(
                    f_body,
                    text="Digite o código NCM para consultar a tabela TIPI.",
                    font=("Segoe UI", 8, "italic"),
                    fg="#64748b",
                    wraplength=500,
                    justify="left"
                )
                lbl_tipi_desc.pack(anchor="w", pady=(0, 10))

                def validar_tipi_assistente(e=None):
                    val = re.sub(r"\D", "", entry_ncm.get().strip())
                    if len(val) >= 4:
                        res = consultar_ncm_tipi(val)
                        if res.get("encontrado") and res.get("registro"):
                            reg = res["registro"]
                            lbl_tipi_desc.config(
                                text=f"✅ TIPI: {reg['DESCRICAO']} (Alíq. IPI: {reg['ALIQUOTA']}%)",
                                fg="#15803d"
                            )
                        else:
                            lbl_tipi_desc.config(
                                text=f"ℹ️ NCM {val} não encontrada na tabela TIPI local.",
                                fg="#b45309"
                            )
                    else:
                        lbl_tipi_desc.config(
                            text="Digite ao menos 4 a 8 dígitos para consultar a TIPI.",
                            fg="#64748b"
                        )

                entry_ncm.bind("<KeyRelease>", validar_tipi_assistente)
                if item.get("novo_ncm"):
                    validar_tipi_assistente()

                acao = {"status": "cancelar"}

                def confirmar():
                    valor = re.sub(r"\D", "", entry_ncm.get().strip())
                    if not valor:
                        messagebox.showwarning("Aviso", "Por favor, digite o código NCM ou clique em 'Pular'.", parent=dlg)
                        return
                    item["novo_ncm"] = valor
                    acao["status"] = "ok"
                    dlg.destroy()

                def pular():
                    acao["status"] = "pular"
                    dlg.destroy()

                def cancelar():
                    acao["status"] = "cancelar"
                    dlg.destroy()

                f_btns = tk.Frame(dlg, bg="#f1f5f9", padx=12, pady=10)
                f_btns.pack(fill=tk.X, side=tk.BOTTOM)

                tk.Button(f_btns, text="✔️ Confirmar", font=("Segoe UI", 9, "bold"), bg="#15803d", fg="white", padx=12, pady=4, command=confirmar).pack(side=tk.RIGHT, padx=4)
                tk.Button(f_btns, text="⏭️ Pular", font=("Segoe UI", 9), bg="#d97706", fg="white", padx=10, pady=4, command=pular).pack(side=tk.RIGHT, padx=4)
                tk.Button(f_btns, text="Encerrar Assistente", font=("Segoe UI", 9), bg="#64748b", fg="white", padx=8, pady=4, command=cancelar).pack(side=tk.LEFT, padx=4)

                dlg.bind("<Return>", lambda e: confirmar())
                dlg.bind("<Escape>", lambda e: cancelar())

                win.wait_window(dlg)

                if acao["status"] == "ok":
                    alterados += 1
                    preencher_tabela(entry_filtro.get())
                elif acao["status"] == "cancelar":
                    break

            preencher_tabela(entry_filtro.get())
            messagebox.showinfo("Assistente", f"Assistente finalizado! {alterados} item(ns) tiveram o NCM definido.", parent=win)

        def preencher_em_lote():
            """Permite aplicar um código NCM padrão a todos os itens selecionados ou pendentes."""
            sel = tree.selection()
            alvos = []
            if sel:
                for item_id in sel:
                    alvos.append(inconsistencias[int(item_id)])
            else:
                resp = messagebox.askyesno(
                    "Preencher em Lote",
                    f"Nenhum item selecionado individualmente.\nDeseja aplicar o NCM a todos os {len(inconsistencias)} itens da lista?",
                    parent=win
                )
                if not resp:
                    return
                alvos = inconsistencias

            ncm_input = simpledialog.askstring(
                "Preencher NCM em Lote",
                f"Digite o código NCM correto para aplicar aos {len(alvos)} itens selecionados:",
                parent=win
            )
            if not ncm_input:
                return

            ncm_limpo = re.sub(r"\D", "", ncm_input.strip())
            if not ncm_limpo:
                messagebox.showwarning("Aviso", "Código NCM inválido.", parent=win)
                return

            for it in alvos:
                it["novo_ncm"] = ncm_limpo

            preencher_tabela(entry_filtro.get())
            messagebox.showinfo("Sucesso", f"NCM {ncm_limpo} aplicado a {len(alvos)} item(ns) com sucesso!", parent=win)

        def consultar_tipi():
            """Abre janela de consulta rápida à tabela TIPI."""
            dlg = tk.Toplevel(win)
            dlg.title("Consulta à Tabela TIPI / NCM")
            dlg.geometry("700x480")
            dlg.transient(win)

            f_t = tk.Frame(dlg, bg="#1e293b", padx=12, pady=8)
            f_t.pack(fill=tk.X)
            tk.Label(f_t, text="Consulta de Códigos NCM / TIPI", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e293b").pack(anchor="w")

            f_s = tk.Frame(dlg, padx=10, pady=8)
            f_s.pack(fill=tk.X)
            tk.Label(f_s, text="Buscar Termo / NCM:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
            ent_b = ttk.Entry(f_s, width=30)
            ent_b.pack(side=tk.LEFT, padx=(0, 8))

            f_res = tk.Frame(dlg, padx=10, pady=6)
            f_res.pack(fill=tk.BOTH, expand=True)

            cols_tipi = ("ncm", "ex", "aliq", "descr")
            tree_tipi = ttk.Treeview(f_res, columns=cols_tipi, show="headings")
            tree_tipi.heading("ncm", text="NCM")
            tree_tipi.heading("ex", text="EX")
            tree_tipi.heading("aliq", text="Alíq. IPI (%)")
            tree_tipi.heading("descr", text="Descrição do Produto")

            tree_tipi.column("ncm", width=100, anchor="center")
            tree_tipi.column("ex", width=45, anchor="center")
            tree_tipi.column("aliq", width=80, anchor="center")
            tree_tipi.column("descr", width=420, anchor="w")

            vsb_t = ttk.Scrollbar(f_res, orient=tk.VERTICAL, command=tree_tipi.yview)
            tree_tipi.configure(yscrollcommand=vsb_t.set)
            tree_tipi.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            vsb_t.pack(side=tk.RIGHT, fill=tk.Y)

            def buscar_tipi(e=None):
                termo = ent_b.get().strip().upper()
                termo_limpo = re.sub(r"\D", "", termo)
                for i in tree_tipi.get_children():
                    tree_tipi.delete(i)
                caminho_db = localizar_tipi_db()
                if not caminho_db or not os.path.exists(str(caminho_db)):
                    return
                try:
                    conn = sqlite3.connect(str(caminho_db))
                    c = conn.cursor()
                    if termo_limpo and len(termo_limpo) >= 2:
                        c.execute("SELECT NCM, EX, ALIQUOTA, DESCRICAO FROM Tipi WHERE REPLACE(REPLACE(NCM, '.', ''), ' ', '') LIKE ? OR UPPER(DESCRICAO) LIKE ? LIMIT 100", (f"%{termo_limpo}%", f"%{termo}%"))
                    elif termo:
                        c.execute("SELECT NCM, EX, ALIQUOTA, DESCRICAO FROM Tipi WHERE UPPER(DESCRICAO) LIKE ? LIMIT 100", (f"%{termo}%",))
                    else:
                        c.execute("SELECT NCM, EX, ALIQUOTA, DESCRICAO FROM Tipi LIMIT 100")
                    rows = c.fetchall()
                    conn.close()
                    for r in rows:
                        tree_tipi.insert("", tk.END, values=(r[0], r[1] or "-", r[2] or "0", r[3]))
                except Exception as ex:
                    print(f"Erro ao buscar TIPI: {ex}")

            ent_b.bind("<KeyRelease>", buscar_tipi)
            tk.Button(f_s, text="Buscar", font=("Segoe UI", 9, "bold"), bg="#1d4ed8", fg="white", command=buscar_tipi).pack(side=tk.LEFT)
            buscar_tipi()

        def salvar_e_aplicar():
            """Aplica todas as alterações no arquivo SPED carregado."""
            definidos = [it for it in inconsistencias if it.get("novo_ncm", "").strip()]
            if not definidos:
                messagebox.showwarning(
                    "Aviso",
                    "Nenhum item teve o novo NCM definido ainda.\nInforme o NCM correto para ao menos um item antes de aplicar.",
                    parent=win
                )
                return

            pendentes = len(inconsistencias) - len(definidos)
            if pendentes > 0:
                conf = messagebox.askyesno(
                    "Confirmar Aplicação Parcial",
                    f"Você definiu o NCM para {len(definidos)} item(ns), mas ainda restam {pendentes} item(ns) pendentes (que permanecerão sem alteração).\n\nDeseja aplicar as {len(definidos)} alterações no arquivo SPED agora?",
                    parent=win
                )
                if not conf:
                    return
            else:
                conf = messagebox.askyesno(
                    "Confirmar Aplicação",
                    f"Deseja aplicar o NCM correto nos {len(definidos)} registros 0200 do arquivo SPED?",
                    parent=win
                )
                if not conf:
                    return

            # Aplicar alterações em self.data
            alterados = 0
            for it in definidos:
                l_idx = it["linha_idx"]
                if l_idx < len(self.data):
                    line = self.data[l_idx].strip()
                    fields = line.split("|")
                    idx_ncm = it["idx_ncm"]
                    while len(fields) <= idx_ncm:
                        fields.append("")
                    fields[idx_ncm] = it["novo_ncm"]
                    self.data[l_idx] = "|".join(fields) + "\n"
                    alterados += 1

            self.display_data()
            self.is_modified = True
            self.update_window_title()
            win.destroy()
            messagebox.showinfo(
                "Sucesso",
                f"Total de {alterados} registro(s) 0200 foram atualizados com sucesso no Campo 8 (NCM)!"
            )

        # Botões de Ação Inferiores
        frame_actions = tk.Frame(win, bg="#f1f5f9", padx=12, pady=10, relief="groove", bd=1)
        frame_actions.pack(fill=tk.X, side=tk.BOTTOM)

        tk.Button(
            frame_actions,
            text="💾 Salvar e Aplicar no SPED",
            font=("Segoe UI", 9, "bold"),
            bg="#15803d",
            fg="white",
            padx=12,
            pady=5,
            command=salvar_e_aplicar
        ).pack(side=tk.RIGHT, padx=6)

        tk.Button(
            frame_actions,
            text="✏️ Informar Valor (Item)",
            font=("Segoe UI", 9, "bold"),
            bg="#1d4ed8",
            fg="white",
            padx=10,
            pady=5,
            command=editar_selecionado
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            frame_actions,
            text="🤖 Assistente (Perguntar 1 a 1)",
            font=("Segoe UI", 9, "bold"),
            bg="#6d28d9",
            fg="white",
            padx=10,
            pady=5,
            command=assistente_passo_a_passo
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            frame_actions,
            text="📦 Preencher em Lote",
            font=("Segoe UI", 9),
            bg="#0891b2",
            fg="white",
            padx=10,
            pady=5,
            command=preencher_em_lote
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            frame_actions,
            text="🔍 Consultar Tabela TIPI",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="white",
            padx=10,
            pady=5,
            command=consultar_tipi
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            frame_actions,
            text="Fechar",
            font=("Segoe UI", 9),
            bg="#94a3b8",
            fg="white",
            padx=10,
            pady=5,
            command=win.destroy
        ).pack(side=tk.RIGHT, padx=4)

        # Inicializa a tabela
        preencher_tabela()

    def update_0200_field_7(self):
        """Atualiza o campo 7 do registro 0200 (Tipo de Item) com base nas regras de classificação."""
        alteracoes = 0
        def modify_line(line):
            nonlocal alteracoes
            fields = line.strip().split("|")
            if len(fields) > 8 and (fields[1] == "0200" or fields[0] == "0200"):
                idx_offset = 1 if fields[0] == "" else 0
                campo_2 = fields[idx_offset + 2] if len(fields) > (idx_offset + 2) else ""
                campo_6 = fields[idx_offset + 6].strip().upper() if len(fields) > (idx_offset + 6) else ""
                campo_8 = fields[idx_offset + 8] if len(fields) > (idx_offset + 8) else ""
                idx_c7 = idx_offset + 7
                orig_c7 = fields[idx_c7] if len(fields) > idx_c7 else ""

                if len(campo_2) >= 8 and campo_2[:8] == campo_8:
                    fields[idx_c7] = "99"
                else:
                    primeiro_digito = campo_2[0] if len(campo_2) > 0 else ""
                    if primeiro_digito == "1":
                        fields[idx_c7] = "01"
                    elif primeiro_digito == "2":
                        fields[idx_c7] = "06"
                    elif primeiro_digito == "3":
                        fields[idx_c7] = "03"
                    elif primeiro_digito in ["6", "7"]:
                        fields[idx_c7] = "00"
                if campo_2[:2] in ("A0", "B0", "B1", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "C0", "D0", "E0", "J1", "J3", "J4", "J5", "J6", "J7", "J8", "J9", "F0", "G0", "H0", "I0", "J0", "K0", "L0", "M0", "N0", "O0", "P0", "Q0", "R0", "S0", "T1", "U0", "V0", "W0", "X0", "Y0", "Z0"):
                    fields[idx_c7] = "04"
                if campo_6 == "SV":
                    fields[idx_c7] = "09"

                if fields[idx_c7] != orig_c7:
                    alteracoes += 1
            return "|".join(fields) + "\n"
        self.data = [modify_line(line) for line in self.data]
        self.display_data()
        self.is_modified = True
        self.update_window_title()
        messagebox.showinfo("Sucesso", f"Total de {alteracoes} registro(s) 0200 alterado(s) no Campo 7 com base nas regras.")

    def pesquisar_e_substituir(self):
        search_window = tk.Toplevel(self.root)
        search_window.title("Pesquisar e Substituir")
        frame = tk.Frame(search_window)
        frame.pack(padx=10, pady=10)
        tk.Label(frame, text="Valor a ser localizado:").grid(row=0, column=0, padx=5, pady=5)
        entry_search = tk.Entry(frame)
        entry_search.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(frame, text="Coluna:").grid(row=1, column=0, padx=5, pady=5)
        entry_column = tk.Entry(frame)
        entry_column.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(frame, text="Valor de substituição:").grid(row=2, column=0, padx=5, pady=5)
        entry_replace = tk.Entry(frame)
        entry_replace.grid(row=2, column=1, padx=5, pady=5)
        btn_confirm = tk.Button(frame, text="Confirmar", command=lambda: self._substituir_valor(entry_search.get(), entry_column.get(), entry_replace.get()))
        btn_confirm.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

    def _substituir_valor(self, search_value, column, replace_value):
        try:
            column = int(column)
        except ValueError:
            messagebox.showerror("Erro", "A coluna deve ser um número inteiro.")
            return
        # A verificação da coluna foi ajustada para ser mais robusta, já que não temos 'self.headers' no contexto
        # Simplificando a lógica para o mínimo de funcionalidade:
        lines_altered = 0
        for i, line in enumerate(self.data):
            fields = line.strip().split("|")
            if column >= 1 and column < len(fields):
                # O índice da lista é coluna - 1
                if fields[column] == search_value:
                    fields[column] = replace_value
                    self.data[i] = "|".join(fields) + "\n"
                    lines_altered += 1
        
        self.display_data() # Atualiza a exibição após as alterações
        messagebox.showinfo("Sucesso", f"Valor '{search_value}' substituído por '{replace_value}' na coluna {column}. Total de {lines_altered} linhas alteradas.")

    def insert_line(self):
        def submit_new_line():
            try:
                new_line = entry_new_line.get().strip()
                line_number = int(entry_line_number.get().strip())
                if not new_line:
                    raise ValueError("A nova linha não pode ser vazia.")
                if line_number < 1 or line_number > len(self.data) + 1:
                    raise ValueError("Número de linha inválido.")
                self.inserted_lines.insert(line_number - 1, [line_number, new_line])
                self.data.insert(line_number - 1, new_line + "\n")
                self.display_data()
                messagebox.showinfo("Sucesso", f"Linha {line_number} inserida com sucesso.")
            except ValueError as e:
                messagebox.showerror("Erro", str(e))
            insert_window.destroy()

        insert_window = tk.Toplevel(self.root)
        insert_window.title("Inserir Linha")
        tk.Label(insert_window, text="Linhas Disponíveis:").pack(pady=5)
        text_widget = tk.Text(insert_window, width=80, height=15, wrap=tk.WORD)
        text_widget.pack(pady=5)
        # O código original não implementava a exibição de linhas inseridas corretamente, 
        # mas vou manter a intenção
        available_lines = "\n".join([f"Linha {idx}: {line.strip()}" for idx, line in enumerate(self.data, 1)])
        text_widget.insert(tk.END, available_lines)
        text_widget.config(state=tk.DISABLED)
        tk.Label(insert_window, text="Digite a nova linha:").pack(pady=5)
        entry_new_line = tk.Entry(insert_window, width=80)
        entry_new_line.pack(pady=5)
        tk.Label(insert_window, text="Digite o número da linha onde deseja inserir:").pack(pady=5)
        entry_line_number = tk.Entry(insert_window, width=10)
        entry_line_number.pack(pady=5)
        btn_submit = tk.Button(insert_window, text="Inserir", command=submit_new_line)
        btn_submit.pack(pady=5)

    def strip_spaces_between_pipes(self):
        def clean_line(line):
            fields = line.strip().split("|")
            cleaned_fields = [field.strip() for field in fields]
            return "|".join(cleaned_fields) + "\n"
        self.data = [clean_line(line) for line in self.data]
        self.display_data()
        messagebox.showinfo("Sucesso", "Espaços extras entre os delimitadores foram removidos.")

    def apply_changes(self):
        self.data = [line for line in self.data if line is not None]
        # O código original não tinha as listas 'deleted_lines' e 'inserted_lines' inicializadas 
        # fora do load_file, o que pode causar erros. Assumindo que elas existem:
        total_deleted = len(getattr(self, 'deleted_lines', []))
        total_inserted = len(getattr(self, 'inserted_lines', []))
        total_changes = total_deleted + total_inserted
        
        # Limpar as listas (assumindo que o load_file as inicializa ou que a classe é reiniciada)
        if hasattr(self, 'deleted_lines'):
            self.deleted_lines.clear()
        if hasattr(self, 'inserted_lines'):
            self.inserted_lines.clear()
        
        self.display_data()
        self.strip_spaces_between_pipes()
        messagebox.showinfo("Alterações Aplicadas", f"Total de alterações aplicadas (exclusão/inserção): {total_changes}")

    def save_file(self):
        """Salva direto no caminho atual. Se não houver, pede caminho."""
        if not self.file_path:
            self.save_file_as()
            return
            
        try:
            with open(self.file_path, "w", encoding="latin-1") as file:
                file.writelines([line for line in self.data if line is not None])
            self.clear_modified()
            messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{self.file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")

    def save_file_as(self):
        """Abre janela para escolher onde salvar (antigo save_file)."""
        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if not save_path:
            return
        
        try:
            with open(save_path, "w", encoding="latin-1") as file:
                file.writelines([line for line in self.data if line is not None])
            
            self.file_path = save_path
            self.clear_modified()
            self.add_to_recent(save_path)
            self.lbl_status.config(text=f"Arquivo carregado: {self.file_path}")
            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {save_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")
            
    def gerar_relatorio_totalizador_simples(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado.")
            return

        output_file = filedialog.asksaveasfilename(
            title="Salvar relatório SPED como",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not output_file:
            messagebox.showinfo("Cancelado", "Geração do relatório cancelada.")
            return

        registros_c190 = []
        notas_canceladas = set()
        numeros_notas_saida = set()
        
        for line in self.data:
            try:
                campos = line.strip().split('|')
                if not campos or len(campos) < 2:
                    continue
                
                if campos[1] == 'C100':
                    if len(campos) > 8:
                        cod_sit = campos[6]
                        ind_oper = campos[2]
                        try:
                            numero_nota = int(campos[8])
                            if cod_sit == '02':
                                notas_canceladas.add(numero_nota)
                            if ind_oper == '1':
                                numeros_notas_saida.add(numero_nota)
                        except (ValueError, IndexError):
                            continue
                
                elif campos[1] == 'C190':
                    registros_c190.append(campos)
            
            except (ValueError, IndexError) as e:
                print(f"Aviso: Linha mal formatada ou incompleta foi ignorada. Erro: {e}")
                continue

        if not registros_c190:
            messagebox.showwarning("Aviso", "Nenhum registro C190 encontrado no arquivo.")
            return

        faltantes = []
        if numeros_notas_saida:
            inicio = min(numeros_notas_saida)
            fim = max(numeros_notas_saida)
            numeros_considerados = numeros_notas_saida - notas_canceladas
            faltantes = sorted(list(set(range(inicio, fim + 1)) - numeros_considerados))

        cfop_data = agrupar_por_cfop(registros_c190)

        try:
            gerar_pdf(cfop_data, list(notas_canceladas), faltantes, len(self.data), output_file, self.data)
            messagebox.showinfo("Sucesso", f"Relatório gerado em: {output_file}")
            os.startfile(output_file)
        except Exception as e:
            messagebox.showerror("Erro ao gerar PDF", f"Ocorreu um erro ao criar o relatório: {e}")


    def gerar_relatorio_totalizador(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado.")
            return

        output_file = filedialog.asksaveasfilename(
            title="Salvar relatorio comparador como",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not output_file:
            messagebox.showinfo("Cancelado", "Geracao do relatorio cancelada.")
            return

        totais_sped_info = consolidar_totais_sped(self.data)
        if not totais_sped_info:
            messagebox.showwarning("Aviso", "Nenhum registro C190 encontrado no arquivo.")
            return

        caminho_txt_padrao = r"C:\Users\Wesley.Raimundo\Desktop\vendas.txt"
        caminho_txt = caminho_txt_padrao if os.path.exists(caminho_txt_padrao) else filedialog.askopenfilename(
            title="Selecione o arquivo Athena para comparacao",
            filetypes=[("Arquivo texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not caminho_txt:
            messagebox.showinfo("Cancelado", "Selecao do arquivo Athena cancelada.")
            return

        output_xls = os.path.splitext(output_file)[0] + ".xls"

        try:
            athena_cfop = carregar_totais_athena_txt_v2(caminho_txt)
            athena_notas = carregar_athena_por_nota_txt(caminho_txt)
            athena_pis_cofins = carregar_pis_cofins_athena_por_nota_txt(caminho_txt)
            sped_cfop = totais_sped_info['cfop_data']
            sped_notas = carregar_sped_por_nota_cfop(self.data)
            sped_pis_cofins = carregar_pis_cofins_c100_por_nota(self.data)
            linhas_comparacao = montar_linhas_comparacao(sped_cfop, athena_cfop)
            linhas_por_nota = montar_linhas_comparacao_por_nota(sped_notas, athena_notas)
            linhas_pis_cofins = montar_linhas_pis_cofins_por_nota(sped_pis_cofins, athena_pis_cofins)
            totais_sped = calcular_totais_cfop(sped_cfop)
            totais_athena = calcular_totais_cfop(athena_cfop)
            totais_por_tipo_sped = calcular_totais_cfop_por_tipo(sped_cfop)
            totais_por_tipo_athena = calcular_totais_cfop_por_tipo(athena_cfop)

            gerar_pdf_comparacao_athena_sped(
                linhas_comparacao,
                totais_sped,
                totais_athena,
                totais_por_tipo_sped,
                totais_por_tipo_athena,
                output_file,
                caminho_txt,
                linhas_por_nota,
                linhas_pis_cofins
            )
            exportar_comparacao_xls(output_xls, linhas_comparacao, totais_sped, totais_athena, totais_por_tipo_sped, totais_por_tipo_athena, linhas_por_nota, linhas_pis_cofins)
            messagebox.showinfo(
                "Sucesso",
                f"Relatorio comparador gerado com sucesso.\n\nPDF: {output_file}\nXLS: {output_xls}"
            )
            os.startfile(output_file)
        except Exception as e:
            messagebox.showerror("Erro ao gerar relatorio", f"Ocorreu um erro ao criar o comparador: {e}")

    def gerar_relatorio_totalizador_contribuicoes(self):
        if not self.data:
            messagebox.showwarning("Aviso", "Nenhum arquivo SPED carregado.")
            return

        output_file = filedialog.asksaveasfilename(
            title="Salvar relatorio comparador SPED Contribuicoes como",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not output_file:
            messagebox.showinfo("Cancelado", "Geracao do relatorio cancelada.")
            return

        totais_sped_info = consolidar_totais_sped_contribuicoes(self.data)
        if not totais_sped_info or not totais_sped_info['cfop_data']:
            messagebox.showwarning("Aviso", "Nenhum registro de faturamento/PIS/COFINS encontrado no arquivo.")
            return

        caminho_txt_padrao = r"C:\Users\Wesley.Raimundo\Desktop\vendas.txt"
        caminho_txt = caminho_txt_padrao if os.path.exists(caminho_txt_padrao) else filedialog.askopenfilename(
            title="Selecione o arquivo Athena para comparacao",
            filetypes=[("Arquivo texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        if not caminho_txt:
            messagebox.showinfo("Cancelado", "Selecao do arquivo Athena cancelada.")
            return

        output_xls = os.path.splitext(output_file)[0] + ".xls"

        try:
            try:
                athena_notas = carregar_athena_contribuicoes_por_nota_txt(caminho_txt)
            except Exception:
                athena_notas = carregar_athena_por_nota_txt(caminho_txt)

            sped_cfop = totais_sped_info['cfop_data']
            sped_notas = carregar_sped_contribuicoes_por_nota(self.data)

            linhas_comparacao = montar_linhas_comparacao_contribuicoes(sped_cfop, athena_cfop)
            linhas_por_nota = montar_linhas_comparacao_por_nota_contribuicoes(sped_notas, athena_notas)

            totais_sped = calcular_totais_cfop(sped_cfop)
            totais_athena = calcular_totais_cfop(athena_cfop)
            totais_por_tipo_sped = calcular_totais_cfop_por_tipo(sped_cfop)
            totais_por_tipo_athena = calcular_totais_cfop_por_tipo(athena_cfop)

            gerar_pdf_comparacao_athena_sped_contribuicoes(
                linhas_comparacao,
                totais_sped,
                totais_athena,
                totais_por_tipo_sped,
                totais_por_tipo_athena,
                output_file,
                caminho_txt,
                linhas_por_nota,
                totais_sped_info.get('totais_bloco_m')
            )
            exportar_comparacao_xls_contribuicoes(
                output_xls,
                linhas_comparacao,
                totais_sped,
                totais_athena,
                totais_por_tipo_sped,
                totais_por_tipo_athena,
                linhas_por_nota,
                totais_sped_info.get('totais_bloco_m')
            )
            messagebox.showinfo(
                "Sucesso",
                f"Relatorio comparador SPED Contribuicoes gerado com sucesso.\n\nPDF: {output_file}\nXLS: {output_xls}"
            )
            os.startfile(output_file)
        except Exception as e:
            messagebox.showerror("Erro ao gerar relatorio", f"Ocorreu um erro ao criar o comparador: {e}")



if __name__ == "__main__":
    root = tk.Tk()
    app = SPEDFileEditor(root)
    root.mainloop()
