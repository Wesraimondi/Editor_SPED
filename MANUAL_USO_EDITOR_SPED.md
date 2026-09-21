# Manual Completo de Uso - Editor SPED Fiscal & Contribuições

Este manual contém a documentação completa, passo a passo e explicativa de todas as funcionalidades, atalhos, módulos de auditoria, tratamentos por bloco e inteligência fiscal do **Editor de Arquivos SPED**.

---

## 📋 Sumário
1. [Visão Geral & Carregamento de Arquivos](#1-visão-geral--carregamento-de-arquivos)
2. [Navegação, Zoom e Títulos Dinâmicos da Tabela](#2-navegação-zoom-e-títulos-dinâmicos-da-tabela)
3. [Filtros de Pesquisa Combinada (Registro + Texto)](#3-filtros-de-pesquisa-combinada-registro--texto)
4. [Módulos de Auditoria e Pré-Validação Fiscal](#4-módulos-de-auditoria-e-pré-validação-fiscal)
5. [Tratamento e Correções por Bloco (0, C, D, H, 9)](#5-tratamento-e-correções-por-bloco-0-c-d-h-9)
6. [Inteligência Fiscal & Dashboards](#6-inteligência-fiscal--dashboards)
7. [Ferramentas Avançadas (Excel, Leiautes e Macros)](#7-ferramentas-avançadas-excel-leiautes-e-macros)
8. [Tabela Completa de Atalhos de Teclado](#8-tabela-completa-de-atalhos-de-teclado)

---

## 1. Visão Geral & Carregamento de Arquivos

O **Editor SPED** foi desenvolvido para tratar arquivos textuais da **Escrituração Fiscal Digital (EFD ICMS/IPI)** e **EFD Contribuições (PIS/COFINS)**.

- **Abertura Assíncrona (`Ctrl+O`)**: Arquivos SPED extensos (com dezenas de megabytes ou centenas de milhares de linhas) são lidos em segundo plano (Threading), acompanhados por uma **Barra de Progresso** no painel de status sem travar a interface.
- **Indicador de Alteração (`*`)**: Quando uma modificação for realizada no arquivo, um asterisco (`*`) aparecerá no título da janela. Ao fechar sem salvar, o sistema exibirá uma caixa de confirmação.

---

## 2. Navegação, Zoom e Títulos Dinâmicos da Tabela

- **Mapeamento Automático de Cabeçalhos (270 Registros Oficial v3.2.2)**:
  - Ao clicar em qualquer linha da tabela (`0000`, `0150`, `0200`, `C100`, `C170`, `D100`, `E110`, `H010`, `K200`, `M200`, `1601`, etc.), o sistema reconhece o código do registro (suportando zeros à esquerda) e substitui os nomes genéricos das colunas (`C1`, `C2`, `C3`...) pelas **nomenclaturas oficiais dos campos** (`REG`, `IND_OPER`, `COD_PART`, `VL_DOC`, `CST_PIS`, `VL_BC_PIS`, etc.).
- **Editor de Campos da Linha (Duplo-Clique)**:
  - Ao dar duplo-clique em qualquer linha da tabela principal, abre-se uma janela pop-up exibindo cada campo com seu número, nome oficial e valor atual para alteração rápida.
- **Controle de Zoom da Tabela**:
  - `Ctrl + Plus` (`Ctrl + =`): Aumenta o tamanho da fonte e a altura proporcional das linhas (`rowheight`).
  - `Ctrl + Minus` (`Ctrl + -`): Diminui o tamanho da fonte e a altura das linhas.
  - `Ctrl + 0`: Redefine a fonte para o padrão inicial (`10pt`).

---

## 3. Filtros de Pesquisa Combinada (Registro + Texto)

No painel superior da aplicação você encontrará o sistema de **Filtro Combinado Simultâneo**:
1. **Filtro Registro**: Selecione no Combobox o registro desejado (ex: `C100`, `C170`, `0200`, `H010`).
2. **Filtro Texto**: Digite no campo de texto qualquer valor (ex: NCM `38151210`, CFOP `5102` ou Razão Social).
3. A tabela filtrará apenas as linhas que atendem a ambos os critérios em tempo real.

---

## 4. Módulos de Auditoria e Pré-Validação Fiscal

Acesse pelo menu `Ferramentas`:

1. **Auditoria de Campos Obrigatórios em Branco**:
   - Varre o arquivo localizando descrições nulas ou códigos ausentes em registros cruciais (ex: `0200` Campo 3 - Descrição, `0150` Campo 3 - Nome, `0000` Razão Social).
   - O relatório **permanece aberto** ao dar duplo-clique em um erro, abrindo a janela de edição do campo e permitindo corrigir vários erros em sequência.
2. **Módulo de Pré-Validação Fiscal (PVA)**:
   - Valida com algoritmo Módulo 11 o Dígito Verificador (DV) de Chaves de Acesso de NFe/CTe (44 dígitos) nos registros `C100`, `D100` e `C800`.
   - Valida os dígitos verificadores de CNPJ (14 dígitos) e CPF (11 dígitos) nos registros `0000` e `0150`.
3. **Auditoria de Matriz Tributária (CFOP x CST)**:
   - Cruza produtos tributados com CFOPs de isenção ou ST.
   - Possui botão **"Corrigir Inconsistências em Lote (1-Clique)"** para alterar o CST para `060` automaticamente.
4. **Verificação de Registros Duplicados**:
   - Detecta e remove registros idênticos repetidos (ex: `0600`).

---

## 5. Tratamento e Correções por Bloco (0, C, D, H, 9)

- **Bloco 0**:
  - `0200 - Corrigir NCM Vazio / Zerado '00' (Campo 8)`: Varre todos os registros `0200` e localiza itens onde a NCM esteja em branco ou preenchida com `'00'`/zeros. Abre janela interativa com tabela de produtos, permitindo visualizar o item, consultar a tabela TIPI em tempo real, informar o valor correto individualmente (duplo-clique), usar o assistente passo a passo ou preencher em lote.
  - `0200 - Preencher / Atualizar CEST`: Atualiza o CEST (campo 13) a partir da tabela oficial `CEST12.xlsx` (742 correspondências NCM x CEST).
  - `0150 - Município / IE MG`: Formata Inscrições Estaduais de Minas Gerais e preenche o código de município IBGE.
- **Bloco C**:
  - `C100 - Preencher Série Padrão (000)`: Preenche séries nulas.
  - `C170 - Vincular Plano de Contas (Campo 38)`: Vincula automaticamente o código da conta contábil.
  - `C170 - Alterar CST PIS/COFINS em Lote`: Atualização massiva de CSTs.
  - `C190 - Gerar Ajustes C195 / C197`: Cria registros de ajuste automaticamente abaixo do `C190`.
- **Bloco D**: Recálculo de PIS e COFINS nos registros `D101`/`D105` e conta contábil de frete no `D100`.
- **Bloco H**: Sincronização automática das unidades de medida do Inventário `H010` com os produtos do `0200`.
- **Bloco 9 (Recálculo Automático)**:
  - Ao salvar o arquivo (`Ctrl+S`), o sistema recalcula e gera a estrutura completa dos registros `9900` (contadores por tipo de registro), `9990` (total do Bloco 9) e `9999` (total de linhas do arquivo).

---

## 6. Inteligência Fiscal & Dashboards

- **Dashboard Executivo de Apuração**:
  - Menu `Ferramentas -> Inteligência - Dashboard Executivo de Apuração`.
  - Exibe cards com o total de **Vendas (Saídas)** vs **Compras (Entradas)**, impostos destacados (ICMS, IPI, PIS, COFINS) e o resumo de Apuração de ICMS do Bloco `E110` (débitos, créditos e saldo credor/devedor).
- **Comparador de Dois Arquivos SPED**:
  - Menu `Ferramentas -> Auditoria - Comparador de Dois Arquivos SPED`.
  - Cruza o SPED gerado pelo ERP com o SPED transmitido/validado, apontando notas faltantes e divergências de valores.

---

## 7. Ferramentas Avançadas (Excel, Leiautes e Macros)

- **Exportador para Excel (`.xlsx`)**:
  - Menu `Ferramentas -> Exportar Registro Selecionado para Excel (.xlsx)`.
  - Gera uma planilha Excel formatada com o nome oficial dos campos no cabeçalho.
- **Gerenciador de Leiautes SPED**:
  - Menu `Cadastros -> Gerenciar Leiautes de Registros SPED`.
  - Permite visualizar, pesquisar e cadastrar novos registros SPED com salvamento local em `sped_layouts_custom.json`.
- **Gerenciador de Macros Personalizadas**:
  - Menu `Ferramentas -> Inteligência - Gerenciador de Macros Personalizadas`.
  - Permite criar e salvar regras de edição recorrentes em `sped_rules_custom.json` para aplicação em 1-clique.

---

## 8. Tabela Completa de Atalhos de Teclado

| Atalho | Ação Executada |
| :--- | :--- |
| **F1** | Abre o Manual Completo de Uso Interativo |
| **Ctrl + O** | Abrir Arquivo SPED... |
| **Ctrl + S** | Salvar Alterações (Recalcula o Bloco 9) |
| **Ctrl + Shift + S** | Salvar Como... |
| **Ctrl + Z** | Desfazer última alteração de edição |
| **Ctrl + F** | Focar no campo de Busca de Texto |
| **F3** | Localizar próxima ocorrência de texto |
| **Ctrl + Plus / =** | Aumentar Tamanho da Fonte e Altura das Linhas da Tabela |
| **Ctrl + Minus / -** | Diminuir Tamanho da Fonte e Altura das Linhas da Tabela |
| **Ctrl + 0** | Redefinir Tamanho da Fonte para o Padrão (`10pt`) |
| **Duplo-Clique (Tabela)** | Abre o Editor de Linha Completo |
| **Duplo-Clique (Auditorias)**| Seleciona a linha e abre o Editor de Campos mantendo o relatório aberto |
