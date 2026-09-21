# Editor_SPED

# 🏛️ Sistema Wesley / ACHILES Fiscal & Contábil Suite

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)
![GUI](https://img.shields.io/badge/Interface-Tkinter%20%7C%20CustomTkinter-brightgreen?style=for-the-badge)
![Database](https://img.shields.io/badge/Database-SQLite3-lightgrey?style=for-the-badge&logo=sqlite)
![Status](https://img.shields.io/badge/Status-Em%20Produção-success?style=for-the-badge)

Uma suíte completa e integrada de ferramentas fiscais, contábeis e de automação tributária voltada para auditoria de arquivos SPED (Fiscal e Contribuições), emissão e gestão de Guias de Recolhimento (GNRE, DUA, DARE), conciliação bancária/contábil e monitoramento de documentos fiscais eletrônicos (NF-e, NFS-e, CT-e).

---

## 📌 Sumário

- [Visão Geral dos Módulos](#-visão-geral-dos-módulos)
  - [1. Módulo Fiscal & Editor SPED](#1-módulo-fiscal--editor-sped)
  - [2. Módulo de Guias & GNRE Nacional](#2-módulo-de-guias--gnre-nacional)
  - [3. Módulo Contábil & Lotes Bancos](#3-módulo-contábil--lotes-bancos)
  - [4. Radar Fiscal & Documentos Eletrônicos](#4-radar-fiscal--documentos-eletrônicos)
  - [5. Menu Centralizador de Aplicações](#5-menu-centralizador-de-aplicações)
- [🛠️ Tecnologias Utilizadas](#️-tecnologias-utilizadas)
- [📁 Estrutura de Diretórios](#-estrutura-de-diretórios)
- [🚀 Instalação e Configuração](#-instalação-e-configuração)
- [💻 Como Executar](#-como-executar)
- [🛡️ Diretrizes de Desenvolvimento e Backup](#️-diretrizes-de-desenvolvimento-e-backup)
- [📄 Licença](#-licença)

---

## 🧭 Visão Geral dos Módulos

### 1. Módulo Fiscal & Editor SPED
Localizado em `01 - FISCAL/Sped/Editor Sped.py`.
- **Auditoria e Edição SPED:** Manipulação avançada de registros dos blocos `0`, `C`, `D`, `E`, `F`, `H` e `9` de arquivos EFD ICMS/IPI e EFD Contribuições.
- **Relatório Comparador Athena x SPED:**
  - Confronto detalhado entre faturamento Athena (`vendas.txt`) e SPED Fiscal por CFOP e por Nota Fiscal.
  - Rateio inteligente de PIS/COFINS por CFOP priorizando operações tributadas (eliminando divergências em notas de industrialização e retornos como 5124/5902).
  - Exportação completa em relatórios paginados em PDF (ReportLab) e planilhas formatadas em Excel (.XLS).
- **Validação de Inventário (H010 vs 0200):**
  - Checagem automática entre os itens do inventário físico (`H010`) e a tabela de cadastro de itens (`0200`).
  - Identificação de itens faltantes com interface visual e importação direta a partir de arquivos TXT/SPED de referência.
- **Matriz de Análise Tributária:** Visualização e auditoria em grade dos registros C170, CFOPs, alíquotas e CSTs de PIS/COFINS.
- **Recálculo Automático de Blocos:** Recálculo em lote dos totalizadores de encerramento (`0990`) e de todo o Bloco 9 (`9001`, `9900`, `9990`, `9999`).

### 2. Módulo de Guias & GNRE Nacional
Localizado em `01 - FISCAL/Guia Nacional/` e `GNRE_System/`.
- **Emissão Automatizada de Guias:** Geração em lote de guias GNRE (Portal Nacional e SEFAZ-PE), DUA (Espírito Santo) e DARE.
- **Integração com Web Services SEFAZ:** Consulta de status, envio de lotes de guias e download automático de comprovantes/PDFs.
- **Histórico e Repositório Local:** Armazenamento em banco SQLite (`gnre_nacional_local.db` e `DADOS_GNRE.db`) com rastreamento completo de guias emitidas e pagas.

### 3. Módulo Contábil & Lotes Bancos
Localizado em `01 - FISCAL/` e `Contabilidade/`.
- **Conciliação e Classificação de Extratos:** Importação de extratos e lotes contábeis para classificação de partidas dobradas (Débito / Crédito).
- **Motor de Sugestão e Histórico Inteligente:** Base de conhecimento (`historico_classificacoes.db`) e dicionários JSON (`codigos_contabeis.json`, `codigos_sugestao.json`) para classificação automática de lançamentos bancários recorrentes.
- **Integração Domínio Sistemas:** Exportação e conversão de arquivos formatados para importação no software Domínio Contábil.

### 4. Radar Fiscal & Documentos Eletrônicos
Localizado em `Radar_Fiscal/`, `Obs NFe/` e `XML/`.
- **Monitoramento e Processamento de XMLs:** Leitura de NF-e, NFS-e e CT-e, extração de dados fiscais e cálculo de impostos retidos.
- **Gestão de CIAP e Variação Cambial:** Controle de créditos de ativo permanente e apuração de variações cambiais.

### 5. Menu Centralizador de Aplicações
Localizado em `sistema_menus.py` / `main_menu.py`.
- **Interface Launcher:** Painel gráfico central com categorização de ferramentas, histórico de acessos, busca dinâmica, favoritos e personalização de temas.

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** [Python 3.10+](https://www.python.org/)
- **Interface Gráfica:** `Tkinter`, `ttk`, `CustomTkinter`
- **Manipulação de Dados:** `pandas`, `openpyxl`, `xlrd`, `lxml`
- **Geração de Relatórios:** `ReportLab` (PDFs estilizados com tabelas e cabeçalhos corporativos)
- **Banco de Dados:** `SQLite3`
- **Web & Comunicação:** `requests`, `urllib3`, `zeep` (SOAP/WSDL SEFAZ)
- **Empacotamento:** `PyInstaller` (geração de executáveis standalone `.exe`)

---

## 📁 Estrutura de Diretórios

```text
├── 01 - FISCAL/               # Módulos fiscais, SPED, GNRE e Lotes Bancos
│   ├── Sped/                  # Editor SPED, Comparador Athena e relatórios
│   ├── Guia Nacional/         # Emissão de GNRE e comunicação SEFAZ
│   └── ...                    # Lançamentos contábeis e bancos
├── 03 - INTEGRAÇÕES/          # Conectores e importadores externos
├── 04 - CONTÁBIL/             # Rotinas e conciliações contábeis
├── 05 - SISTEMAS E MODULOS/   # Módulos auxiliares de negócio
├── 05 - UTILITÁRIOS/          # Scripts e utilitários de suporte
├── Contabilidade/             # Sistema de plano de contas e conciliações
├── GNRE_System/               # Motores legados e auxiliares de GNRE
├── Menu_Sistema/              # Recursos gráficos do launcher
├── Radar_Fiscal/              # Monitoramento de NFe e tributos
├── scripts/                   # Scripts de backup e automação
├── sistema_menus.py           # Launcher principal do sistema
├── requirements.txt           # Dependências do projeto
├── AGENTS.md                  # Regras de manutenção e backup
└── README.md                  # Documentação do repositório
