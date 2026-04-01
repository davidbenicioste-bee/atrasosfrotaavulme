# 🔧 Extrator M.EMB — Manutenção Embarcada

Ferramenta web (Streamlit) para extrair automaticamente todos os registros de
**M.EMB. (Manutenção Embarcada)** do Mapa de Atrasos, com datas, quantidades,
anotações e legenda.

---

## ✅ Requisitos

- Python 3.8 ou superior instalado
- Conexão com internet (apenas na primeira instalação)

---

## 🚀 Passo a Passo

### 1. Instale as dependências
Abra o **Prompt de Comando** (Windows) ou **Terminal** (Mac/Linux) na pasta da ferramenta e execute:

```
pip install -r requirements.txt
```

> ⏳ Aguarde — o Streamlit será instalado automaticamente (pode levar alguns minutos na primeira vez).

---

### 2. Inicie a ferramenta

```
streamlit run app.py
```

O navegador abrirá automaticamente em **http://localhost:8501**

---

### 3. Use a ferramenta

1. **Faça o upload** — arraste o arquivo `.xlsm` ou `.xlsx` ou clique para selecionar
2. **Aguarde o processamento** — a leitura é automática
3. **Veja os resultados** — métricas, tabela com filtros e legenda na tela
4. **Baixe o Excel** — clique em "⬇ Baixar Excel (.xlsx)"

---

## 📋 O que é extraído

| Campo         | Descrição                                              |
|---------------|--------------------------------------------------------|
| **Núcleo**    | Aba da planilha (ex: Núcleo 1, Núcleo 4...)            |
| **Linha**     | Código da linha de ônibus                              |
| **Período**   | 1º Período (manhã) ou 2º Período (tarde)               |
| **Data**      | Data do atraso (ex: 02/03/2026)                        |
| **Qtd Atrasos** | Quantidade registrada na célula                      |
| **Anotação**  | Comentário da célula (nº do veículo + motivo)          |

O Excel gerado contém **2 abas**:
- `M.EMB - Manut. Embarcada` → todos os registros com atrasos
- `Legenda` → itens de legenda (MAN - MANUTENÇÃO / MAN - EMBARCADA)

---

## 📁 Arquivos

```
memb_streamlit/
├── app.py            ← aplicação principal
├── requirements.txt  ← dependências Python
└── README.md         ← este arquivo
```

---

## ⚠️ Observações

- A planilha deve seguir o formato padrão do **Mapa de Atrasos** com abas de Núcleo
- O **Núcleo 2** normalmente não possui a categoria M.EMB. (estrutura diferente)
- Funciona com `.xlsm` e `.xlsx`
- Para encerrar a ferramenta: pressione `Ctrl + C` no terminal
