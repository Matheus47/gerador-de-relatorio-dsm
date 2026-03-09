# 📊 Relatório Leads / MQL — App Streamlit

App web que substitui o Google Colab para gerar relatórios de Leads/MQL a partir de uma base Excel (RD Station / CRM).

---

## 📁 Arquivos do projeto

```
leads_app/
├── app.py                    ← Interface web (Streamlit)
├── relatorio_leads_mql.py    ← Lógica de análise (não mexa aqui)
├── requirements.txt          ← Dependências Python
└── README.md                 ← Este arquivo
```

---

## 🚀 Como publicar no Streamlit Community Cloud (GRÁTIS)

### Passo 1 — Criar conta no GitHub
Acesse https://github.com e crie uma conta gratuita se ainda não tiver.

### Passo 2 — Criar um repositório
1. Clique em **"New repository"**
2. Dê um nome, ex: `relatorio-leads-mql`
3. Deixe como **Public**
4. Clique em **"Create repository"**

### Passo 3 — Fazer upload dos arquivos
1. Dentro do repositório, clique em **"uploading an existing file"**
2. Arraste os 3 arquivos: `app.py`, `relatorio_leads_mql.py`, `requirements.txt`
3. Clique em **"Commit changes"**

### Passo 4 — Deploy no Streamlit Cloud
1. Acesse https://share.streamlit.io
2. Faça login com sua conta GitHub
3. Clique em **"New app"**
4. Selecione o repositório `relatorio-leads-mql`
5. No campo **"Main file path"**, digite: `app.py`
6. Clique em **"Deploy!"**

Aguarde ~2 minutos e o link do app estará disponível. 🎉

---

## 💻 Como rodar localmente (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## 🔒 Privacidade

Os arquivos carregados no app **não são salvos em nenhum servidor** — são processados em memória e descartados após o download.

---

## ✏️ Personalizações futuras

- Alterar as tags em `relatorio_leads_mql.py` → variável `OFFICIAL_TAGS`
- Mudar cores/layout → arquivo `app.py`
