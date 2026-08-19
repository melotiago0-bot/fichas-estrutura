# Fichas de Estrutura

App Flask que gera **fichas de estrutura** de UCs (Universidade Europeia) em Word, a partir de um formulário web. Preenche um template `.docx` preservando 100% a formatação, editando o XML cru por número de linha.

## Stack

- Python + Flask (`app:app`), servido por gunicorn
- Manipulação Word: edição direta do `word/document.xml` (sem python-docx) + reempacotamento via `scripts/pack.py`
- `defusedxml` + `lxml` nas deps (parsing seguro)
- Deploy: Railway via Dockerfile (`python:3.12-slim`)

## Ficheiros principais

- `app.py` — Flask; rotas `/` (form), `/generate` (POST → .docx), `/sugestao` (POST → Google Sheets). Mapas `HEADER` e `UE_MAP` = números de linha no `document.xml` a substituir.
- `static/index.html` — formulário (frontend)
- `template_clean.docx` — template Word original (fonte da verdade da formatação)
- `clean_unpacked/` — o template já descompactado (o zip .docx aberto); a geração copia esta pasta, substitui o texto e reempacota
- `scripts/pack.py` / `scripts/unpack.py` — (des)empacotam o .docx; `scripts/validators/` e `scripts/schemas/` validam OOXML
- `requirements.txt`, `Dockerfile`, `.dockerignore`

## Como funciona a geração

1. `/generate` recebe JSON `{nomeUC, autor, apresentacao, palavrasChave[], ues:[{numero,titulo,descricao,eatividades:[{titulo,tipo}]}]}`
2. `fill_and_pack()` copia `clean_unpacked/` para tmp, substitui texto por número de linha (`HEADER`/`UE_MAP`), reempacota com `pack.py --original template_clean.docx`
3. Devolve o `.docx` como download (`<nomeUC>_ficha_estrutura.docx`)

**FRÁGIL:** `UE_MAP`/`HEADER` são números de linha fixos no `document.xml`. Reformatar o template desalinha tudo — reeditar o template implica remapear as linhas.

## Deployment

```bash
cd ~/Documents/claude/fichas-estrutura
git add <ficheiros>
git commit -m "descrição"
railway up --detach
```

Repo GitHub: `melotiago0-bot/fichas-estrutura`. No hub dentro do grupo **UE**.
