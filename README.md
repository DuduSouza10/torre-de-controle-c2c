# Torre de Controle C2C - Railway

Projeto adaptado para hospedagem no Railway sem dependências externas de Python.

## Deploy no Railway

1. Extraia o ZIP.
2. Envie a pasta para um repositório GitHub ou use o fluxo de deploy do Railway.
3. Crie um novo projeto no Railway e selecione o repositório/pasta.
4. O Railway executará `python app.py` e fornecerá a variável `PORT` automaticamente.
5. O health check está disponível em `/health`.

O projeto também inclui `Dockerfile`, então pode ser implantado pelo builder Docker caso prefira.

## Persistência dos dados

- O painel continua lendo XLSX/CSV diretamente no navegador.
- Os arquivos carregados pelo usuário não são enviados ao servidor Railway.
- A base processada fica salva no navegador usando IndexedDB.
- Cada navegador/dispositivo mantém sua própria base local.

## Rodar localmente

```bash
python app.py
```

Abra `http://localhost:8080`.
