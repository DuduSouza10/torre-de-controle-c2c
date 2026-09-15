# Torre de Controle C2C - Railway

Painel preparado para rodar no Railway com **base compartilhada entre todos os usuários**.

## O que mudou nesta versão

- Filtro por **dias sem movimentação desde o último bipe operacional**:
  - informe apenas o mínimo para ver, por exemplo, todos com `5+ dias`;
  - informe mínimo e máximo para uma faixa, por exemplo, `3 até 7 dias`;
  - quando esse filtro é usado, entram apenas pedidos ainda não entregues/devolvidos e com data válida da última operação.
- O upload de XLSX/CSV agora atualiza uma **base central no servidor**.
- Pedidos repetidos são atualizados pelo **número do pedido**.
- Quem abrir o site depois já recebe a versão mais recente da base.
- Quem já estiver com o painel aberto verifica automaticamente se existe uma nova versão a cada 30 segundos e também ao voltar para a aba.
- Botão **Sincronizar agora** para forçar uma atualização imediata.

## Persistência obrigatória no Railway

Para os uploads continuarem existindo depois de restart/redeploy, adicione um **Railway Volume** ao mesmo serviço.

### Recomendado

1. Faça o deploy do projeto.
2. No serviço do Railway, adicione um **Volume**.
3. Defina o Mount Path como:

```text
/data
```

O Railway fornece `RAILWAY_VOLUME_MOUNT_PATH` automaticamente e o `app.py` usa esse caminho. Se nenhum volume estiver conectado, o projeto ainda funciona, mas os dados gravados no filesystem podem ser perdidos em um novo deploy/restart.

Você também pode definir manualmente a variável `DATA_DIR` se quiser usar outro diretório.

## Deploy

1. Extraia o ZIP.
2. Envie a pasta `torre_controle_c2c_railway` para um repositório GitHub.
3. No Railway, crie um serviço a partir do repositório.
4. Adicione o Volume conforme a seção acima.
5. Em **Networking**, gere o domínio público do serviço.

O projeto usa:

```text
python app.py
```

O Railway fornece a variável `PORT` automaticamente.

## Endpoints

- `/` - painel
- `/health` - health check
- `/api/meta` - versão atual da base compartilhada
- `/api/data` - base compartilhada atual em CSV
- `/api/upload` - recebe as linhas normalizadas do novo upload e faz merge no servidor

## Rodar localmente

```bash
python app.py
```

Abra:

```text
http://localhost:8080
```

Localmente, se `DATA_DIR` não for definido, os dados são gravados em `data/runtime/`.

## Observação sobre concorrência

O servidor usa lock durante a atualização e gravação atômica do CSV/metadata, evitando que dois uploads simultâneos corrompam a base.
