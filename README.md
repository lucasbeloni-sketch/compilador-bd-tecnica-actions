# compilador-bd-tecnica-actions

Le todos os `.xlsx` de uma pasta do Google Drive (Shared Drive), normaliza o cabecalho
conforme a referencia da planilha destino, junta tudo, **deduplica por TRABALHO mantendo
o arquivo mais recente** e cola na aba `BD_Tecnica`.

## O que faz (passo a passo)

1. Le o cabecalho de referencia em `BD_Tecnica!D5:DK5` do destino (nomes + ordem alvo, 112 col).
2. Lista os `.xlsx` da pasta do Drive (`driveFolderId`) — funciona em Shared Drive.
3. Pra cada arquivo:
   - extrai a data do **nome** (`2026.06.25 - BASE SIRTEC_TECNICA` -> 25/06/2026);
   - le a 1a planilha (cabecalho na linha 1, `A1:DH1`);
   - **reordena** as colunas pra ordem da referencia (por nome do cabecalho);
   - colunas extras do arquivo sao ignoradas;
   - se faltar coluna da referencia (cabecalho nao bate) ou nao tiver data no nome -> **pula e loga**.
4. Junta tudo e deduplica por `TRABALHO`: pra cada TRABALHO repetido, **prevalece o arquivo
   com data mais recente** (mesmo que as outras colunas sejam diferentes).
5. Limpa `D6:DK` e cola o resultado a partir de `D6`.

Guarda anti-apagao: se compilar menos que `minRows`, aborta **sem** limpar o destino.

## Configuracao

Tudo em [`config.json`](./config.json). IDs ja preenchidos:

| campo | valor |
|---|---|
| `driveFolderId` | `1FhiHfhVK5lH7LAzbtYmMm85FKWBNoxLt` |
| `destSpreadsheetId` | `1YtcYEFgrxVW59xG-H57gV25hWpZtROwnabZgVNrPnz0` |
| `destSheet` | `BD_Tecnica` |
| `headerRefRange` | `D5:DK5` |
| `destStartCell` / `destClearRange` | `D6` / `D6:DK` |
| `keyColumn` | `TRABALHO` |

## Pre-requisitos (Service Account)

Reusa a **mesma Service Account** dos outros robos (env `GOOGLE_CREDENTIALS`).
Compartilhe com o email da SA:

- a **pasta origem** no Drive (acesso de leitura) — ela esta num Shared Drive, entao
  adicione a SA como membro do Shared Drive ou compartilhe a pasta;
- a **planilha destino** (acesso de editor).

Scopes usados: `drive.readonly` + `spreadsheets`.

## Rodar local

```bash
npm install

# coloque a chave da SA num arquivo e exporte (PowerShell):
#   $env:GOOGLE_CREDENTIALS = Get-Content credentials.json -Raw

# teste sem escrever no destino:
DRY_RUN=1 npm start      # bash
#   $env:DRY_RUN="1"; npm start    # PowerShell

# rodar de verdade:
npm start

# testes unitarios:
npm test
```

## GitHub Actions

`.github/workflows/compilar.yml` roda de hora em hora (+ botao manual). Configure o secret
`GOOGLE_CREDENTIALS` (JSON da SA) em Settings > Secrets and variables > Actions.
Em falha abre/comenta uma issue; ao voltar a funcionar, fecha.
