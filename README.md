# geapa-desligamentos-suspensoes

Modulo Apps Script responsavel apenas por solicitacoes voluntarias de vinculo no ecossistema GEAPA.

## Escopo atual

- Suspensao voluntaria temporaria.
- Desligamento voluntario imediato.
- Desligamento voluntario ao final do semestre.

Fora de escopo nesta etapa:

- desligamento por faltas;
- suspensao ou desligamento administrativo;
- controle direto de presenca, faltas ou inelegibilidade.

## Contrato institucional obrigatorio

A trilha oficial compartilhada entre modulos para mudancas de vinculo e a key `MEMBER_EVENTOS_VINCULO`, apontando para `PESSOAS > Membros_Eventos`.

Regras praticas:

- este modulo escreve eventos nela;
- o modulo de atividades deve ler dela;
- integracoes em membros atuais/historico nao substituem essa trilha;
- sempre que houver execucao efetiva, o append institucional e obrigatorio.

Helpers oficiais desta camada:

- `off_appendSuspensionEvent_(payload)`
- `off_appendReturnEvent_(payload)`
- `off_appendDismissalEvent_(payload)`

## Arquitetura

Entrada:

- uma unica aba bruta de respostas do formulario, obtida via `OFFBOARD_RESPONSES`.

Filas operacionais:

- `PEDIDOS_SUSPENSAO`, obtida via `OFFBOARD_QUEUE_SUSPENSIONS`
- `PEDIDOS_DESLIGAMENTO`, obtida via `OFFBOARD_QUEUE_DISMISSALS`

Fluxo:

1. `off_handleFormSubmit(e)` le a nova resposta.
2. A submissao e normalizada, recebe `ID_SOLICITACAO` e passa por validacoes leves.
3. A linha e roteada para a fila correta, sem duplicar `ORIGEM_LINHA_FORM`.
4. `off_onEditDecision(e)` reage ao preenchimento de `DECISAO_DIRETORIA`.
5. `off_processDailyLifecycleQueue()` aplica o que estiver vencido no dia.

## Arquivos principais

- [00_off_config.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/00_off_config.js)
- [01_off_registry.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/01_off_registry.js)
- [02_off_form_router.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/02_off_form_router.js)
- [03_off_validation.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/03_off_validation.js)
- [04_off_decision.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/04_off_decision.js)
- [05_off_execution_suspension.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/05_off_execution_suspension.js)
- [06_off_execution_dismissal.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/06_off_execution_dismissal.js)
- [07_off_members_integration.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/07_off_members_integration.js)
- [08_off_lifecycle_events.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/08_off_lifecycle_events.js)
- [09_off_notifications.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/09_off_notifications.js)
- [10_off_jobs.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/10_off_jobs.js)
- [50_off_install.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/50_off_install.js)

## Regras por fluxo

### Suspensao voluntaria

- Exige `DATA_INICIO_SUSPENSAO` e `DATA_FIM_SUSPENSAO`.
- Calcula `QTDE_DIAS_SUSPENSAO`.
- Valida periodo minimo de 30 dias.
- Ao deferir:
  - se a data chegou, aplica na hora;
  - se a data e futura, fica `AGUARDANDO_EXECUCAO`.
- Ao aplicar:
  - registra `SUSPENSAO` com status `HOMOLOGADO`;
  - tenta integracao opcional com `GEAPA_MEMBROS`;
  - nao remove o membro da base atual.
- No fim do periodo:
  - registra `RETORNO` com status `HOMOLOGADO`;
  - tenta integracao opcional com `GEAPA_MEMBROS`.

### Desligamento imediato

- Continua usando `GEAPA_MEMBROS.members_offboardApprovedImmediateExit(payload)`.
- O payload legado foi preservado para esse caso.
- O e-mail final e enviado no momento da execucao.
- O evento `DESLIGAMENTO_VOLUNTARIO` e sempre obrigatorio.

### Desligamento ao fim do semestre

- Ao deferir, calcula ou reaproveita `DATA_EFETIVA_DESLIGAMENTO`.
- Fica `AGUARDANDO_EXECUCAO`.
- O job diario executa quando a data chegar.
- Na execucao, envia e-mail final, integra com membros e registra o evento institucional.

## Integracao futura com atividades

Os contratos/stubs desta fase estao em [03_off_validation.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/03_off_validation.js):

- `off_checkPresentationConflict_(rga, context)`
- `off_checkPendingFiles_(rga, context)`
- `off_checkOtherPendingItems_(rga, context)`
- `off_checkPresentationAndFilePending_(rga, periodoRef, context)`

Comportamento atual:

- tenta ler `PREVISAO_APRESENTACAO_NO_PERIODO` como filtro rapido, se a base existir;
- se necessario, tenta confirmar em `Atividades_Apresentacoes`;
- se a integracao nao estiver disponivel, retorna `NAO_VERIFICADO` sem quebrar o fluxo.

## Compatibilidade Semantica de Ocupacao

Este modulo nao depende diretamente hoje de colunas de `Cargo/Função` para executar seus fluxos principais.
As interacoes reais com a base de membros continuam centradas em identidade (`RGA` e `EMAIL`) e em e-mail group institucional.

Mesmo assim, o modulo passou a expor uma camada centralizada de compatibilidade para a transicao semantica:

- aliases aceitos para leitura: `Ocupação`, `Ocupacao`, `Cargo/Função`, `Cargo/Funcao`
- escrita preferencial em `Ocupação`/`Ocupacao`, com fallback para o legado
- helpers centralizados em [01_off_registry.js](/C:/Users/Windows%2010/geapa-desligamentos-suspensoes/01_off_registry.js):
  - `off_getOccupationHeaderAliases_()`
  - `off_findOccupationColumn_(headerMap)`
  - `off_getOccupationValue_(rowCtx, defaultValue)`
  - `off_setOccupationValue_(sheet, rowNumber, value)`

Com isso, o modulo fica preparado para uma futura etapa de renomeacao fisica dos cabecalhos sem espalhar acesso hardcoded por `Cargo/Função`.

## Triggers

Instalados por `off_installTriggers()`:

- `off_handleFormSubmit` em `onFormSubmit`;
- `off_onEditDecision` em `onEdit`;
- `off_processDailyLifecycleQueue` em trigger diario.

## UX operacional

O modulo agora possui uma camada reaplicavel de UX para as filas operacionais:

- `applyOffboardingSheetUx()`
- `reapplyOffboardingSheetUx()`

Essa camada aplica:

- freeze da linha de cabecalho;
- filtro;
- notas por cabecalho;
- grupos de cor por blocos funcionais;
- dropdowns nas colunas operacionais;
- alinhamento central dos dados;
- compactacao visual nas colunas de texto longo, sem reajustar automaticamente a altura das linhas.

## Migracao

- A aba bruta continua sendo a entrada oficial do Forms.
- Crie as keys de registry:
  - `OFFBOARD_RESPONSES`
  - `OFFBOARD_QUEUE_SUSPENSIONS`
  - `OFFBOARD_QUEUE_DISMISSALS`
- Opcional/recomendada conforme integracoes:
  - `MEMBERS_ATUAIS`
  - `ATIVIDADES_PRESENCAS`
  - `ATIVIDADES_APRESENTACOES`
- As keys das filas devem apontar para as abas operacionais:
  - `OFFBOARD_QUEUE_SUSPENSIONS` -> `PEDIDOS_SUSPENSAO`
  - `OFFBOARD_QUEUE_DISMISSALS` -> `PEDIDOS_DESLIGAMENTO`
- Se a key existir mas a aba ainda nao existir fisicamente, o modulo cria a aba com o `sheetName` configurado e injeta os cabecalhos.
- Rode `off_installTriggers()` apos publicar esta versao.
- Se houver pedidos antigos apenas na aba bruta, eles nao sao retro-roteados automaticamente nesta etapa.

## Teste manual rapido

1. Envie uma resposta de suspensao e confirme que a linha vai para `PEDIDOS_SUSPENSAO`.
2. Marque `DECISAO_DIRETORIA = DEFERIDO` numa suspensao com inicio no dia atual e confirme append de `SUSPENSAO`.
3. Ajuste a data final para hoje e rode `off_processDailyLifecycleQueue()` para confirmar append de `RETORNO`.
4. Envie um desligamento imediato, defera a linha e confirme e-mail final, integracao com membros e evento `DESLIGAMENTO_VOLUNTARIO`.
5. Envie um desligamento de fim de semestre, defera a linha e confirme `AGUARDANDO_EXECUCAO`; depois rode o job com a data efetiva vencida.

## Riscos e TODOs conhecidos

- A deteccao de campos do formulario depende de aliases; se o Forms mudar muito de texto, sera preciso ampliar a lista em `OFF_CFG.RAW_FIELDS`.
- A integracao com atividades esta isolada e ainda parcial; `NAO_VERIFICADO` e um estado esperado nesta versao.
- O backfill de pedidos antigos da aba bruta ficou como etapa futura.
