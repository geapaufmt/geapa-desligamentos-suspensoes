/***************************************
 * 00_off_config.gs
 * MÓDULO: Desligamento & Suspensão (OFFBOARD)
 *
 * Este arquivo só guarda configuração.
 * Nada de abrir planilha nem enviar e-mail aqui.
 ***************************************/

const OFF_KEYS = Object.freeze({
  RESPONSES: 'OFFBOARD_RESPONSES',   // planilha de respostas do Forms
  MEMBERS:   'MEMBERS_ATUAIS',       // base de membros (para achar secretários)
});

const OFF_CFG = Object.freeze({
  TZ: 'America/Cuiaba',

  // Cabeçalhos na planilha de respostas do Forms
  RESP: Object.freeze({
    HEADER_ROW: 1,

    // cabeçalhos (devem bater com a planilha)
    TYPE: 'Quero solicitar:',
    NAME: 'Nome Completo',
    RGA: 'RGA',
    EMAIL: 'Endereço de e-mail',

    // controle do fluxo
    NOTIFIED_SECS: 'Notificação aos secretários enviada?',
    APPROVED: 'Decisão da Diretoria',
    SENT_FINAL: 'E-mail de desligamento enviado?',
    DEFER_DATE: 'Data/Hora do deferimento',
  }),

  // Valores esperados / regras
  VALUES: Object.freeze({
    YES: 'SIM',
    DECISAO_DEFERIDO: 'DEFERIDO',
    TYPE_OFFBOARD: 'Desligamento',
    TYPE_SUSPEND: 'Suspensão',
  }),

  // Como identificar secretários na base de membros
  MEMBERS: Object.freeze({
    COL_ROLE: 'Cargo/função atual',
    COL_EMAIL: 'EMAIL',
    SECRETARY_ROLES: [
      'Secretário(a) Executivo(a)',
      'Secretário(a) Geral',
    ],
  }),

  EMAIL: Object.freeze({
    FROM_NAME: 'GEAPA',

    // assunto de notificação aos secretários (prefixo + tipo)
    NOTIFY_SECS_SUBJECT_PREFIX: 'GEAPA — Pedido aguardando análise: ',

    // assunto do e-mail final ao membro
    FINAL_SUBJECT_OFFBOARD: 'Confirmação de desligamento – GEAPA',
    FINAL_SUBJECT_SUSPEND:  'Confirmação de suspensão – GEAPA',
  }),
});