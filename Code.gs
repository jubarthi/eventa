// ============================================================
//  EVENTA — Backend Google Apps Script
//  Versão 1.1 | Com Autenticação
// ============================================================

const CONFIG = {
  FOLDER_NAME:     'MARCELO_e_RAFA',
  FOLDER_ID:       '1ulRjtioJRIcqCmfjkbcG779jQ_LvZsVz',
  SHEET_REGISTROS: 'Registros',
  SHEET_CONFIG:    'Configuracao',
  MAX_FILE_MB:     20,
  SESSION_HORAS:   24,
  RESET_MINUTOS:   60
};

// Credenciais padrão dos administradores
const ADMIN_USERS = [
  {
    email: 'othiagoschwanz@gmail.com',
    senha: '@#TDSe2026*',
    nome:  'Thiago'
  },
  {
    email: 'avelarmarcelo@gmail.com',
    senha: 'pheijao0609',
    nome:  'Marcelo'
  }
];

const ADMIN_DEFAULT = ADMIN_USERS[0];

// ─────────────────────────────────────────
//  ROTEADOR GET
// ─────────────────────────────────────────

function doGet(e) {
  inicializarAdmin();
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  try {
    switch (action) {
      // Públicas (convidado)
      case 'getEvent': return respond(getEvent(params.code));
      case 'ping':     return respond({ ok: true, ts: new Date().toISOString() });

      // Protegidas (painel)
      case 'getStats':  return respondAuth(e || { parameter: {} }, () => getStats());
      case 'getConfig': return respondAuth(e || { parameter: {} }, () => getConfig());
      case 'getLinks':  return respondAuth(e || { parameter: {} }, () => getLinks());

      default: return respond({ ok: true, msg: 'EVENTA API Online' });
    }
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ─────────────────────────────────────────
//  ROTEADOR POST
// ─────────────────────────────────────────

function doPost(e) {
  inicializarAdmin();
  try {
    const data   = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const action = data.action || '';
    switch (action) {
      // Públicas (convidado)
      case 'uploadFile':     return respond(uploadFile(data));
      case 'finalizarEnvio': return respond(finalizarEnvio(data));

      // Auth
      case 'login':          return respond(login(data));
      case 'logout':         return respond(logout(data));
      case 'esqueceuSenha':  return respond(esqueceuSenha(data));
      case 'redefinirSenha': return respond(redefinirSenha(data));

      // Protegidas (painel)
      case 'saveConfig':  return respondAuthPost(data, () => saveConfig(data));
      case 'saveLink':    return respondAuthPost(data, () => saveLink(data));
      case 'deleteLink':  return respondAuthPost(data, () => deleteLink(data));

      default: return respond({ error: 'Ação inválida' });
    }
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Verifica token e executa ação (GET)
function respondAuth(e, fn) {
  const token = e.parameter.token || '';
  if (!verificarSessao(token)) return respond({ error: 'Sessão inválida. Faça login novamente.' });
  return respond(fn());
}

// Verifica token e executa ação (POST)
function respondAuthPost(data, fn) {
  if (!verificarSessao(data.token || '')) return respond({ error: 'Sessão inválida. Faça login novamente.' });
  return respond(fn());
}

const EVENTO_DEFAULT = {
  titulo:            'Casamento Rafa & Marcelo',
  subtitulo:         'Celebrando o Amor 💒',
  data_evento:       '2026-09-06',
  codigo_evento:     'RAFMAR06',
  mensagem:          'Sejam muito bem-vindos ao nosso casamento! Compartilhe conosco todas as fotos e vídeos que você tirar hoje para guardarmos para sempre. 💛',
  msg_agradecimento: 'Muito obrigado por fazer parte desse dia inesquecível e por compartilhar suas memórias conosco! 💛',
  tema:              'classico',
  cor_accent:        '#B8975A',
  cor_bg:            '#F9F8F6',
  ativo:             'true'
};

// ─────────────────────────────────────────
//  AUTENTICAÇÃO
// ─────────────────────────────────────────

function inicializarAdmin() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('admin_email')) {
    props.setProperty('admin_email',      ADMIN_DEFAULT.email);
    props.setProperty('admin_senha_hash', hashSenha(ADMIN_DEFAULT.senha));
    props.setProperty('admin_nome',       ADMIN_DEFAULT.nome);
  }

  // Garantir que RAFMAR06 está na lista de links
  const linksRaw = props.getProperty('eventa_links');
  const links    = linksRaw ? JSON.parse(linksRaw) : [];
  const temRafa = links.some(l => l.codigo === 'RAFMAR06');
  if (!temRafa) {
    links.unshift({
      id:        'rafmar06-fixo',
      codigo:    'RAFMAR06',
      nome:      'Casamento Rafa & Marcelo',
      link:      '',
      expiry:    '',
      criado_em: new Date().toISOString()
    });
    props.setProperty('eventa_links', JSON.stringify(links));
  }
}

function login(data) {
  const { email, senha } = data;
  if (!email || !senha) return { error: 'Preencha email e senha.' };

  const props      = PropertiesService.getScriptProperties();
  const inputEmail = String(email).trim().toLowerCase();
  const inputHash  = hashSenha(senha);

  // 1. Checa usuários pré-configurados
  const userPadrao = ADMIN_USERS.find(u => u.email.toLowerCase() === inputEmail && hashSenha(u.senha) === inputHash);

  // 2. Checa se o admin alterou senha nas propriedades
  const adminEmail = (props.getProperty('admin_email') || '').toLowerCase();
  const adminHash  = props.getProperty('admin_senha_hash');
  const matchAdmin = (inputEmail === adminEmail && inputHash === adminHash);

  let nomeUsuario = '';
  if (userPadrao) {
    nomeUsuario = userPadrao.nome;
  } else if (matchAdmin) {
    nomeUsuario = props.getProperty('admin_nome') || 'Admin';
  } else {
    return { error: 'Email ou senha incorretos.' };
  }

  // Gerar token de sessão
  const token  = Utilities.getUuid();
  const expiry = new Date();
  expiry.setHours(expiry.getHours() + CONFIG.SESSION_HORAS);

  props.setProperty('session_token',  token);
  props.setProperty('session_expiry', expiry.toISOString());

  return {
    ok:    true,
    token: token,
    nome:  nomeUsuario
  };
}

function verificarSessao(token) {
  if (!token) return false;
  const props  = PropertiesService.getScriptProperties();
  const saved  = props.getProperty('session_token');
  const expiry = props.getProperty('session_expiry');
  if (!saved || !expiry || token !== saved) return false;
  return new Date() <= new Date(expiry);
}

function logout(data) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('session_token');
  props.deleteProperty('session_expiry');
  return { ok: true };
}

function esqueceuSenha(data) {
  const { email, painelUrl } = data;
  const inputEmail = String(email || '').trim().toLowerCase();
  const props      = PropertiesService.getScriptProperties();
  const adminEmail = (props.getProperty('admin_email') || '').toLowerCase();

  const emailsValidos = [...ADMIN_USERS.map(u => u.email.toLowerCase()), adminEmail];

  if (emailsValidos.includes(inputEmail)) {
    const token  = Utilities.getUuid();
    const expiry = new Date();
    expiry.setMinutes(expiry.getMinutes() + CONFIG.RESET_MINUTOS);

    props.setProperty('reset_token',  token);
    props.setProperty('reset_expiry', expiry.toISOString());

    const resetUrl = (painelUrl || '') + '?reset=' + token;

    GmailApp.sendEmail(inputEmail, 'EVENTA — Redefinição de Senha', '', {
      htmlBody: `
        <div style="font-family:'Helvetica Neue',sans-serif;max-width:520px;margin:0 auto;color:#1C1C1E">
          <div style="background:#1C1C1E;padding:32px;border-radius:16px 16px 0 0;text-align:center">
            <p style="color:#B8975A;font-size:13px;letter-spacing:3px;margin:0;text-transform:uppercase">EVENTA</p>
            <p style="color:#fff;font-size:20px;font-weight:600;margin:8px 0 0">Redefinição de Senha</p>
          </div>
          <div style="background:#F9F8F6;padding:32px;border-radius:0 0 16px 16px">
            <p style="font-size:15px;color:#6E6E73;margin:0 0 20px">
              Clique no botão abaixo para criar uma nova senha.<br>
              O link expira em <strong>1 hora</strong>.
            </p>
            <a href="${resetUrl}"
               style="display:block;background:#B8975A;color:#fff;text-decoration:none;padding:16px;border-radius:12px;text-align:center;font-weight:600;font-size:15px">
              Redefinir Minha Senha
            </a>
            <p style="font-size:12px;color:#A8A8AD;margin:20px 0 0;text-align:center">
              Se você não solicitou isso, ignore este email.
            </p>
          </div>
        </div>
      `
    });
  }

  // Sempre retorna ok (não revela se email existe)
  return { ok: true };
}

function redefinirSenha(data) {
  const { token, novaSenha } = data;
  const props       = PropertiesService.getScriptProperties();
  const savedToken  = props.getProperty('reset_token');
  const savedExpiry = props.getProperty('reset_expiry');

  if (!savedToken || token !== savedToken)         return { error: 'Link inválido.' };
  if (new Date() > new Date(savedExpiry))          return { error: 'Link expirado. Solicite um novo.' };
  if (!novaSenha || novaSenha.length < 6)          return { error: 'A nova senha deve ter ao menos 6 caracteres.' };

  props.setProperty('admin_senha_hash', hashSenha(novaSenha));
  props.deleteProperty('reset_token');
  props.deleteProperty('reset_expiry');

  return { ok: true };
}

function hashSenha(senha) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    senha,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ─────────────────────────────────────────
//  CONFIGURAÇÃO DO EVENTO
// ─────────────────────────────────────────

function saveConfig(data) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('eventa_config', JSON.stringify(data));

  // Tenta sincronizar com a planilha de forma segura (se houver permissão)
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_CONFIG);
    sheet.clearContents();
    const campos = [
      ['titulo',            data.titulo            || ''],
      ['subtitulo',         data.subtitulo          || ''],
      ['data_evento',       data.data_evento        || ''],
      ['mensagem',          data.mensagem           || ''],
      ['codigo_evento',     data.codigo_evento      || ''],
      ['msg_agradecimento', data.msg_agradecimento  || 'Muito obrigado pelas memórias! 💛'],
      ['duracao_horas',     data.duracao_horas      || ''],
      ['duracao_fim',       data.duracao_fim        || ''],
      ['gmail_ativo',       data.gmail_ativo        || 'false'],
      ['gmail_dest',        data.gmail_dest         || ''],
      ['whatsapp_ativo',    data.whatsapp_ativo     || 'false'],
      ['noivo_ddd',         data.noivo_ddd          || ''],
      ['noivo_tel',         data.noivo_tel          || ''],
      ['noiva_ddd',         data.noiva_ddd          || ''],
      ['noiva_tel',         data.noiva_tel          || ''],
      ['tema',              data.tema               || 'classico'],
      ['cor_accent',        data.cor_accent         || '#B8975A'],
      ['cor_bg',            data.cor_bg             || '#F9F8F6'],
      ['ativo',             'true'],
      ['criado_em',         new Date().toISOString()]
    ];
    campos.forEach((par, i) => {
      sheet.getRange(i + 1, 1).setValue(par[0]);
      const cell = sheet.getRange(i + 1, 2);
      if (par[0] === 'data_evento' || par[0] === 'duracao_fim' || par[0] === 'criado_em') cell.setNumberFormat('@');
      cell.setValue(par[1]);
    });
  } catch(e) {}

  try { getOrCreateEventFolder(data.codigo_evento); } catch(e) {}
  return { ok: true };
}

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty('eventa_config');
  if (raw) {
    try {
      const cfg = JSON.parse(raw);
      if (cfg && cfg.codigo_evento) return cfg;
    } catch(e) {}
  }
  // Se não existir, salva e retorna o padrão
  props.setProperty('eventa_config', JSON.stringify(EVENTO_DEFAULT));
  return EVENTO_DEFAULT;
}

// ─────────────────────────────────────────
//  DADOS DO EVENTO PARA O CONVIDADO
// ─────────────────────────────────────────

function getEvent(code) {
  const cfg = getConfig();

  if (!cfg.codigo_evento) return { error: 'Evento não configurado' };
  if (cfg.codigo_evento.toString().toUpperCase() !== (code || '').toUpperCase()) {
    return { error: 'Código inválido' };
  }
  if (String(cfg.ativo) !== 'true') return { error: 'Evento encerrado' };

  if (cfg.duracao_fim) {
    if (new Date() > new Date(cfg.duracao_fim)) {
      encerrarEvento();
      return { error: 'Evento encerrado' };
    }
  }

  return {
    ok:               true,
    titulo:           cfg.titulo            || '',
    subtitulo:        cfg.subtitulo         || '',
    data_evento:      String(cfg.data_evento || ''),
    mensagem:         cfg.mensagem          || '',
    msg_agradecimento:cfg.msg_agradecimento || '',
    whatsapp_ativo:   cfg.whatsapp_ativo    || 'false',
    noivo_ddd:        cfg.noivo_ddd         || '',
    noivo_tel:        cfg.noivo_tel         || '',
    noiva_ddd:        cfg.noiva_ddd         || '',
    noiva_tel:        cfg.noiva_tel         || '',
    tema:             cfg.tema              || 'classico',
    cor_accent:       cfg.cor_accent        || '#B8975A',
    cor_bg:           cfg.cor_bg            || '#F9F8F6'
  };
}

function encerrarEvento() {
  const props = PropertiesService.getScriptProperties();
  const cfg   = getConfig();
  cfg.ativo   = 'false';
  props.setProperty('eventa_config', JSON.stringify(cfg));

  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_CONFIG);
    const data  = sheet.getDataRange().getValues();
    data.forEach((row, i) => {
      if (row[0] === 'ativo') sheet.getRange(i + 1, 2).setValue('false');
    });
  } catch(e) {}
}

// ─────────────────────────────────────────
//  UPLOAD DE ARQUIVO
// ─────────────────────────────────────────

function uploadFile(data) {
  const cfg = getConfig();
  const { nome, sobrenome, ddd, telefone, mensagem, fileName, mimeType, fileData, fileSizeMB, codigo } = data;

  if (!cfg.codigo_evento || cfg.codigo_evento.toString().toUpperCase() !== (codigo || '').toUpperCase()) {
    return { error: 'Código inválido' };
  }
  if (String(cfg.ativo) !== 'true') return { error: 'Evento encerrado' };
  if (parseFloat(fileSizeMB) > CONFIG.MAX_FILE_MB) return { error: 'Arquivo muito grande' };

  const folder  = getOrCreateEventFolder(cfg.codigo_evento);
  const decoded = Utilities.base64Decode(fileData);
  const ext      = fileName.split('.').pop().toLowerCase();
  const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const nomeArq  = (nome + '-' + sobrenome + '-' + dataHoje + '.' + ext)
                    .toLowerCase()
                    .normalize('NFD').replace(/[̀-ͯ]/g, '')
                    .replace(/[^a-z0-9.\-]/g, '-')
                    .replace(/-+/g, '-');
  const blob    = Utilities.newBlob(decoded, mimeType, nomeArq);
  const file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const tipo = mimeType.startsWith('video/') ? 'Video' : 'Foto';

  // Grava stats em ScriptProperties (sempre disponível)
  try {
    const props = PropertiesService.getScriptProperties();
    const stats = JSON.parse(props.getProperty('eventa_stats') || '{"fotos":0,"videos":0,"convidados":[],"recentes":[]}');
    if (tipo === 'Video') stats.videos++;
    else stats.fotos++;
    const nomeCompleto = ((nome||'') + ' ' + (sobrenome||'')).trim();
    if (nomeCompleto && !stats.convidados.includes(nomeCompleto)) stats.convidados.push(nomeCompleto);
    stats.recentes.unshift({
      ts: new Date().toISOString(),
      nome: nomeCompleto,
      tipo: tipo,
      arquivo: fileName,
      tamanho: fileSizeMB
    });
    if (stats.recentes.length > 20) stats.recentes = stats.recentes.slice(0, 20);
    props.setProperty('eventa_stats', JSON.stringify(stats));
  } catch(e) {}

  // Grava na planilha se disponível
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_REGISTROS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Nome','Sobrenome','DDD','Telefone','Mensagem','Tipo','Arquivo','Tamanho (MB)','Link Drive']);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#F9F8F6');
    }
    sheet.appendRow([
      new Date(), nome||'', sobrenome||'', ddd||'', telefone||'',
      mensagem||'', tipo, fileName, fileSizeMB, file.getUrl()
    ]);
  } catch(e) {}

  return { ok: true, tipo, fileId: file.getId() };
}

// ─────────────────────────────────────────
//  ESTATÍSTICAS
// ─────────────────────────────────────────

function getStats() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('eventa_stats');
    if (raw) {
      const s = JSON.parse(raw);
      return {
        fotos:      s.fotos || 0,
        videos:     s.videos || 0,
        convidados: (s.convidados || []).length,
        total:      (s.fotos || 0) + (s.videos || 0),
        recentes:   (s.recentes || []).slice(0, 10)
      };
    }
  } catch(e) {}

  return { fotos: 0, videos: 0, convidados: 0, total: 0, recentes: [] };
}

// ─────────────────────────────────────────
//  GMAIL
// ─────────────────────────────────────────

function finalizarEnvio(data) {
  try {
    const cfg = getConfig();

    Logger.log('gmail_ativo: ' + cfg.gmail_ativo);
    Logger.log('gmail_dest: ' + cfg.gmail_dest);

    const gmailAtivo = String(cfg.gmail_ativo).trim().toLowerCase() === 'true';
    const gmailDest  = String(cfg.gmail_dest || '').trim();

    if (!gmailAtivo || !gmailDest) {
      Logger.log('Gmail desativado ou sem destino — email não enviado.');
      return { ok: true, msg: 'gmail_off' };
    }

    const nome         = String(data.nome      || '');
    const sobrenome    = String(data.sobrenome  || '');
    const limpoDDD     = String(data.ddd      || '').replace(/\D/g,'').replace(/^0+/, '');
    const limpoTel     = String(data.telefone || '').replace(/\D/g,'').replace(/^0+/, '');
    const fotos        = parseInt(data.fotos)   || 0;
    const videos       = parseInt(data.videos)  || 0;
    const nomeCompleto = (nome + ' ' + sobrenome).trim();
    const wppNum       = '55' + limpoDDD + limpoTel;
    const wppLink      = 'https://wa.me/' + wppNum;

    const fotoTxt  = fotos  === 1 ? '1 foto'  : fotos  + ' fotos';
    const videoTxt = videos === 1 ? '1 vídeo' : videos + ' vídeos';
    const resumo   = fotos > 0 && videos > 0 ? fotoTxt + ' e ' + videoTxt
                   : fotos  > 0 ? fotoTxt : videoTxt;

    GmailApp.sendEmail(gmailDest, 'Chegaram recordações! ' + nomeCompleto, '', {
      htmlBody: `
        <div style="font-family:'Helvetica Neue',sans-serif;max-width:520px;margin:0 auto;color:#1C1C1E">
          <div style="background:#1C1C1E;padding:32px;border-radius:16px 16px 0 0;text-align:center">
            <p style="color:#B8975A;font-size:13px;letter-spacing:3px;margin:0;text-transform:uppercase">EVENTA</p>
            <p style="color:#fff;font-size:22px;font-weight:600;margin:10px 0 0">Chegaram recordações!</p>
          </div>
          <div style="background:#F9F8F6;padding:32px;border-radius:0 0 16px 16px">
            <p style="font-size:16px;margin:0 0 16px">
              <strong>${nomeCompleto}</strong> acabou de enviar <strong>${resumo}</strong> para o evento.
            </p>
            <p style="font-size:14px;color:#6E6E73;margin:0 0 24px">Entre em contato pelo WhatsApp:</p>
            <a href="${wppLink}"
               style="display:inline-block;background:#25D366;color:#fff;text-decoration:none;padding:14px 28px;border-radius:12px;font-weight:600;font-size:15px;margin-bottom:28px">
              💬 Abrir WhatsApp de ${nomeCompleto}
            </a>
            <hr style="border:none;border-top:1px solid #E8E6E1;margin:0 0 24px"/>
            <p style="font-size:13px;color:#A8A8AD;text-align:center;margin:0">
              Obrigado por escolher a EVENTA para esse momento incrível ✦
            </p>
          </div>
        </div>`
    });

    Logger.log('Email enviado para: ' + gmailDest);
    return { ok: true };

  } catch(err) {
    Logger.log('ERRO finalizarEnvio: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ─────────────────────────────────────────
//  GOOGLE DRIVE
// ─────────────────────────────────────────

function getOrCreateRootFolder() {
  if (CONFIG.FOLDER_ID) {
    try {
      return DriveApp.getFolderById(CONFIG.FOLDER_ID);
    } catch(e) {
      Logger.log('Aviso: pasta por ID não encontrada, usando busca por nome.');
    }
  }
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function getOrCreateEventFolder(codigo) {
  const root = getOrCreateRootFolder();
  // Se já apontamos diretamente para a pasta MARCELO_e_RAFA, grava tudo direto nela!
  if (CONFIG.FOLDER_ID) {
    return root;
  }
  const name   = 'Evento_' + (codigo || 'SEM_CODIGO').toUpperCase();
  const exists = root.getFoldersByName(name);
  if (exists.hasNext()) return exists.next();
  return root.createFolder(name);
}

// ─────────────────────────────────────────
//  GOOGLE SHEETS
// ─────────────────────────────────────────

function getOrCreateSpreadsheet() {
  const root  = getOrCreateRootFolder();
  const files = root.getFilesByName('EVENTA_Dados');
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  const ss = SpreadsheetApp.create('EVENTA_Dados');
  DriveApp.getFileById(ss.getId()).moveTo(root);
  return ss;
}

function getOrCreateSheet(name) {
  const ss    = getOrCreateSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ─────────────────────────────────────────
//  GERENCIAMENTO DE LINKS / CHAVES
// ─────────────────────────────────────────

function saveLink(data) {
  const props = PropertiesService.getScriptProperties();
  const links = JSON.parse(props.getProperty('eventa_links') || '[]');

  // Remove link existente do mesmo evento (evita duplicata)
  const filtrado = links.filter(l => l.codigo !== data.codigo);

  filtrado.unshift({
    id:        Utilities.getUuid(),
    codigo:    data.codigo    || '',
    nome:      data.nome      || '',
    link:      data.link      || '',
    expiry:    data.expiry    || '',
    criado_em: new Date().toISOString()
  });

  props.setProperty('eventa_links', JSON.stringify(filtrado));
  return { ok: true };
}

function getLinks() {
  const props = PropertiesService.getScriptProperties();
  const links = JSON.parse(props.getProperty('eventa_links') || '[]');

  // Remove links expirados automaticamente
  const agora   = new Date();
  const ativos  = links.filter(l => !l.expiry || new Date(l.expiry) > agora);

  if (ativos.length !== links.length) {
    props.setProperty('eventa_links', JSON.stringify(ativos));
  }

  return { ok: true, links: ativos };
}

function deleteLink(data) {
  const props  = PropertiesService.getScriptProperties();
  const links  = JSON.parse(props.getProperty('eventa_links') || '[]');
  const novos  = links.filter(l => l.id !== data.id);
  props.setProperty('eventa_links', JSON.stringify(novos));
  return { ok: true };
}

// ─────────────────────────────────────────
//  TESTE DE EMAIL (rodar manualmente no editor)
// ─────────────────────────────────────────
function testarEmail() {
  GmailApp.sendEmail(
    'othiagoschwanz@gmail.com',
    'EVENTA — Teste de Email',
    'Se chegou, o Gmail está funcionando!'
  );
  Logger.log('Email enviado com sucesso!');
}
