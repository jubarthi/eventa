// ============================================================
//  EVENTA — Backend Google Apps Script
//  Versão 1.1 | Com Autenticação
// ============================================================

const CONFIG = {
  FOLDER_NAME:     'EVENTA',
  SHEET_REGISTROS: 'Registros',
  SHEET_CONFIG:    'Configuracao',
  MAX_FILE_MB:     20,
  SESSION_HORAS:   24,
  RESET_MINUTOS:   60
};

// Credenciais padrão do administrador
// (usadas apenas na primeira execução — depois ficam no Script Properties)
const ADMIN_DEFAULT = {
  email: 'thiago@jubarthi.com.br',
  senha: '@#TDSe2026*',
  nome:  'Thiago'
};

// ─────────────────────────────────────────
//  ROTEADOR GET
// ─────────────────────────────────────────

function doGet(e) {
  inicializarAdmin();
  const action = e.parameter.action || '';
  try {
    switch (action) {
      // Públicas (convidado)
      case 'getEvent': return respond(getEvent(e.parameter.code));
      case 'ping':     return respond({ ok: true, ts: new Date().toISOString() });

      // Protegidas (painel)
      case 'getStats':  return respondAuth(e, () => getStats());
      case 'getConfig': return respondAuth(e, () => getConfig());
      case 'getLinks':  return respondAuth(e, () => getLinks());

      default: return respond({ error: 'Ação inválida' });
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
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || '';
    switch (action) {
      // Públicas (convidado)
      case 'uploadFile': return respond(uploadFile(data));

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
}

function login(data) {
  const { email, senha } = data;
  const props      = PropertiesService.getScriptProperties();
  const adminEmail = props.getProperty('admin_email');
  const adminHash  = props.getProperty('admin_senha_hash');

  if (!email || !senha) return { error: 'Preencha email e senha.' };
  if (email.toLowerCase() !== adminEmail.toLowerCase() || hashSenha(senha) !== adminHash) {
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
    nome:  props.getProperty('admin_nome') || 'Admin'
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
  const props      = PropertiesService.getScriptProperties();
  const adminEmail = props.getProperty('admin_email');

  // Não revelar se email existe (segurança)
  if (email && email.toLowerCase() === adminEmail.toLowerCase()) {
    const token  = Utilities.getUuid();
    const expiry = new Date();
    expiry.setMinutes(expiry.getMinutes() + CONFIG.RESET_MINUTOS);

    props.setProperty('reset_token',  token);
    props.setProperty('reset_expiry', expiry.toISOString());

    const resetUrl = (painelUrl || '') + '?reset=' + token;

    GmailApp.sendEmail(adminEmail, 'EVENTA — Redefinição de Senha', '', {
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
    sheet.getRange(i + 1, 2).setValue(par[1]);
  });

  getOrCreateEventFolder(data.codigo_evento);
  return { ok: true };
}

function getConfig() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  const cfg   = {};
  data.forEach(row => { if (row[0]) cfg[row[0]] = row[1]; });
  return cfg;
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
    data_evento:      cfg.data_evento instanceof Date ? Utilities.formatDate(cfg.data_evento, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(cfg.data_evento || ''),
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
  const sheet = getOrCreateSheet(CONFIG.SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  data.forEach((row, i) => {
    if (row[0] === 'ativo') sheet.getRange(i + 1, 2).setValue('false');
  });
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
  const blob    = Utilities.newBlob(decoded, mimeType, fileName);
  const file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const tipo = mimeType.startsWith('video/') ? 'Video' : 'Foto';

  const sheet = getOrCreateSheet(CONFIG.SHEET_REGISTROS);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp','Nome','Sobrenome','DDD','Telefone','Mensagem','Tipo','Arquivo','Tamanho (MB)','Link Drive']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#F9F8F6');
  }
  sheet.appendRow([
    new Date(), nome||'', sobrenome||'', ddd||'', telefone||'',
    mensagem||'', tipo, fileName, fileSizeMB, file.getUrl()
  ]);

  if (cfg.gmail_ativo === 'true' && cfg.gmail_dest) {
    try { enviarEmailNotificacao(cfg, nome, sobrenome, tipo, fileName, fileSizeMB); } catch(e) {}
  }

  return { ok: true, tipo, fileId: file.getId() };
}

// ─────────────────────────────────────────
//  ESTATÍSTICAS
// ─────────────────────────────────────────

function getStats() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_REGISTROS);
  const rows  = sheet.getLastRow();

  if (rows <= 1) return { fotos: 0, videos: 0, convidados: 0, total: 0, recentes: [] };

  const data     = sheet.getRange(2, 1, rows - 1, 10).getValues();
  let fotos      = 0, videos = 0;
  const nomes    = new Set();
  const recentes = [];

  data.forEach(row => {
    if (!row[0]) return;
    const tipo = (row[6] || '').toString();
    if (tipo === 'Foto')  fotos++;
    if (tipo === 'Video') videos++;
    const nome = ((row[1]||'') + ' ' + (row[2]||'')).trim();
    if (nome) nomes.add(nome);
    recentes.push({ ts: row[0] ? new Date(row[0]).toISOString() : '', nome, tipo, arquivo: row[7]||'', tamanho: row[8]||'' });
  });

  recentes.sort((a, b) => new Date(b.ts) - new Date(a.ts));
  return { fotos, videos, convidados: nomes.size, total: fotos + videos, recentes: recentes.slice(0, 10) };
}

// ─────────────────────────────────────────
//  GMAIL
// ─────────────────────────────────────────

function enviarEmailNotificacao(cfg, nome, sobrenome, tipo, arquivo, tamanho) {
  const nomeCompleto = (nome + ' ' + sobrenome).trim();
  GmailApp.sendEmail(cfg.gmail_dest, `📸 EVENTA — Nova ${tipo} de ${nomeCompleto}`, '', {
    htmlBody: `
      <div style="font-family:'Helvetica Neue',sans-serif;max-width:520px;margin:0 auto">
        <div style="background:#1C1C1E;padding:32px;border-radius:16px 16px 0 0;text-align:center">
          <p style="color:#B8975A;font-size:13px;letter-spacing:3px;margin:0">EVENTA</p>
          <p style="color:#fff;font-size:20px;font-weight:600;margin:8px 0 0">Nova memória recebida</p>
        </div>
        <div style="background:#F9F8F6;padding:32px;border-radius:0 0 16px 16px">
          <p>De: <strong>${nomeCompleto}</strong></p>
          <p>Tipo: <strong>${tipo}</strong></p>
          <p>Arquivo: <strong>${arquivo}</strong></p>
          <p>Tamanho: <strong>${tamanho} MB</strong></p>
        </div>
      </div>`
  });
}

// ─────────────────────────────────────────
//  GOOGLE DRIVE
// ─────────────────────────────────────────

function getOrCreateRootFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function getOrCreateEventFolder(codigo) {
  const root   = getOrCreateRootFolder();
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
