const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'Manual do Administrador — EBD Escala Dominical';

// ─── PALETTE ───────────────────────────────────────────────
const C = {
  green:      '2d6a4f',
  greenMid:   '40916c',
  greenLight: '74c69d',
  greenPale:  'd8f3dc',
  dark:       '18181b',
  ink:        '27272a',
  muted:      '71717a',
  border:     'e4e4e7',
  surface:    'f4f4f5',
  white:      'ffffff',
  amber:      'b45309',
  amberPale:  'fef3c7',
  rose:       'be123c',
  rosePale:   'ffe4e6',
  sky:        '0369a1',
  skyPale:    'e0f2fe',
};

const makeShadow = () => ({ type:'outer', blur:8, offset:3, angle:135, color:'000000', opacity:0.10 });

// ─── HELPERS ───────────────────────────────────────────────
function bgSolid(slide, color) {
  slide.background = { color };
}

function titleBar(slide, label, dotColor) {
  // top accent strip
  slide.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:0.07, fill:{ color: dotColor||C.green }, line:{ color: dotColor||C.green, width:0 } });
}

function sectionTag(slide, text, x, y, bgColor, textColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:1.6, h:0.28, fill:{ color: bgColor||C.greenPale }, line:{ color: bgColor||C.greenPale }, rectRadius:0.06 });
  slide.addText(text.toUpperCase(), { x, y, w:1.6, h:0.28, fontSize:7.5, bold:true, color: textColor||C.green, align:'center', valign:'middle', margin:0 });
}

function card(slide, x, y, w, h, fillColor) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill:{ color: fillColor||C.white },
    line:{ color: C.border, width:0.75 },
    shadow: makeShadow(),
  });
}

function stepCircle(slide, num, x, y, color) {
  slide.addShape(pres.shapes.OVAL, { x, y, w:0.38, h:0.38, fill:{ color: color||C.green }, line:{ color: color||C.green } });
  slide.addText(String(num), { x, y, w:0.38, h:0.38, fontSize:11, bold:true, color:C.white, align:'center', valign:'middle', margin:0 });
}

function pill(slide, text, x, y, bg, fg) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:1.4, h:0.26, fill:{ color: bg||C.greenPale }, line:{ color: bg||C.greenPale }, rectRadius:0.13 });
  slide.addText(text, { x, y, w:1.4, h:0.26, fontSize:8, bold:true, color: fg||C.green, align:'center', valign:'middle', margin:0 });
}

function iconBox(slide, emoji, x, y, bg) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w:0.5, h:0.5, fill:{ color: bg||C.greenPale }, line:{ color: bg||C.greenPale }, rectRadius:0.12 });
  slide.addText(emoji, { x, y, w:0.5, h:0.5, fontSize:18, align:'center', valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 1 — CAPA
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.green);

  // texture dots pattern (subtle circles)
  for (let r=0; r<6; r++) {
    for (let c=0; c<10; c++) {
      sl.addShape(pres.shapes.OVAL, { x: c*1.05-0.2, y: r*1.0-0.1, w:0.08, h:0.08, fill:{ color:'ffffff', transparency:88 }, line:{ color:'ffffff', transparency:88 } });
    }
  }

  // big white card overlay right side
  sl.addShape(pres.shapes.RECTANGLE, { x:5.8, y:0, w:4.2, h:5.625, fill:{ color:C.white }, line:{ color:C.white } });

  // cross icon left
  sl.addText('✝', { x:0.7, y:1.2, w:1.2, h:1.2, fontSize:64, color:C.white, align:'center', valign:'middle', margin:0 });

  // main title left
  sl.addText('Manual do', { x:0.5, y:2.5, w:5.0, h:0.6, fontSize:28, bold:false, color:'d8f3dc', align:'left', valign:'middle', margin:0 });
  sl.addText('Administrador', { x:0.5, y:3.1, w:5.0, h:0.75, fontSize:40, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });
  sl.addText('EBD Escala · Dominical', { x:0.5, y:3.85, w:5.0, h:0.35, fontSize:14, color:'95d5b2', align:'left', valign:'middle', margin:0 });

  // right panel content
  sl.addText('📋', { x:6.3, y:0.7, w:0.8, h:0.8, fontSize:40, align:'center', valign:'middle', margin:0 });
  sl.addText('Guia Completo', { x:6.1, y:1.5, w:3.4, h:0.5, fontSize:22, bold:true, color:C.dark, align:'center', valign:'middle', margin:0 });
  sl.addText('para o responsável pela\nescala de professores', { x:6.1, y:2.05, w:3.4, h:0.6, fontSize:12, color:C.muted, align:'center', valign:'middle', margin:0 });

  // divider
  sl.addShape(pres.shapes.LINE, { x:6.5, y:2.75, w:2.6, h:0, line:{ color:C.border, width:1 } });

  // topic pills
  const topics = [['🔐','Acesso Admin'],['📅','Escalas'],['🏫','Classes'],['⚙️','Configurações']];
  topics.forEach(([ic,lb], i) => {
    const tx = 6.15 + (i%2)*1.75, ty = 3.05 + Math.floor(i/2)*0.55;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:tx, y:ty, w:1.55, h:0.4, fill:{ color:C.surface }, line:{ color:C.border, width:0.5 }, rectRadius:0.08 });
    sl.addText(`${ic} ${lb}`, { x:tx, y:ty, w:1.55, h:0.4, fontSize:9.5, color:C.ink, align:'center', valign:'middle', margin:0 });
  });

  sl.addText('Versão 2025 · v2', { x:6.1, y:5.1, w:3.4, h:0.28, fontSize:8, color:C.muted, align:'center', valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 2 — VISÃO GERAL DO APP
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.green);

  sectionTag(sl, 'Visão Geral', 0.5, 0.22, C.greenPale, C.green);
  sl.addText('O que é o EBD Escala?', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, align:'left', valign:'middle', margin:0 });
  sl.addText('Um app de gestão de escalas para a Escola Bíblica Dominical — sincronizado em tempo real via Firebase.', { x:0.5, y:1.1, w:9, h:0.35, fontSize:12, color:C.muted, align:'left', valign:'middle', margin:0 });

  // 5 feature cards
  const features = [
    { ic:'📅', title:'Próxima EBD',    desc:'Veja e edite a escala\ndo próximo domingo' },
    { ic:'📆', title:'Por Mês',         desc:'Gerencie todos os\ndomingos do mês' },
    { ic:'📋', title:'Histórico',       desc:'Registros de EBDs\njá realizadas' },
    { ic:'🏫', title:'Classes',         desc:'Turmas e professores\ncadastrados' },
    { ic:'⚙️', title:'Configurações',  desc:'Ajustes gerais\ne segurança' },
  ];
  features.forEach((f, i) => {
    const x = 0.45 + i * 1.84;
    card(sl, x, 1.65, 1.65, 2.8);
    sl.addText(f.ic, { x, y:1.75, w:1.65, h:0.7, fontSize:32, align:'center', valign:'middle', margin:0 });
    sl.addText(f.title, { x, y:2.5, w:1.65, h:0.35, fontSize:11, bold:true, color:C.dark, align:'center', valign:'middle', margin:0 });
    sl.addShape(pres.shapes.LINE, { x: x+0.2, y:2.88, w:1.25, h:0, line:{ color:C.border, width:0.5 } });
    sl.addText(f.desc, { x, y:2.93, w:1.65, h:0.45, fontSize:9, color:C.muted, align:'center', valign:'top', margin:4 });
  });

  // modos de acesso
  sl.addText('Modos de Acesso', { x:0.5, y:4.6, w:3, h:0.3, fontSize:11, bold:true, color:C.dark, margin:0 });
  const modos = [
    { ic:'👑', label:'Admin',      desc:'Acesso total — você',      bg:C.amberPale, fg:C.amber },
    { ic:'👤', label:'Professor',  desc:'Confirma própria presença', bg:C.skyPale,   fg:C.sky },
    { ic:'👁️', label:'Visitante', desc:'Apenas leitura',            bg:C.surface,   fg:C.muted },
  ];
  modos.forEach((m, i) => {
    const x = 0.45 + i * 3.1;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:4.95, w:2.85, h:0.45, fill:{ color: m.bg }, line:{ color:m.bg }, rectRadius:0.1 });
    sl.addText(`${m.ic}  ${m.label} — ${m.desc}`, { x: x+0.1, y:4.95, w:2.75, h:0.45, fontSize:9.5, bold:false, color:m.fg, align:'left', valign:'middle', margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 3 — ACESSO ADMIN (PIN)
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.amber);

  sectionTag(sl, 'Acesso', 0.5, 0.22, C.amberPale, C.amber);
  sl.addText('Como Entrar como Administrador', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });

  // left: steps
  const steps = [
    { n:1, txt:'Toque no botão de perfil\nno topo do app', sub:'(mostra o modo atual: "Visitante", "Admin", etc.)' },
    { n:2, txt:'Selecione "Administrador"\nna lista de modos', sub:'Uma tela de PIN será exibida' },
    { n:3, txt:'Digite o PIN\n(padrão: 1234)', sub:'O PIN pode ser alterado em Configurações' },
    { n:4, txt:'Toque em "Entrar\ncomo Admin"', sub:'O app atualiza — botões de edição aparecem' },
  ];
  steps.forEach((s, i) => {
    const y = 1.25 + i * 0.95;
    stepCircle(sl, s.n, 0.5, y+0.05, C.amber);
    sl.addText(s.txt, { x:1.05, y, w:4.0, h:0.45, fontSize:11.5, bold:true, color:C.dark, margin:0 });
    sl.addText(s.sub, { x:1.05, y: y+0.44, w:4.0, h:0.28, fontSize:9, color:C.muted, margin:0 });
    if (i<3) sl.addShape(pres.shapes.LINE, { x:0.68, y: y+0.5, w:0, h:0.42, line:{ color:C.border, width:1 } });
  });

  // right: dicas de segurança
  card(sl, 5.5, 1.15, 4.1, 4.2, C.amberPale);
  sl.addText('🔐', { x:5.5, y:1.25, w:4.1, h:0.55, fontSize:28, align:'center', margin:0 });
  sl.addText('Segurança do PIN', { x:5.6, y:1.82, w:3.9, h:0.35, fontSize:13, bold:true, color:C.amber, align:'center', margin:0 });

  const tips = [
    '🔄  Troque o PIN padrão (1234) logo na primeira vez',
    '🤫  Não compartilhe o PIN com professores',
    '🔒  O PIN é salvo com criptografia no Firebase',
    '📲  Um PIN por dispositivo — não precisa de login',
    '⚠️  Se esquecer, contate quem instalou o sistema',
  ];
  tips.forEach((t, i) => {
    sl.addText(t, { x:5.75, y: 2.3+i*0.4, w:3.6, h:0.35, fontSize:9.5, color:C.ink, margin:0 });
  });

  // also accessible via config button
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:5.05, w:4.7, h:0.38, fill:{ color:C.surface }, line:{ color:C.border, width:0.5 }, rectRadius:0.08 });
  sl.addText('💡  Também acessível pelo botão "Configurações" no final da aba Próxima', { x:0.55, y:5.05, w:4.6, h:0.38, fontSize:9, color:C.muted, align:'left', valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 4 — ESCALANDO PROFESSORES (PRÓXIMA EBD)
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.green);

  sectionTag(sl, 'Próxima EBD', 0.5, 0.22, C.greenPale, C.green);
  sl.addText('Escalando Professores', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('A aba "Próxima" mostra a escala do domingo que está chegando. É aqui que você faz a maior parte do trabalho.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // flow: 3 action cards
  const actions = [
    { ic:'✏️', title:'Editar slot por slot',   steps:['Na aba Próxima, toque em ✏️ Editar no card da classe','Digite o nome do professor','Escolha o status (Confirmado, Pendente…)','Toque Salvar Escala'] },
    { ic:'📝', title:'Editar escala completa', steps:['Na aba Por Mês, toque em ✏️ Editar no domingo','Preencha todos os professores de uma vez','Adicione uma observação se necessário','Toque 💾 Salvar Tudo'] },
    { ic:'➕', title:'Novo domingo / data',    steps:['Na aba Por Mês, toque em + Adicionar Domingo','Escolha a data desejada','Confirme — a escala é criada em branco','Edite depois pelo botão ✏️ Editar'] },
  ];
  actions.forEach((a, i) => {
    const x = 0.45 + i * 3.15;
    card(sl, x, 1.55, 2.9, 3.85);
    // header
    sl.addShape(pres.shapes.RECTANGLE, { x, y:1.55, w:2.9, h:0.55, fill:{ color:C.green }, line:{ color:C.green } });
    sl.addText(`${a.ic}  ${a.title}`, { x: x+0.1, y:1.55, w:2.7, h:0.55, fontSize:10.5, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });
    a.steps.forEach((s, j) => {
      stepCircle(sl, j+1, x+0.15, 2.25+j*0.72, C.greenMid);
      sl.addText(s, { x: x+0.62, y:2.2+j*0.72, w:2.15, h:0.6, fontSize:9, color:C.ink, margin:2 });
    });
  });

  // status legend
  sl.addText('Status disponíveis:', { x:0.5, y:5.5, w:2.2, h:0.25, fontSize:9.5, bold:true, color:C.dark, margin:0 });
  const statuses = [['✓ Confirmado','dcfce7','15803d'],['⏳ Pendente','fef9c3','a16207'],['○ Disponível','dbeafe','1d4ed8'],['✗ Ausente','fee2e2','b91c1c']];
  statuses.forEach(([lbl, bg, fg], i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.85+i*1.75, y:5.48, w:1.55, h:0.27, fill:{ color:bg }, line:{ color:bg }, rectRadius:0.07 });
    sl.addText(lbl, { x: 2.85+i*1.75, y:5.48, w:1.55, h:0.27, fontSize:8.5, bold:true, color:fg, align:'center', valign:'middle', margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 5 — FOTOS DE PROFESSORES
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.sky);

  sectionTag(sl, 'Fotos', 0.5, 0.22, C.skyPale, C.sky);
  sl.addText('Fotos dos Professores', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('Cada professor pode ter uma foto de perfil associada ao nome — exibida em todos os cards e listas do app.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // como funciona - 2 columns
  // left col: adicionar foto
  card(sl, 0.45, 1.55, 4.3, 2.7);
  sl.addShape(pres.shapes.RECTANGLE, { x:0.45, y:1.55, w:4.3, h:0.5, fill:{ color:C.sky }, line:{ color:C.sky } });
  sl.addText('📷  Adicionar ou trocar foto', { x:0.55, y:1.55, w:4.1, h:0.5, fontSize:11, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });

  const addSteps = ['1. No modo Admin, clique em ✏️ Editar em qualquer slot de professor','2. Verifique que o nome do professor está digitado','3. Toque no botão 📷 Foto na área de avatar','4. Selecione uma imagem da galeria / câmera','5. A foto é comprimida (200×200px) e salva automaticamente'];
  addSteps.forEach((s, i) => {
    sl.addText(s, { x:0.6, y: 2.15+i*0.38, w:4.0, h:0.34, fontSize:9.5, color:C.ink, margin:0 });
  });

  // right col: comportamento
  card(sl, 5.25, 1.55, 4.3, 2.7);
  sl.addShape(pres.shapes.RECTANGLE, { x:5.25, y:1.55, w:4.3, h:0.5, fill:{ color:'0891b2' }, line:{ color:'0891b2' } });
  sl.addText('✨  Como funciona o avatar', { x:5.35, y:1.55, w:4.1, h:0.5, fontSize:11, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });

  const behaviors = [
    { ic:'🖼️', txt:'Com foto cadastrada — exibe a foto do professor' },
    { ic:'🔤', txt:'Sem foto — exibe as iniciais do nome com cor única' },
    { ic:'🔗', txt:'A foto é vinculada ao NOME — aparece em todo o app' },
    { ic:'💾', txt:'Salva no Firebase — sincronizada entre dispositivos' },
    { ic:'🗑️', txt:'Pode ser removida com o botão 🗑 Remover' },
  ];
  behaviors.forEach((b, i) => {
    iconBox(sl, b.ic, 5.35, 2.15+i*0.38, C.skyPale);
    sl.addText(b.txt, { x:5.95, y: 2.15+i*0.38, w:3.45, h:0.34, fontSize:9.5, color:C.ink, valign:'middle', margin:0 });
  });

  // tip bottom
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:4.45, w:9.1, h:0.9, fill:{ color:C.skyPale }, line:{ color:C.skyPale }, rectRadius:0.12 });
  sl.addText('💡  Dica', { x:0.65, y:4.52, w:1.0, h:0.28, fontSize:10, bold:true, color:C.sky, margin:0 });
  sl.addText('Se vários professores tiverem o mesmo nome, o avatar aparecerá igual para ambos. Para diferenciar, use o nome completo (ex.: "João Silva" e "João Costa") ao registrar na escala.', { x:0.65, y:4.78, w:8.7, h:0.48, fontSize:9, color:C.sky, margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 6 — GERENCIANDO CLASSES
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.greenMid);

  sectionTag(sl, 'Classes', 0.5, 0.22, C.greenPale, C.green);
  sl.addText('Gerenciando as Turmas', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('A aba "Classes" lista todas as turmas cadastradas com histórico de professores.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // 3 operation cards
  const ops = [
    {
      ic:'➕', title:'Criar nova classe', color: C.green,
      items:['Toque em "+ Classe" (canto superior direito)','Digite o nome da turma (ex: Berçário)','Escolha um emoji representativo','Selecione uma cor de destaque','Confirme em "Adicionar Classe"'],
    },
    {
      ic:'✏️', title:'Editar classe existente', color: C.greenMid,
      items:['Na aba Classes, toque em ✏️ ao lado da turma','Altere nome, emoji ou cor','Confirme em "Salvar"','As mudanças aparecem imediatamente em todas as escalas'],
    },
    {
      ic:'🗑️', title:'Excluir uma classe', color: C.rose,
      items:['Toque em ✏️ na turma que deseja remover','Role até o botão vermelho "Excluir Classe"','Confirme a exclusão — ação irreversível','⚠️ O histórico desta classe não é apagado'],
    },
  ];
  ops.forEach((op, i) => {
    const x = 0.45 + i * 3.15;
    card(sl, x, 1.55, 2.9, 3.45);
    sl.addShape(pres.shapes.RECTANGLE, { x, y:1.55, w:2.9, h:0.5, fill:{ color:op.color }, line:{ color:op.color } });
    sl.addText(`${op.ic}  ${op.title}`, { x: x+0.1, y:1.55, w:2.7, h:0.5, fontSize:10.5, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });
    op.items.forEach((item, j) => {
      sl.addShape(pres.shapes.OVAL, { x: x+0.15, y: 2.17+j*0.52, w:0.14, h:0.14, fill:{ color:op.color }, line:{ color:op.color } });
      sl.addText(item, { x: x+0.37, y: 2.13+j*0.52, w:2.42, h:0.42, fontSize:9, color:C.ink, margin:0 });
    });
  });

  // dica de ordenação
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:5.1, w:9.1, h:0.35, fill:{ color:C.greenPale }, line:{ color:C.greenPale }, rectRadius:0.08 });
  sl.addText('💡  As classes aparecem na ordem em que foram criadas. Para reordenar, exclua e recrie na ordem desejada.', { x:0.6, y:5.1, w:8.9, h:0.35, fontSize:9, color:C.green, valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 7 — HISTÓRICO E ARQUIVAMENTO
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.muted);

  sectionTag(sl, 'Histórico', 0.5, 0.22, C.surface, C.muted);
  sl.addText('Histórico e Arquivamento', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('Após cada EBD realizada, mova a escala para o histórico oficial para manter o registro organizado.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // archive flow - visual step by step
  sl.addText('Como arquivar uma EBD realizada', { x:0.5, y:1.55, w:5.5, h:0.35, fontSize:13, bold:true, color:C.dark, margin:0 });

  const archSteps = [
    { n:1, title:'Acesse as Configurações', desc:'Toque em ⚙️ Configurações (botão no fim da aba Próxima, ou aba Config se já for Admin)' },
    { n:2, title:'Toque em "Arquivar EBD Realizada"', desc:'Uma lista de escalas passadas ainda não arquivadas será exibida' },
    { n:3, title:'Selecione o domingo', desc:'Toque em "Arquivar →" ao lado da data que deseja mover para o histórico' },
    { n:4, title:'Pronto!', desc:'A escala some das escalas ativas e aparece na aba "Histórico" com data e professores' },
  ];

  archSteps.forEach((s, i) => {
    const y = 2.0 + i * 0.78;
    card(sl, 0.45, y, 5.0, 0.65, C.surface);
    stepCircle(sl, s.n, 0.65, y+0.13, C.ink);
    sl.addText(s.title, { x:1.15, y: y+0.04, w:4.1, h:0.27, fontSize:11, bold:true, color:C.dark, margin:0 });
    sl.addText(s.desc,  { x:1.15, y: y+0.34, w:4.1, h:0.26, fontSize:9, color:C.muted, margin:0 });
  });

  // right: o que o historico mostra
  card(sl, 5.9, 1.55, 3.65, 3.1);
  sl.addShape(pres.shapes.RECTANGLE, { x:5.9, y:1.55, w:3.65, h:0.5, fill:{ color:C.ink }, line:{ color:C.ink } });
  sl.addText('📋  O que o histórico mostra', { x:6.0, y:1.55, w:3.45, h:0.5, fontSize:10.5, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });

  const histItems = [
    '📅  Data do domingo arquivado',
    '👤  Nome do professor por classe',
    '✓  Status de confirmação de cada um',
    '📌  Observação especial (se houver)',
    '🗑  Botão para excluir registro (admin)',
  ];
  histItems.forEach((item, i) => {
    sl.addText(item, { x:6.05, y: 2.18+i*0.38, w:3.35, h:0.34, fontSize:9.5, color:C.ink, margin:0 });
  });

  // aviso
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:5.1, w:9.1, h:0.35, fill:{ color:C.rosePale }, line:{ color:C.rosePale }, rectRadius:0.08 });
  sl.addText('⚠️  Escalas passadas que não forem arquivadas continuam aparecendo em "Por Mês". Archive regularmente para manter o app organizado.', { x:0.6, y:5.1, w:8.9, h:0.35, fontSize:9, color:C.rose, valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 8 — CONFIGURAÇÕES GERAIS
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.amber);

  sectionTag(sl, 'Configurações', 0.5, 0.22, C.amberPale, C.amber);
  sl.addText('Configurações do App', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('Personalize o app para a sua igreja. Todas as configurações exigem acesso de Admin.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // 4 config sections
  const configs = [
    {
      ic:'🏛️', title:'Configurações Gerais', bg: C.amberPale, fg: C.amber,
      fields: [
        { label:'Nome da Igreja', desc:'Ex: Igreja Manancial de Vida' },
        { label:'Horário da EBD', desc:'Formato 24h (ex: 09:30)' },
        { label:'Dia da Semana',  desc:'Geralmente Domingo (padrão)' },
      ],
    },
    {
      ic:'🔐', title:'Segurança (PIN)', bg: C.rosePale, fg: C.rose,
      fields: [
        { label:'PIN atual',    desc:'Visível na tela de configurações' },
        { label:'Novo PIN',     desc:'Mínimo 4 dígitos, máximo 8' },
        { label:'Confirmação',  desc:'Toque "🔐 Alterar PIN" para salvar' },
      ],
    },
    {
      ic:'📌', title:'Observação Geral', bg: C.skyPale, fg: C.sky,
      fields: [
        { label:'Texto livre', desc:'Aparece no card superior da aba Próxima' },
        { label:'Exemplo',    desc:'"Trazer material da lição 15"' },
        { label:'Edição',     desc:'Toque em "Editar" no card Observação Geral' },
      ],
    },
    {
      ic:'⏰', title:'Edição Rápida', bg: C.greenPale, fg: C.green,
      fields: [
        { label:'Botão "Editar"', desc:'No card verde "Próxima EBD" no topo' },
        { label:'O que muda',     desc:'Horário e Dia da Semana rapidamente' },
        { label:'Quando usar',    desc:'Quando a EBD muda de horário pontualmente' },
      ],
    },
  ];
  configs.forEach((cfg, i) => {
    const x = 0.45 + (i%2)*4.8, y = 1.55 + Math.floor(i/2)*2.0;
    card(sl, x, y, 4.3, 1.82);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.15, y: y+0.15, w:0.5, h:0.5, fill:{ color:cfg.bg }, line:{ color:cfg.bg }, rectRadius:0.1 });
    sl.addText(cfg.ic, { x: x+0.15, y: y+0.15, w:0.5, h:0.5, fontSize:20, align:'center', valign:'middle', margin:0 });
    sl.addText(cfg.title, { x: x+0.75, y: y+0.2, w:3.4, h:0.4, fontSize:11.5, bold:true, color:C.dark, margin:0 });
    cfg.fields.forEach((f, j) => {
      sl.addText(`• ${f.label}`, { x: x+0.2, y: y+0.72+j*0.34, w:1.4, h:0.3, fontSize:9, bold:true, color:cfg.fg, margin:0 });
      sl.addText(f.desc, { x: x+1.65, y: y+0.72+j*0.34, w:2.5, h:0.3, fontSize:9, color:C.muted, margin:0 });
    });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 9 — FLUXO SEMANAL RECOMENDADO
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.surface);
  titleBar(sl, '', C.green);

  sectionTag(sl, 'Fluxo de Trabalho', 0.5, 0.22, C.greenPale, C.green);
  sl.addText('Rotina Semanal Recomendada', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('Siga este fluxo para manter as escalas sempre organizadas e os professores informados a tempo.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // 5 day timeline
  const days = [
    { day:'Domingo', icon:'⛪', color:C.green,    tasks:['Realize a EBD', 'Observe presenças'] },
    { day:'Segunda', icon:'📦', color:'0369a1',   tasks:['Archive a EBD\ndo domingo passado'] },
    { day:'Terça-Quarta', icon:'📋', color:'7c3aed', tasks:['Verifique a escala\ndo próximo domingo','Defina professores\npendentes'] },
    { day:'Quinta-Sexta', icon:'📲', color: C.amber, tasks:['Confirme com os\nprofessores escalados','Atualize status\npara Confirmado'] },
    { day:'Sábado', icon:'✅', color:C.greenMid, tasks:['Verifique escala\nfinal','Adicione observação\nse necessário'] },
  ];

  days.forEach((d, i) => {
    const x = 0.45 + i * 1.83;
    // column header
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y:1.55, w:1.65, h:0.55, fill:{ color:d.color }, line:{ color:d.color }, rectRadius:0.1 });
    sl.addText(d.icon, { x, y:1.55, w:1.65, h:0.3, fontSize:18, align:'center', margin:0 });
    sl.addText(d.day, { x, y:1.85, w:1.65, h:0.25, fontSize:8.5, bold:true, color:C.white, align:'center', margin:0 });

    // tasks
    card(sl, x, 2.18, 1.65, 3.2, C.white);
    d.tasks.forEach((t, j) => {
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.1, y: 2.3+j*1.1, w:1.45, h:0.9, fill:{ color: d.color, transparency:90 }, line:{ color:d.color, width:0.5 }, rectRadius:0.08 });
      sl.addText(t, { x: x+0.12, y: 2.3+j*1.1, w:1.41, h:0.9, fontSize:9, color:C.ink, align:'center', valign:'middle', margin:4 });
    });

    // connector arrow (except last)
    if (i<4) {
      sl.addShape(pres.shapes.LINE, { x: x+1.65, y:1.82, w:0.18, h:0, line:{ color:C.border, width:1.5 } });
      sl.addText('›', { x: x+1.74, y:1.72, w:0.15, h:0.2, fontSize:12, color:C.border, align:'center', margin:0 });
    }
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 10 — ACESSO DOS PROFESSORES
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.sky);

  sectionTag(sl, 'Professores', 0.5, 0.22, C.skyPale, C.sky);
  sl.addText('Como os Professores Usam o App', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });
  sl.addText('Os professores usam o mesmo link/app — sem PIN. Eles selecionam o próprio nome para ver e confirmar presenças.', { x:0.5, y:1.1, w:9, h:0.32, fontSize:11, color:C.muted, margin:0 });

  // left: passo a passo do professor
  sl.addText('Passo a passo do professor', { x:0.5, y:1.55, w:4.5, h:0.35, fontSize:13, bold:true, color:C.dark, margin:0 });

  const profSteps = [
    { n:1, txt:'Abrir o mesmo link do app' },
    { n:2, txt:'Tocar no botão de perfil (topo direito)' },
    { n:3, txt:'Selecionar o próprio nome na lista' },
    { n:4, txt:'Visualizar em quais escalas está' },
    { n:5, txt:'Tocar em "✔ Confirmar" no domingo desejado' },
  ];
  profSteps.forEach((s, i) => {
    const y = 2.0 + i * 0.6;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y, w:4.5, h:0.5, fill:{ color:C.surface }, line:{ color:C.border, width:0.5 }, rectRadius:0.1 });
    stepCircle(sl, s.n, 0.65, y+0.06, C.sky);
    sl.addText(s.txt, { x:1.15, y: y+0.05, w:3.65, h:0.38, fontSize:10.5, color:C.ink, valign:'middle', margin:0 });
  });

  // right: diferença de permissões
  card(sl, 5.3, 1.55, 4.25, 3.9);
  sl.addShape(pres.shapes.RECTANGLE, { x:5.3, y:1.55, w:4.25, h:0.5, fill:{ color:C.sky }, line:{ color:C.sky } });
  sl.addText('O que cada perfil pode fazer', { x:5.4, y:1.55, w:4.05, h:0.5, fontSize:11, bold:true, color:C.white, align:'left', valign:'middle', margin:0 });

  const perms = [
    { label:'Ver escalas',         admin:true,  prof:true,  visit:true },
    { label:'Confirmar presença',  admin:true,  prof:true,  visit:false },
    { label:'Editar professor',    admin:true,  prof:false, visit:false },
    { label:'Criar/excluir EBD',   admin:true,  prof:false, visit:false },
    { label:'Gerenciar classes',   admin:true,  prof:false, visit:false },
    { label:'Alterar PIN',         admin:true,  prof:false, visit:false },
    { label:'Arquivar histórico',  admin:true,  prof:false, visit:false },
  ];

  // header row
  ['Ação','👑 Admin','👤 Prof.','👁️ Visit.'].forEach((h, i) => {
    sl.addText(h, { x: 5.4+[0,1.65,2.6,3.4][i], y:2.16, w:[1.55,0.85,0.8,0.7][i], h:0.3, fontSize:8.5, bold:true, color:C.muted, align:i===0?'left':'center', margin:0 });
  });
  sl.addShape(pres.shapes.LINE, { x:5.4, y:2.48, w:3.95, h:0, line:{ color:C.border, width:0.5 } });

  perms.forEach((p, i) => {
    const y = 2.55 + i*0.38;
    const bg = i%2===0 ? C.surface : C.white;
    sl.addShape(pres.shapes.RECTANGLE, { x:5.3, y, w:4.25, h:0.37, fill:{ color:bg }, line:{ color:bg } });
    sl.addText(p.label, { x:5.42, y, w:1.6, h:0.37, fontSize:9, color:C.ink, valign:'middle', margin:0 });
    [p.admin, p.prof, p.visit].forEach((has, j) => {
      sl.addText(has ? '✓' : '—', { x: [7.02,7.88,8.62][j], y, w:0.5, h:0.37, fontSize:10, bold:has, color: has ? C.green : C.border, align:'center', valign:'middle', margin:0 });
    });
  });

  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.45, y:5.1, w:4.5, h:0.35, fill:{ color:C.skyPale }, line:{ color:C.skyPale }, rectRadius:0.08 });
  sl.addText('💡  Compartilhe o link do app com todos os professores — nenhum cadastro extra necessário.', { x:0.6, y:5.1, w:4.35, h:0.35, fontSize:9, color:C.sky, valign:'middle', margin:0 });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 11 — PERGUNTAS FREQUENTES
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.white);
  titleBar(sl, '', C.green);

  sectionTag(sl, 'FAQ', 0.5, 0.22, C.greenPale, C.green);
  sl.addText('Perguntas Frequentes', { x:0.5, y:0.55, w:9, h:0.55, fontSize:26, bold:true, color:C.dark, margin:0 });

  const faqs = [
    { q:'Esqueci o PIN. O que faço?', a:'Acesse o Firebase Console do projeto e edite o campo "adminPin" em dados_app > ebd_app_v2. Ou peça ao desenvolvedor que instalou o sistema.' },
    { q:'As escalas aparecem para todos automaticamente?', a:'Sim! O Firebase sincroniza em tempo real. Qualquer alteração feita pelo Admin aparece instantaneamente nos dispositivos de todos.' },
    { q:'Posso usar em celular e computador ao mesmo tempo?', a:'Sim. O app é um PWA (Progressive Web App) que funciona em qualquer navegador moderno — iOS, Android, Windows, Mac.' },
    { q:'Como adiciono uma data fora do domingo padrão?', a:'Na aba "Por Mês", toque em "+ Adicionar Domingo ao Mês" e selecione qualquer data. Útil para EBDs especiais em sábados ou feriados.' },
    { q:'O app funciona sem internet?', a:'Parcialmente. Os dados ficam em cache local. Mas para sincronizar mudanças, é preciso reconectar. Um banner de "modo offline" avisa quando não há conexão.' },
    { q:'Como excluir um registro do histórico?', a:'Na aba Histórico, estando como Admin, cada card tem um botão 🗑 vermelho para exclusão. A ação não pode ser desfeita.' },
  ];

  faqs.forEach((f, i) => {
    const x = 0.45 + (i%2)*4.8;
    const y = 1.55 + Math.floor(i/2)*1.45;
    card(sl, x, y, 4.3, 1.3);
    sl.addShape(pres.shapes.OVAL, { x: x+0.15, y: y+0.15, w:0.3, h:0.3, fill:{ color:C.green }, line:{ color:C.green } });
    sl.addText('?', { x: x+0.15, y: y+0.15, w:0.3, h:0.3, fontSize:12, bold:true, color:C.white, align:'center', valign:'middle', margin:0 });
    sl.addText(f.q, { x: x+0.55, y: y+0.12, w:3.6, h:0.35, fontSize:10, bold:true, color:C.dark, margin:0 });
    sl.addShape(pres.shapes.LINE, { x: x+0.15, y: y+0.52, w:4.0, h:0, line:{ color:C.border, width:0.5 } });
    sl.addText(f.a, { x: x+0.15, y: y+0.58, w:4.0, h:0.62, fontSize:9, color:C.muted, margin:0 });
  });
}

// ══════════════════════════════════════════════════════════════
// SLIDE 12 — ENCERRAMENTO
// ══════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  bgSolid(sl, C.green);

  // texture
  for (let r=0; r<6; r++) {
    for (let c=0; c<10; c++) {
      sl.addShape(pres.shapes.OVAL, { x: c*1.05-0.2, y: r*1.0-0.1, w:0.08, h:0.08, fill:{ color:'ffffff', transparency:88 }, line:{ color:'ffffff', transparency:88 } });
    }
  }

  sl.addShape(pres.shapes.RECTANGLE, { x:2.3, y:1.1, w:5.4, h:3.4, fill:{ color:C.white, transparency:8 }, line:{ color:'ffffff', transparency:70 } });

  sl.addText('✝', { x:3.8, y:1.3, w:2.4, h:1.1, fontSize:56, color:'d8f3dc', align:'center', valign:'middle', margin:0 });
  sl.addText('Deus abençoe seu\nministério na EBD!', { x:2.5, y:2.35, w:5.0, h:0.95, fontSize:24, bold:true, color:C.white, align:'center', valign:'middle', margin:0 });
  sl.addShape(pres.shapes.LINE, { x:3.5, y:3.38, w:3.0, h:0, line:{ color:'95d5b2', width:1 } });
  sl.addText('"Instruí o jovem no caminho em que deve andar,\ne, quando envelhecer, não se desviará dele."\nProvérbios 22:6', { x:2.5, y:3.48, w:5.0, h:0.7, fontSize:9.5, color:'95d5b2', italic:true, align:'center', margin:0 });

  sl.addText('EBD Escala · Dominical  ·  Manual do Administrador  ·  v2025', { x:0.5, y:5.2, w:9, h:0.25, fontSize:8, color:'52b788', align:'center', margin:0 });
}

// ─── WRITE ─────────────────────────────────────────────────
pres.writeFile({ fileName: '/mnt/user-data/outputs/Manual_Admin_EBD.pptx' })
  .then(() => console.log('✅  Manual_Admin_EBD.pptx gerado com sucesso!'))
  .catch(e => { console.error('❌', e); process.exit(1); });
