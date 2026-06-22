// ═══════════════════════════════════════════════════════════════════════════════
//  PPT AUTOMATION — Migração de template para apresentações GSlides
// ═══════════════════════════════════════════════════════════════════════════════
//  O script copia cada slide da apresentação ativa para um novo arquivo baseado
//  no template definido em TEMPLATE_ID, mapeando automaticamente o tipo de cada
//  slide (capa, seção, padrão escuro, claro, obrigado) para o slide de molde
//  correspondente e replicando o conteúdo com a fonte e cores do novo template.
//
//  O processamento é feito em lotes (definido em BATCH_SIZE) com triggers agendados 
//  para contornar o limite de execução de ~6 minutos do Apps Script.
// ───────────────────────────────────────────────────────────────────────────────
//  FLUXO DE USO
// ───────────────────────────────────────────────────────────────────────────────
//
//  Abra o arquivo de ORIGEM no GSlides e execute as funções abaixo em ordem, uma a uma.
//  As funções do passo 1, 2 e 3 são de diagnóstico apenas leem — não alteram nenhum arquivo.
//
//  PASSO 1 — verificarDimensoes()
//    Confirma que origem e template têm as mesmas dimensões de página (ex: 960×540pt).
//    Se forem diferentes, todo o conteúdo copiado ficará desalinhado.
//    → Ajuste o template ou a origem antes de continuar.
//
//  PASSO 2 — identificarImagensTemplate()
//    Lista todas as imagens da apresentação com frequência de aparição.
//    Imagens marcadas com 🚩 (3+ slides) são candidatas a logos/ícones do template
//    antigo e devem estar em IMAGENS_BLOQUEADAS_TITULO ou IMAGENS_BLOQUEADAS_DIMENSOES.
//    → Revise a lista e atualize os bloqueios se necessário.
//
//  PASSO 3 — estatisticasBackground()
//    Mostra a distribuição de tipos de fundo (RGB, THEME, GRADIENT) por contagem.
//    Útil para identificar cores inesperadas que possam ser mal classificadas como
//    claro ou escuro — e calibrar LAYOUTS_WHITE ou os limiares em slideEhClaro.
//    → Se encontrar fundos estranhos, rode debugWhiteSlides() ou debugBlank8()
//      para investigar slide a slide antes de prosseguir.
//
//  PASSO 4 — transformarComTemplate()
//    Inicia o job. O primeiro batch roda imediatamente; os seguintes são agendados
//    automaticamente a cada ~15s até o processamento terminar.
//    → Não feche o editor de scripts durante a execução.
//    → Para cancelar um job em andamento, execute limparJob().
//
//  PASSO 5 — Acompanhamento
//    Após a primeira execução iniciar, acompanhe o progresso pelo painel de execuções
//    (Execuções > Ver execuções do projeto). Cada batch loga os slides processados
//    e o tipo detectado para cada um. Ao final, o log exibe a URL do arquivo criado.
//
// ───────────────────────────────────────────────────────────────────────────────
//  FUNÇÕES DE DIAGNÓSTICO DISPONÍVEIS
// ───────────────────────────────────────────────────────────────────────────────
//
//  verificarDimensoes()          — compara dimensões de página origem vs template
//  identificarImagensTemplate()  — lista imagens por frequência (detecta logos recorrentes)
//  estatisticasBackground()      — distribuição de tipos de fundo por contagem
//  listarLayoutsTemplate()       — lista nomes de layout do template (para configurar LAYOUTS_*)
//  listarSlidesTemplate()        — lista slides do template com prévia de texto (para conferir IDX_*)
//  debugWhiteSlides()            — mostra como cada slide é classificado (white/escuro/gradient)
//  debugBlank8()                 — foca nos slides BLANK_8 para distinguir white vs black
//  debugImagensBloqueadas()      — mostra quais imagens seriam bloqueadas e por quê
//  debugSlideEspecifico(N)       — dump completo do slide N: fundo, layout, runs, imagens
//
// ───────────────────────────────────────────────────────────────────────────────
//  AVISO
// ───────────────────────────────────────────────────────────────────────────────
//
//  Este script está em construção. As heurísticas de classificação de slides
//  (detecção de fundo claro, seção, fonte grande etc.) foram calibradas para um
//  conjunto específico de apresentações e podem não funcionar corretamente em
//  outros arquivos sem ajustes.
//
//  Antes de usar em uma apresentação nova, rode os passos 1–3 acima, revise os
//  resultados e ajuste as constantes e limiares conforme necessário. Resultados
//  inesperados no arquivo final quase sempre têm origem em alguma variável ou
//  threshold que precisa ser revisado para o contexto da nova apresentação.
//
// ═══════════════════════════════════════════════════════════════════════════════

// ─── CONFIGURAÇÕES ───────────────────────────────────────
//
// ⚠️  ANTES DE RODAR QUALQUER FUNÇÃO:
//  1. Preencha TEMPLATE_ID com o ID do seu template.
//  2. Execute verificarDimensoes(), identificarImagensTemplate() e
//     estatisticasBackground() para mapear a apresentação de origem.
//  3. Ajuste os índices IDX_*, layouts LAYOUTS_* e listas de bloqueio
//     conforme o resultado do mapeamento.
//  Só então execute transformarComTemplate().
//
const FONTE_NOVA   = "Inter";
const TEMPLATE_ID  = " "; // ID do template (parte da URL após /d/ e antes de /edit)

// Índices dos slides no template (0 = primeiro).
// Estes slides servem de "molde": o conteúdo do original é copiado para cima deles.
// Se o template mudar de ordem, basta ajustar estes índices.
const IDX_CAPA    = 1;  // slide de capa (azul)
const IDX_SECAO   = 8;  // slide de seção (curvas à direita)
const IDX_PADRAO  = 9;  // slide padrão dark
const IDX_WHITE   = 5;  // slide claro (fundo claro)
const IDX_THANKS  = 10; // slide Thank You

// Nomes de layout são mais estáveis que índices — vêm do próprio arquivo .pptx e não mudam
// ao reordenar slides. Use listarLayoutsTemplate() para descobrir os nomes do template atual.
const LAYOUTS_CAPA  = ["BLANK_1_2"];
const LAYOUTS_SECAO = ["TITLE_5", "TITLE_10"];
// BLANK_8 pode ser claro ou escuro dependendo do fundo real; o fallback visual em slideEhClaro
// resolve os casos ambíguos quando o layout sozinho não basta.
const LAYOUTS_WHITE = ["BLANK_8"];

// Texto branco sobre fundo escuro; texto escuro (#1c1d2f) sobre fundo claro.
// A cor é aplicada uniformemente a todos os runs — cores originais são descartadas
// para garantir contraste correto no novo template.
const COR_TEXTO_ESCURO = "#FFFFFF";
const COR_TEXTO_CLARO  = "#1c1d2f";

// Logos e ícones do template antigo que não devem migrar para o novo.
// Bloqueio por título: confiável quando a imagem tem metadado de nome.
// → Preencha esta lista com os títulos retornados por identificarImagensTemplate()
//   (coluna title:"..."). Candidatos: imagens marcadas com 🚩 (3+ slides).
const IMAGENS_BLOQUEADAS_TITULO = [
  "Prancheta 6@2x.png",
  "Prancheta 6@1.5x.png",
  "Prancheta 7 cópia@2x.png",
  "Prancheta 7@1.5x.png",
  "Brand@2x.png",
  "Artboard 1@2x.png",
  "Prancheta 6.png",
];

// Bloqueio por dimensão: fallback para imagens sem título (decorações do template antigo
// que o Slides exporta sem metadado). Só aplica quando title está vazio para não bloquear
// conteúdo legítimo que coincida com essas dimensões.
const IMAGENS_BLOQUEADAS_DIMENSOES = [
  { w: 38,  h: 38  },
  { w: 44,  h: 44  },
  { w: 80,  h: 28  },
  // { w: 297, h: 284 },
  // { w: 261, h: 249 },
];

// O Slides API não preserva o padding interno do shape original ao usar insertTextBox.
// Este valor compensa a diferença expandindo a caixa em todas as direções.
const PADDING = 7.2;

// Caixas de corpo/lista têm font size variado no original; forçar 14pt garante
// que o texto caiba no espaço do template sem estourar a caixa.
const FONTE_CORPO_LISTA = 14;

// ─── LIMIARES E TIMEOUTS ─────────────────────────────────
// Limiar RGB para detectar fundo claro (>240 = branco/off-white puro)
const LIMIAR_CLARO_RGB = 240;
// Limiar RGB para detectar texto escuro (<128 = preto/off-black escuro)
const LIMIAR_ESCURO_RGB = 128;
// Limiar de font size para distinguir títulos (≥40pt) de corpo (≤22pt)
const LIMIAR_TITULO_FONTSIZE = 40;
const LIMIAR_CORPO_FONTSIZE = 22;
// Percentual de runs escuros na heurística de texto
const PERCENTUAL_RUNS_ESCUROS = 0.6;
// Limiar de shapes com texto para detectar slides de seção (≤3 shapes)
const LIMIAR_SHAPES_SECAO = 3;
// Line spacing: máximo permitido (cap) e padrão fallback
const LINE_SPACING_MAX = 130;
const LINE_SPACING_DEFAULT = 115;

// ─── DELAYS E THROTTLES ──────────────────────────────────
// Apps Script tem limite de ~6 min por execução. Com 800ms de sleep por slide,
// 50 slides consomem ~40s só em espera — longe do limite, mas seguro para
// apresentações grandes onde cada slide pode ter muitas chamadas de API.
const BATCH_SIZE = 50;
// Delay antes de abrir arquivo após makeCopy (eventual consistency)
const DELAY_APOS_MAKECOPY_MS = 3000;
// Throttle entre slides para não estourar quota de chamadas
const THROTTLE_POR_SLIDE_MS = 800;
// Throttle ao remover slides (mais conservador)
const THROTTLE_AO_REMOVER_MS = 500;
// Delay em caso de erro ao remover
const DELAY_ERRO_REMOCAO_MS = 2000;
// Delay do próximo batch agendado (segundos)
const DELAY_PROXIMO_BATCH_S = 15;
// ─────────────────────────────────────────────────────────


function transformarComTemplate() {
  const props = PropertiesService.getScriptProperties();

  // ScriptProperties persistem entre execuções — se o job foi interrompido
  // (timeout, erro, execução manual repetida), retoma de onde parou em vez
  // de criar um segundo arquivo destino.
  if (props.getProperty('job_destinoId')) {
    Logger.log("📋 Job em andamento detectado. Continuando...");
    processarBatch();
    return;
  }

  // ── Inicia novo job ──
  const origem       = SlidesApp.getActivePresentation();
  const nomeOrigem   = origem.getName();
  const NOME_ARQUIVO = "FIX_ID_" + nomeOrigem;
  const slidesOrigem = origem.getSlides();
  const total        = slidesOrigem.length;

  Logger.log("🚀 Criando arquivo baseado no template...");
  const arquivo = DriveApp.getFileById(TEMPLATE_ID).makeCopy(NOME_ARQUIVO);

  // Drive é eventually consistent: abrir o arquivo imediatamente após makeCopy
  // pode retornar uma apresentação ainda não inicializada.
  Utilities.sleep(DELAY_APOS_MAKECOPY_MS);

  const destino = SlidesApp.openById(arquivo.getId());
  const templateSlides = destino.getSlides();

  // Salva os objectIds dos slides do template antes de qualquer appendSlide.
  // objectId é estável mesmo após reordenações; índice não é.
  // Este array é a "lista de remoção" usada no final para limpar o molde.
  const templateSlideIds = templateSlides.map(s => s.getObjectId());

  props.setProperty('job_destinoId', arquivo.getId());
  props.setProperty('job_origemId', origem.getId());
  props.setProperty('job_templateIds', JSON.stringify(templateSlideIds));
  props.setProperty('job_currentIndex', '0');
  props.setProperty('job_total', String(total));

  Logger.log(`📋 Job iniciado. ${total} slides. Batch de ${BATCH_SIZE}.`);
  processarBatch();
}


function processarBatch() {
  const props = PropertiesService.getScriptProperties();
  const destinoId        = props.getProperty('job_destinoId');
  const origemId         = props.getProperty('job_origemId');
  const templateSlideIds = JSON.parse(props.getProperty('job_templateIds'));
  const currentIndex     = parseInt(props.getProperty('job_currentIndex'));
  const total            = parseInt(props.getProperty('job_total'));

  const endIndex = Math.min(currentIndex + BATCH_SIZE, total);

  const origem  = SlidesApp.openById(origemId);
  const destino = SlidesApp.openById(destinoId);

  // Reconstrói referências aos slides de template por objectId — não por índice,
  // porque appendSlide empurra os slides existentes e invalida qualquer índice salvo.
  const allDestinoSlides = destino.getSlides();
  const tplMap = {};
  allDestinoSlides.forEach(s => {
    const id = s.getObjectId();
    if (templateSlideIds.includes(id)) tplMap[id] = s;
  });

  const slideCapa   = tplMap[templateSlideIds[IDX_CAPA]];
  const slideSecao  = tplMap[templateSlideIds[IDX_SECAO]];
  const slidePadrao = tplMap[templateSlideIds[IDX_PADRAO]];
  const slideWhite  = tplMap[templateSlideIds[IDX_WHITE]];
  const slideThanks = tplMap[templateSlideIds[IDX_THANKS]];

  const slidesOrigem = origem.getSlides();

  Logger.log(`📋 Processando slides ${currentIndex + 1} a ${endIndex} de ${total}...`);

  for (let i = currentIndex; i < endIndex; i++) {
    const slideOrigem = slidesOrigem[i];
    Logger.log(`   ${i + 1} de ${total}...`);

    // Slides ocultos no original (ex: rascunhos, slides de backup) são preservados
    // ocultos no destino sem transformação — não faz sentido retemplatizar conteúdo
    // que o apresentador decidiu esconder.
    if (slideOrigem.isSkipped()) {
      try {
        const copia = destino.appendSlide(slideOrigem);
        copia.setSkipped(true);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} oculto: ${e.message}`); }
      continue;
    }

    // Concatena todo o texto do slide em maiúsculas para detecção por palavra-chave.
    // É uma heurística rápida — não depende de layout name nem de cor de fundo.
    const textoSlide = slideOrigem.getShapes().map(s => {
      try { return s.getText().asString(); } catch(e) { return ""; }
    }).join(" ").toUpperCase();

    const isObrigado   = textoSlide.includes("OBRIGADO") || textoSlide.includes("THANK YOU");
    const isComunidade = textoSlide.includes("JUNTE-SE");

    const shapesComTexto = slideOrigem.getShapes().filter(s => {
      try { return s.getText().asString().trim().length > 0; } catch(e) { return false; }
    }).length;

    // Slides de módulo costumam ter título + número + talvez subtítulo (≤LIMIAR_SHAPES_SECAO shapes).
    // O limite evita que slides de conteúdo que mencionam "módulo" em rodapé
    // sejam classificados erroneamente como divisores de seção.
    const isSecaoModulo = shapesComTexto <= LIMIAR_SHAPES_SECAO &&
      (textoSlide.includes("MÓDULO") || textoSlide.includes("MODULE") || textoSlide.includes("MODULO"));

    // Detecta seção por "cara": 1 caixa de texto + fonte grande (≥LIMIAR_TITULO_FONTSIZE)
    // Slides de seção minimalistas não têm keyword "MÓDULO" mas visualmente são
    // divisores: um único texto em display size (≥LIMIAR_TITULO_FONTSIZE). Abaixo é corpo.
    const isSecaoSimples = (() => {
      if (shapesComTexto !== 1) return false;
      try {
        const shapes = slideOrigem.getShapes();
        for (const sh of shapes) {
          try {
            const text = sh.getText();
            if (!text.asString().trim()) continue;
            const runs = text.getRuns();
            for (const run of runs) {
              try {
                const size = run.getTextStyle().getFontSize();
                if (size && size >= LIMIAR_TITULO_FONTSIZE) return true;
              } catch(e) {}
            }
          } catch(e) {}
        }
      } catch(e) {}
      return false;
    })();

    const isSecao = isSecaoModulo || isSecaoSimples;

    // ── Ordem de prioridade na classificação ──────────────────────────────────
    // 1. "Obrigado/Thank You" → sempre substitui pelo slide fixo de encerramento
    // 2. "Junte-se"/seção    → slide de seção do template + conteúdo copiado
    // 3. Layout explícito (LAYOUTS_CAPA, LAYOUTS_SECAO)
    // 4. Fundo gradient      → seção visual (mesmo sem keyword de módulo)
    // 5. Fundo claro         → slide white
    // 6. Padrão              → slide escuro
    // ─────────────────────────────────────────────────────────────────────────

    if (isObrigado) {
      // Slide "Obrigado" é sempre substituído pelo template fixo — não há conteúdo
      // variável para copiar. O original vai oculto como referência de auditoria.
      try {
        const slideRef = destino.appendSlide(slideOrigem);
        slideRef.setSkipped(true);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} ref: ${e.message}`); }

      try {
        destino.appendSlide(slideThanks);
        Logger.log(`🙏 Slide ${i+1}: Thank You`);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} thanks: ${e.message}`); }

    } else if (isComunidade || isSecao) {
      // Slide de comunidade ("Junte-se") recebe o mesmo tratamento visual de seção —
      // fundo escuro com curvas — pois estruturalmente funciona como divisor de bloco.
      try {
        const slideRef = destino.appendSlide(slideOrigem);
        slideRef.setSkipped(true);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} ref: ${e.message}`); }

      try {
        const novoSlide = destino.appendSlide(slideSecao);
        copiarConteudo(slideOrigem, novoSlide, COR_TEXTO_ESCURO);
        const tipo = isComunidade ? 'comunidade' : (isSecaoSimples ? 'seção (simples)' : 'seção (módulo)');
        Logger.log(`${isComunidade ? '🤝' : '📑'} Slide ${i+1}: ${tipo}`);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1}: ${e.message}`); }

    } else {
      try {
        const slideRef = destino.appendSlide(slideOrigem);
        slideRef.setSkipped(true);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} ref: ${e.message}`); }

      try {
        const layoutNome = slideOrigem.getLayout().getLayoutName();
        let templateBase, corTexto;

        if (LAYOUTS_CAPA.includes(layoutNome)) {
          templateBase = slideCapa;
          corTexto     = COR_TEXTO_ESCURO;
          Logger.log(`   → Capa`);
        } else if (LAYOUTS_SECAO.includes(layoutNome)) {
          templateBase = slideSecao;
          corTexto     = COR_TEXTO_ESCURO;
          Logger.log(`   → Seção (layout)`);
        } else if (slideEhGradient(slideOrigem)) {
          // Gradiente linear é assinatura visual de slides de seção no template antigo —
          // não existe equivalente no novo template, então mapeia para o slide de seção.
          templateBase = slideSecao;
          corTexto     = COR_TEXTO_ESCURO;
          Logger.log(`   → Seção (gradient)`);
        } else if (slideEhClaro(slideOrigem)) {
          templateBase = slideWhite;
          corTexto     = COR_TEXTO_CLARO;
          Logger.log(`   → White (claro)`);
        } else {
          templateBase = slidePadrao;
          corTexto     = COR_TEXTO_ESCURO;
          Logger.log(`   → Padrão (escuro)`);
        }

        const novoSlide = destino.appendSlide(templateBase);
        copiarConteudo(slideOrigem, novoSlide, corTexto);
      } catch(e) { Logger.log(`⚠️ Slide ${i+1} novo: ${e.message}`); }
    }

    // Throttle entre slides para não estourar a cota de chamadas da Slides API
    // ("Service invoked too many times"). Cada slide pode gerar dezenas de chamadas
    // (appendSlide + insertTextBox + N calls de estilo).
    Utilities.sleep(THROTTLE_POR_SLIDE_MS);
  }

  if (endIndex < total) {
    props.setProperty('job_currentIndex', String(endIndex));
    Logger.log(`⏸️ Batch ${currentIndex + 1}-${endIndex} done. Faltam ${total - endIndex} slides.`);
    agendarProximoBatch();
  } else {
    // ── Fim do job: remove os slides do molde ──
    // Os slides do template foram usados como "origem" de appendSlide, então ainda
    // existem no arquivo destino. Precisam ser removidos após o processamento completo
    // pois o appendSlide copia, não move.
    Logger.log("🧹 Removendo slides do template...");
    Utilities.sleep(DELAY_APOS_MAKECOPY_MS);

    const destinoFinal = SlidesApp.openById(destinoId);
    const slidesFinais = destinoFinal.getSlides();
    let removidos = 0;

    slidesFinais.forEach(s => {
      try {
        if (templateSlideIds.includes(s.getObjectId())) {
          s.remove();
          removidos++;
          // Throttle: remover muitos slides em sequência rápida pode gerar erros de quota.
          Utilities.sleep(THROTTLE_AO_REMOVER_MS);
        }
      } catch(e) {
        Logger.log(`⚠️ Erro ao remover: ${e.message}`);
        // Espera mais longa no erro para dar tempo à API de se recuperar.
        Utilities.sleep(DELAY_ERRO_REMOCAO_MS);
      }
    });

    props.deleteProperty('job_destinoId');
    props.deleteProperty('job_origemId');
    props.deleteProperty('job_templateIds');
    props.deleteProperty('job_currentIndex');
    props.deleteProperty('job_total');

    Logger.log(`🧹 ${removidos} slides do template removidos`);
    Logger.log("✅ Criado: " + DriveApp.getFileById(destinoId).getUrl());
  }
}


function agendarProximoBatch() {
  // Remove triggers pendentes antes de criar um novo — evita acúmulo de triggers
  // duplicados caso transformarComTemplate seja chamado várias vezes ou o trigger
  // anterior ainda não tenha disparado.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processarBatch') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // DELAY_PROXIMO_BATCH_S é suficiente para a API "respirar" entre batches sem atrasar demais
  // o processamento total. O novo contexto de execução zera o contador de tempo.
  ScriptApp.newTrigger('processarBatch')
    .timeBased()
    .after(DELAY_PROXIMO_BATCH_S * 1000)
    .create();

  Logger.log(`⏰ Próximo batch agendado em ~${DELAY_PROXIMO_BATCH_S}s`);
}


// Chamada manual para cancelar um job travado ou iniciado por engano.
// Limpa tanto o estado quanto os triggers pendentes — sem isso, processarBatch
// continuaria sendo invocado mesmo após o arquivo destino ser deletado.
function limparJob() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('job_destinoId');
  props.deleteProperty('job_origemId');
  props.deleteProperty('job_templateIds');
  props.deleteProperty('job_currentIndex');
  props.deleteProperty('job_total');

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processarBatch') {
      ScriptApp.deleteTrigger(t);
    }
  });

  Logger.log("🧹 Job cancelado e triggers removidos.");
}


// Retorna true se a imagem deve ser omitida do arquivo destino.
// Estratégia de dois níveis: título primeiro (mais preciso), dimensão como fallback
// para imagens sem metadado. Dimensão só bloqueia quando não há título para evitar
// falsos positivos em conteúdo legítimo com tamanhos coincidentes.
function deveBloquear(image) {
  const title = image.getTitle() || "";
  const w = Math.round(image.getWidth());
  const h = Math.round(image.getHeight());

  if (IMAGENS_BLOQUEADAS_TITULO.some(t => title.includes(t))) return true;

  if (!title) {
    return IMAGENS_BLOQUEADAS_DIMENSOES.some(d => d.w === w && d.h === h);
  }

  return false;
}


// Determina se o fundo do slide é claro o suficiente para usar texto escuro.
// A API retorna a cor do fundo como RGB direto ou como referência a uma cor do tema —
// cada caso precisa de tratamento diferente porque asRgbColor() falha em cores de tema.
function slideEhClaro(slide) {
  try {
    const bg = slide.getBackground();
    if (bg.getType().toString() !== "SOLID") return false;

    const color = bg.getSolidFill().getColor();
    const colorType = color.getColorType().toString();

    if (colorType === "RGB") {
      const hex = color.asRgbColor().asHexString().toUpperCase();
      const r = parseInt(hex.slice(1, 3), 16);
      const g = parseInt(hex.slice(3, 5), 16);
      const b = parseInt(hex.slice(5, 7), 16);
      // Limiar alto (LIMIAR_CLARO_RGB) para garantir que só brancos e off-whites muito claros
      // entrem — cinzas médios ficam no template escuro.
      return r > LIMIAR_CLARO_RGB && g > LIMIAR_CLARO_RGB && b > LIMIAR_CLARO_RGB;
    }

    if (colorType === "THEME") {
      try {
        const themeColor = color.asThemeColor();
        const themeType = themeColor.getThemeColorType().toString();
        if (themeType === "LIGHT1" || themeType === "BACKGROUND1" || themeType === "LIGHT2") return true;
        // Atenção: neste tema customizado DARK1 está mapeado para um tom claro (creme/off-white),
        // ao contrário do comportamento padrão do GSlides onde DARK1 é escuro.
        if (themeType === "DARK1") return true;
        if (themeType === "DARK2") return false;
      } catch(e) {}

      // Quando o tipo do tema não é conclusivo, infere a cor de fundo pela cor do texto:
      // texto escuro → fundo claro.
      return textoEhEscuro(slide);
    }

    return false;
  } catch (e) {
    return false;
  }
}


// Heurística de último recurso: se a maioria do texto é escuro, o fundo provavelmente é claro.
// Usada quando a cor do fundo não pode ser determinada diretamente (cor de tema não resolvível).
// Exige mínimo de 2 runs para evitar decisões baseadas em texto único não representativo.
function textoEhEscuro(slide) {
  try {
    const shapes = slide.getShapes();
    let runsChecados = 0;
    let runsEscuros = 0;

    for (const shape of shapes) {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) continue;

        const runs = text.getRuns();
        for (const run of runs) {
          try {
            const t = run.asString();
            if (!t.trim()) continue;

            const style = run.getTextStyle();
            const fg = style.getForegroundColor();
            if (!fg) continue;

            const colorType = fg.getColorType().toString();

            if (colorType === "RGB") {
              const hex = fg.asRgbColor().asHexString().toUpperCase();
              const r = parseInt(hex.slice(1, 3), 16);
              const g = parseInt(hex.slice(3, 5), 16);
              const b = parseInt(hex.slice(5, 7), 16);
              runsChecados++;
              if (r < LIMIAR_ESCURO_RGB && g < LIMIAR_ESCURO_RGB && b < LIMIAR_ESCURO_RGB) runsEscuros++;
            } else if (colorType === "THEME") {
              try {
                const themeType = fg.asThemeColor().getThemeColorType().toString();
                runsChecados++;
                if (themeType === "DARK1" || themeType === "TEXT1" || themeType === "DARK2") {
                  runsEscuros++;
                }
              } catch(e) {}
            }
          } catch(e) {}
        }
      } catch(e) {}
    }

    // PERCENTUAL_RUNS_ESCUROS é o limiar: suficientemente alto para ignorar runs isolados
    // com cor herdada/indefinida, mas não tão alto a ponto de falhar em slides
    // com mistura de texto colorido (destaques, links).
    if (runsChecados >= 2 && runsEscuros >= Math.ceil(runsChecados * PERCENTUAL_RUNS_ESCUROS)) {
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}


// Gradiente linear é assinatura dos slides de seção no template antigo.
// Não há como replicar o gradiente no novo template, então esses slides
// sempre recebem o slide de seção padrão.
function slideEhGradient(slide) {
  try {
    return slide.getBackground().getType().toString() === "LINEAR_GRADIENT";
  } catch (e) {
    return false;
  }
}


// Copia texto, formatação, imagens e anotações do slide de origem para o destino.
// O destino já existe (foi criado por appendSlide a partir do template) — esta função
// apenas sobrepõe conteúdo, não altera o fundo nem o layout do template.
function copiarConteudo(slideOrigem, slideDestino, corTexto = "#FFFFFF") {

  // ── Shapes com texto ──
  slideOrigem.getShapes().forEach(shape => {
    try {
      const texto = shape.getText().asString();
      if (!texto.trim()) return;

      // insertTextBox não herda o padding interno do shape original, causando clipping
      // de texto nas bordas. O PADDING expande a caixa para compensar.
      const novoShape = slideDestino.insertTextBox(
        texto,
        shape.getLeft()   - PADDING,
        shape.getTop()    - PADDING,
        shape.getWidth()  + PADDING * 2,
        shape.getHeight() + PADDING * 2
      );

      try {
        novoShape.setContentAlignment(shape.getContentAlignment());
      } catch(e) {
        // ContentAlignment pode não estar disponível em shapes simples; TOP é o padrão seguro.
        novoShape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
      }

      const srcText = shape.getText();
      const dstText = novoShape.getText();
      const srcParas = srcText.getParagraphs();
      const dstParas = dstText.getParagraphs();

      const linhasComTexto = srcParas.filter(p => p.getRange().asString().trim()).length;
      const isMultiParagraph = linhasComTexto > 1;

      // Usado para distinguir títulos multi-linha (fonte grande) de listas/corpo (fonte pequena).
      const maxFontSize = (() => {
        let max = 0;
        srcText.getRuns().forEach(run => {
          try {
            const s = run.getTextStyle().getFontSize();
            if (s && s > max) max = s;
          } catch(e) {}
        });
        return max;
      })();

      // Força FONTE_CORPO_LISTA apenas em caixas de lista/corpo: multi-linha E sem texto grande.
      // Caixas com fonte acima de LIMIAR_CORPO_FONTSIZE são títulos multi-linha (ex: título + subtítulo)
      // e devem manter seus tamanhos originais.
      const forcarCorpo = isMultiParagraph && maxFontSize > 0 && maxFontSize <= LIMIAR_CORPO_FONTSIZE;

      // Define o font size baseline do shape inteiro antes de trabalhar run a run,
      // pois alguns runs podem não ter tamanho explícito e herdariam o padrão do shape.
      try {
        let baseSize = srcText.getTextStyle().getFontSize();
        if (forcarCorpo) baseSize = FONTE_CORPO_LISTA;
        if (baseSize) dstText.getTextStyle().setFontSize(baseSize);
      } catch(e) {}

      // O cursor rastreia a posição no texto de destino — necessário porque
      // getRange() opera por índice de caractere, não por run.
      let cursor = 0;
      const dstTextLength = dstText.asString().length;
      srcText.getRuns().forEach(run => {
        const t = run.asString();
        if (!t) return;

        // Guard: o texto de destino pode ser ligeiramente menor que o de origem
        // se houver caracteres especiais não suportados. Evita IndexOutOfBounds.
        const endCursor = Math.min(cursor + t.length, dstTextLength);
        if (cursor >= dstTextLength) return;

        const srcStyle = run.getTextStyle();
        const dstStyle = dstText.getRange(cursor, endCursor).getTextStyle();

        try { dstStyle.setFontFamily(FONTE_NOVA); }        catch(e) {}
        try { dstStyle.setBold(srcStyle.isBold()); }       catch(e) {}
        try { dstStyle.setItalic(srcStyle.isItalic()); }   catch(e) {}
        try { dstStyle.setUnderline(srcStyle.isUnderline()); } catch(e) {}
        // Cor original é descartada — usa sempre a cor do template para garantir
        // contraste correto (branco em fundo escuro, escuro em fundo claro).
        try { dstStyle.setForegroundColor(corTexto); }     catch(e) {}

        try {
          let s = srcStyle.getFontSize();
          if (forcarCorpo) s = FONTE_CORPO_LISTA;
          if (s) dstStyle.setFontSize(s);
        } catch(e) {}

        cursor += t.length;
      });

      // ParagraphStyle opera na unidade de parágrafo, não de run — feito separado do loop de runs.
      for (let i = 0; i < srcParas.length; i++) {
        const srcPara = srcParas[i];
        const dstPara = dstParas[i];
        if (!dstPara) continue;

        try {
          const srcPS = srcPara.getRange().getParagraphStyle();
          const dstPS = dstPara.getRange().getParagraphStyle();

          try {
            const align = srcPS.getAlignment();
            dstPS.setAlignment(align || SlidesApp.ParagraphAlignment.START);
          } catch(e) {}
          try { dstPS.setIndentStart(srcPS.getIndentStart()); }         catch(e) {}
          try { dstPS.setIndentFirstLine(srcPS.getIndentFirstLine()); } catch(e) {}
          try { dstPS.setIndentEnd(srcPS.getIndentEnd()); }             catch(e) {}
          try { dstPS.setSpaceAbove(srcPS.getSpaceAbove()); }           catch(e) {}
          try { dstPS.setSpaceBelow(srcPS.getSpaceBelow()); }           catch(e) {}
          try {
            const ls = srcPS.getLineSpacing();
            // Cap em LINE_SPACING_MAX: line spacing acima disso estoura a caixa do template,
            // que é mais compacta que o template original. LINE_SPACING_DEFAULT é o padrão seguro.
            dstPS.setLineSpacing(ls ? Math.min(ls, LINE_SPACING_MAX) : LINE_SPACING_DEFAULT);
          } catch(e) {}

          // Font size por parágrafo precisa ser aplicado depois do run-a-run
          // porque setFontSize no range de parágrafo pode sobrescrever variações
          // intra-parágrafo — só faz sentido quando forçamos corpo uniforme.
          try {
            if (forcarCorpo) {
              dstPara.getRange().getTextStyle().setFontSize(FONTE_CORPO_LISTA);
            } else {
              let paraSize = srcPara.getRange().getTextStyle().getFontSize();
              if (paraSize) dstPara.getRange().getTextStyle().setFontSize(paraSize);
            }
          } catch(e) {}

        } catch(e) {}
      }

      // Preserva fill sólido de shapes que têm fundo colorido (ex: caixas de destaque).
      // Shapes sem fill ficam transparentes para não cobrir o fundo do template.
      try {
        const fill = shape.getFill();
        if (fill.getType() === SlidesApp.FillType.SOLID)
          novoShape.getFill().setSolidFill(fill.getSolidFill().getColor().asRgbColor().asHexString());
        else
          novoShape.getFill().setTransparent();
      } catch(e) {}

      // Remove borda se o original não a tinha — insertTextBox cria borda padrão por default.
      try {
        if (!shape.getBorder().isVisible()) novoShape.getBorder().setTransparent();
      } catch(e) {}

    } catch(e) { Logger.log(`⚠️ Shape: ${e.message}`); }
  });


  // ── Imagens ──
  // Exporta como PNG para garantir compatibilidade — o formato original pode ser
  // EMF/WMF (do PowerPoint) que o Slides não renderiza corretamente.
  slideOrigem.getImages().forEach(image => {
    try {
      if (deveBloquear(image)) {
        Logger.log(`🚫 Bloqueada: "${image.getTitle()}" ${Math.round(image.getWidth())}x${Math.round(image.getHeight())}`);
        return;
      }
      slideDestino.insertImage(
        image.getAs('image/png'),
        image.getLeft(), image.getTop(),
        image.getWidth(), image.getHeight()
      );
    } catch(e) { Logger.log(`⚠️ Imagem: ${e.message}`); }
  });


  // ── Anotações do apresentador ──
  try {
    const notes = slideOrigem.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
    if (notes) slideDestino.getNotesPage().getSpeakerNotesShape().getText().setText(notes);
  } catch(e) { Logger.log(`⚠️ Notes: ${e.message}`); }

}


// ─── FUNÇÕES AUXILIARES ──────────────────────────────────
// Rode estas funções no arquivo de ORIGEM antes de configurar as constantes
// para entender a distribuição de tipos de slide e quais imagens bloquear.

// Agrupa slides por tipo de fundo e mostra contagem + lista de slides por grupo.
// Útil para calibrar LAYOUTS_WHITE e identificar se há tipos de fundo inesperados.
function estatisticasBackground() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const grupos = {};

  slides.forEach((slide, i) => {
    let chave;
    try {
      const bg = slide.getBackground();
      const tipo = bg.getType().toString();

      if (tipo === "SOLID") {
        try {
          const hex = bg.getSolidFill().getColor().asRgbColor().asHexString().toUpperCase();
          chave = `RGB ${hex}`;
        } catch (e) {
          try {
            const themeType = bg.getSolidFill().getColor().getColorType().toString();
            chave = `Theme ${themeType}`;
          } catch (e2) {
            chave = `Cor não-resolvível`;
          }
        }
      } else {
        chave = `Tipo: ${tipo}`;
      }
    } catch (e) {
      chave = `Erro: ${e.message}`;
    }

    if (!grupos[chave]) grupos[chave] = 0;
    grupos[chave]++;
  });

  Logger.log(`📊 ESTATÍSTICAS DE BACKGROUND (${slides.length} slides)\n`);
  Object.entries(grupos)
    .sort((a, b) => b[1] - a[1])
    .forEach(([chave, count]) => {
      Logger.log(`${count}x | ${chave}`);
    });
}


// Lista todas as imagens do arquivo com frequência de aparição.
// Imagens que aparecem em 3+ slides são candidatas a IMAGENS_BLOQUEADAS
// (elementos do template antigo que se repetem como decoração).
function identificarImagensTemplate() {
  const slides   = SlidesApp.getActivePresentation().getSlides();
  const contagem = {};
  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;
    slide.getImages().forEach(img => {
      const title = img.getTitle()       || "(sem título)";
      const desc  = img.getDescription() || "(sem desc)";
      const w     = Math.round(img.getWidth());
      const h     = Math.round(img.getHeight());
      const chave = `title:"${title}" | desc:"${desc}" | ${w}x${h}`;
      if (!contagem[chave]) contagem[chave] = { count: 0, primeiroSlide: i + 1 };
      contagem[chave].count++;
    });
  });
  Object.entries(contagem)
    .sort((a, b) => b[1] - a[1])
    .forEach(([chave, info]) => {
      Logger.log(`${info.count >= 3 ? "🚩" : "  "} ${info.count}x | ${chave} | 1º slide: ${info.primeiroSlide}`);
    });
}

// Confirma que origem e template têm as mesmas dimensões de página.
// Diferença de dimensões causaria desalinhamento de todos os elementos copiados.
function verificarDimensoes() {
  const origem  = SlidesApp.getActivePresentation();
  const template = SlidesApp.openById(TEMPLATE_ID);
  Logger.log(`Origem:  ${origem.getPageWidth()}pt x ${origem.getPageHeight()}pt`);
  Logger.log(`Destino: ${template.getPageWidth()}pt x ${template.getPageHeight()}pt`);
}

// Lista os nomes de layout do template para atualizar LAYOUTS_CAPA / LAYOUTS_SECAO / LAYOUTS_WHITE.
function listarLayoutsTemplate() {
  SlidesApp.openById(TEMPLATE_ID).getLayouts().forEach((l, i) => {
    Logger.log(`${i}. ${l.getLayoutName()}`);
  });
}

// Lista os slides do template com prévia de texto para conferir se IDX_* estão corretos.
function listarSlidesTemplate() {
  SlidesApp.openById(TEMPLATE_ID).getSlides().forEach((s, i) => {
    const texto = s.getShapes().map(sh => {
      try { return sh.getText().asString().trim(); } catch(e) { return ""; }
    }).filter(t => t).join(" | ").substring(0, 60);
    Logger.log(`${i}. ${texto || "(sem texto)"}`);
  });
}


// ─── DIAGNÓSTICO: slideEhClaro ───────────────────────────
// Mostra slide a slide o que o slideEhClaro retorna e por quê.
// Omite slides com fundo escuro óbvio (THEME não-RGB) para não poluir o log.
function debugWhiteSlides() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  let countWhite = 0, countGradient = 0, countEscuro = 0;

  Logger.log("🔍 DIAGNÓSTICO slideEhClaro / slideEhGradient\n");

  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;

    let bgType = "?", colorType = "?", hex = "?", ehClaro = false, ehGrad = false;
    let layoutNome = "?";
    let motivo = "";

    try { layoutNome = slide.getLayout().getLayoutName(); } catch(e) {}

    try {
      const bg = slide.getBackground();
      bgType = bg.getType().toString();

      if (bgType === "SOLID") {
        const color = bg.getSolidFill().getColor();
        colorType = color.getColorType().toString();

        if (colorType === "RGB") {
          hex = color.asRgbColor().asHexString().toUpperCase();
          const r = parseInt(hex.slice(1, 3), 16);
          const g = parseInt(hex.slice(3, 5), 16);
          const b = parseInt(hex.slice(5, 7), 16);
          ehClaro = r > LIMIAR_CLARO_RGB && g > LIMIAR_CLARO_RGB && b > LIMIAR_CLARO_RGB;
          motivo = `RGB ${hex} | R=${r} G=${g} B=${b} | claro=${ehClaro}`;
        } else {
          motivo = `ColorType=${colorType} (não é RGB, assume escuro)`;
        }
      } else if (bgType === "LINEAR_GRADIENT") {
        ehGrad = true;
        motivo = "LINEAR_GRADIENT → seção";
      } else {
        motivo = `Tipo=${bgType} → escuro`;
      }
    } catch(e) {
      motivo = `Erro: ${e.message}`;
    }

    let templateEscolhido;
    if (LAYOUTS_CAPA.includes(layoutNome)) {
      templateEscolhido = "CAPA";
    } else if (LAYOUTS_SECAO.includes(layoutNome)) {
      templateEscolhido = "SECAO (layout)";
    } else if (ehGrad) {
      templateEscolhido = "SECAO (gradient)";
      countGradient++;
    } else if (ehClaro) {
      templateEscolhido = "WHITE ✅";
      countWhite++;
    } else {
      templateEscolhido = "PADRAO (escuro)";
      countEscuro++;
    }

    if (ehClaro || ehGrad || bgType !== "SOLID" || colorType === "RGB") {
      Logger.log(`Slide ${i+1} | layout=${layoutNome} | bg=${bgType} | ${motivo} | → ${templateEscolhido}`);
    }
  });

  Logger.log(`\n📊 RESUMO:`);
  Logger.log(`  WHITE: ${countWhite}`);
  Logger.log(`  GRADIENT (seção): ${countGradient}`);
  Logger.log(`  ESCURO (padrão): ${countEscuro}`);
}


// ─── DIAGNÓSTICO: Imagens bloqueadas ─────────────────────
// Mostra quais imagens estão sendo bloqueadas em cada slide
function debugImagensBloqueadas() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  let totalBloqueadas = 0, totalCopiadas = 0;

  Logger.log("🔍 DIAGNÓSTICO de imagens bloqueadas vs copiadas\n");

  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;

    const imagens = slide.getImages();
    if (imagens.length === 0) return;

    imagens.forEach(img => {
      const title = img.getTitle() || "(sem título)";
      const w = Math.round(img.getWidth());
      const h = Math.round(img.getHeight());
      const bloqueada = deveBloquear(img);

      let motivo = "";
      if (bloqueada) {
        totalBloqueadas++;
        const matchTitulo = IMAGENS_BLOQUEADAS_TITULO.find(t => title.includes(t));
        const matchDim = IMAGENS_BLOQUEADAS_DIMENSOES.find(d => d.w === w && d.h === h);
        if (matchTitulo) motivo = `TÍTULO match: "${matchTitulo}"`;
        else if (matchDim) motivo = `DIMENSÃO match: ${matchDim.w}x${matchDim.h}`;
        else motivo = "???";
        Logger.log(`🚫 Slide ${i+1} | ${w}x${h} | "${title}" | ${motivo}`);
      } else {
        totalCopiadas++;
      }
    });
  });

  Logger.log(`\n📊 RESUMO:`);
  Logger.log(`  Bloqueadas: ${totalBloqueadas}`);
  Logger.log(`  Copiadas: ${totalCopiadas}`);
  Logger.log(`\n💡 Procure no log acima por imagens que NÃO deveriam estar bloqueadas.`);
  Logger.log(`   Cada linha 🚫 mostra: slide | dimensão | título | motivo do bloqueio`);
}


// ─── DIAGNÓSTICO: Slide específico ───────────────────────
// Mostra TODOS os detalhes de um slide: background, layout, theme color
// Uso: debugSlideEspecifico(6) para investigar o slide 6
function debugSlideEspecifico(num) {
  const slide = SlidesApp.getActivePresentation().getSlides()[num - 1];
  if (!slide) {
    Logger.log(`❌ Slide ${num} não existe`);
    return;
  }

  Logger.log(`🔍 DETALHES DO SLIDE ${num}\n`);

  try {
    Logger.log(`Layout: ${slide.getLayout().getLayoutName()}`);
  } catch(e) {
    Logger.log(`Layout: Erro - ${e.message}`);
  }

  Logger.log(`Skipped: ${slide.isSkipped()}`);

  try {
    const bg = slide.getBackground();
    const bgType = bg.getType().toString();
    Logger.log(`\nBackground type: ${bgType}`);

    if (bgType === "SOLID") {
      const color = bg.getSolidFill().getColor();
      const colorType = color.getColorType().toString();
      Logger.log(`Color type: ${colorType}`);

      if (colorType === "RGB") {
        const hex = color.asRgbColor().asHexString().toUpperCase();
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        Logger.log(`Hex: ${hex} | R=${r} G=${g} B=${b}`);
        Logger.log(`É claro (RGB)? ${r > LIMIAR_CLARO_RGB && g > LIMIAR_CLARO_RGB && b > LIMIAR_CLARO_RGB}`);
      } else if (colorType === "THEME") {
        Logger.log(`\nTentando resolver theme color...`);
        try {
          const themeColor = color.asThemeColor();
          const themeType = themeColor.getThemeColorType().toString();
          Logger.log(`Theme color type: ${themeType}`);
          Logger.log(`É claro (THEME)? ${themeType === "LIGHT1" || themeType === "BACKGROUND1"}`);
        } catch(e) {
          Logger.log(`Erro ao resolver theme: ${e.message}`);
        }

        try {
          Logger.log(`\nTentando asRgbColor() no theme...`);
          const rgb = color.asRgbColor().asHexString();
          Logger.log(`RGB forçado: ${rgb}`);
        } catch(e) {
          Logger.log(`asRgbColor falhou: ${e.message}`);
        }
      }
    } else if (bgType === "LINEAR_GRADIENT") {
      Logger.log(`→ Gradiente (seção)`);
    }
  } catch(e) {
    Logger.log(`Erro no background: ${e.message}`);
  }

  Logger.log(`\nslideEhClaro retorna: ${slideEhClaro(slide)}`);
  Logger.log(`slideEhGradient retorna: ${slideEhGradient(slide)}`);

  Logger.log(`\nShapes com texto:`);
  slide.getShapes().forEach((sh, i) => {
    try {
      const t = sh.getText().asString().trim();
      if (!t) return;
      Logger.log(`  ${i}: "${t.substring(0, 50)}"`);

      sh.getText().getRuns().forEach((run, ri) => {
        try {
          const rt = run.asString().trim();
          if (!rt) return;
          const style = run.getTextStyle();
          const fg = style.getForegroundColor();
          if (!fg) {
            Logger.log(`      run ${ri}: (sem cor explícita)`);
            return;
          }
          const ct = fg.getColorType().toString();
          if (ct === "RGB") {
            const hex = fg.asRgbColor().asHexString().toUpperCase();
            const r = parseInt(hex.slice(1, 3), 16);
            const g = parseInt(hex.slice(3, 5), 16);
            const b = parseInt(hex.slice(5, 7), 16);
            const escuro = r < LIMIAR_ESCURO_RGB && g < LIMIAR_ESCURO_RGB && b < LIMIAR_ESCURO_RGB;
            Logger.log(`      run ${ri}: RGB ${hex} | escuro=${escuro}`);
          } else if (ct === "THEME") {
            try {
              const tt = fg.asThemeColor().getThemeColorType().toString();
              const escuro = tt === "DARK1" || tt === "TEXT1" || tt === "DARK2";
              Logger.log(`      run ${ri}: THEME ${tt} | escuro=${escuro}`);
            } catch(e) {
              Logger.log(`      run ${ri}: THEME (não resolvível)`);
            }
          }
        } catch(e) {}
      });
    } catch(e) {}
  });

  Logger.log(`\nImagens:`);
  slide.getImages().forEach((img, i) => {
    const w = Math.round(img.getWidth());
    const h = Math.round(img.getHeight());
    const title = img.getTitle() || "(sem título)";
    Logger.log(`  ${i}: ${w}x${h} | "${title}" | bloqueada=${deveBloquear(img)}`);
  });
}


// ─── DIAGNÓSTICO: Slides BLANK_8 (white vs black) ───────
// Mostra todos os slides com layout BLANK_8 e suas características
// para distinguir quais são white e quais são black
function debugBlank8() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  let count = 0;

  Logger.log("🔍 DIAGNÓSTICO: Slides BLANK_8\n");
  Logger.log("Legenda: BG=background | TXT=cor do 1º run de texto\n");

  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;

    let layoutNome = "?";
    try { layoutNome = slide.getLayout().getLayoutName(); } catch(e) {}
    if (layoutNome !== "BLANK_8") return;

    count++;
    let bgType = "?", bgColorType = "?", bgHex = "?", bgTheme = "?";
    let txtColorType = "?", txtHex = "?", txtTheme = "?";
    let ehGradient = false;
    let primeiroTexto = "";

    try {
      const bg = slide.getBackground();
      bgType = bg.getType().toString();
      if (bgType === "SOLID") {
        const color = bg.getSolidFill().getColor();
        bgColorType = color.getColorType().toString();
        if (bgColorType === "RGB") {
          bgHex = color.asRgbColor().asHexString().toUpperCase();
        } else if (bgColorType === "THEME") {
          try { bgTheme = color.asThemeColor().getThemeColorType().toString(); } catch(e) {}
        }
      } else if (bgType === "LINEAR_GRADIENT") {
        ehGradient = true;
      }
    } catch(e) {}

    try {
      const shapes = slide.getShapes();
      for (const sh of shapes) {
        try {
          const text = sh.getText();
          const str = text.asString().trim();
          if (!str) continue;
          if (!primeiroTexto) primeiroTexto = str.substring(0, 40);

          const runs = text.getRuns();
          for (const run of runs) {
            try {
              const t = run.asString().trim();
              if (!t) continue;
              const fg = run.getTextStyle().getForegroundColor();
              if (!fg) continue;
              txtColorType = fg.getColorType().toString();
              if (txtColorType === "RGB") {
                txtHex = fg.asRgbColor().asHexString().toUpperCase();
              } else if (txtColorType === "THEME") {
                try { txtTheme = fg.asThemeColor().getThemeColorType().toString(); } catch(e) {}
              }
              break;
            } catch(e) {}
          }
          if (txtColorType !== "?") break;
        } catch(e) {}
      }
    } catch(e) {}

    let classificacao;
    if (ehGradient) {
      classificacao = "SEÇÃO (gradient)";
    } else if (LAYOUTS_WHITE.includes("BLANK_8")) {
      classificacao = "WHITE (layout) ← ATUAL";
    } else {
      classificacao = "PADRÃO (escuro)";
    }

    Logger.log(
      `Slide ${i+1} | BG: ${bgType}${bgType === "SOLID" ? `/${bgColorType}` : ""}` +
      `${bgColorType === "RGB" ? ` ${bgHex}` : ""}` +
      `${bgColorType === "THEME" ? ` ${bgTheme}` : ""}` +
      `${ehGradient ? " (GRADIENT)" : ""}` +
      ` | TXT: ${txtColorType}` +
      `${txtColorType === "RGB" ? ` ${txtHex}` : ""}` +
      `${txtColorType === "THEME" ? ` ${txtTheme}` : ""}` +
      ` | "${primeiroTexto}"` +
      ` | → ${classificacao}`
    );
  });

  Logger.log(`\n📊 Total BLANK_8: ${count}`);
  Logger.log(`\n💡 Procure o padrão: quais slides deveriam ser BLACK mas estão indo para WHITE?`);
  Logger.log(`   Compare BG e TXT dos slides corretos vs incorretos.`);
}
