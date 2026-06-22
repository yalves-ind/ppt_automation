// ═══════════════════════════════════════════════════════════════════════════════
//  PPT AUTOMATION — Validação pós-transformação (contraprova)
// ═══════════════════════════════════════════════════════════════════════════════
//  Inspeciona o arquivo destino ( já transformado por process_batch.js ) e
//  sinaliza inconsistências contra as regras definidas no process_batch.js:
//  fonte Inter, cores #FFFFFF / #1c1d2f, imagens bloqueadas, classificação
//  de slides, remoção do molde, dimensões, notas, espaçamento, bordas.
//
//  AUTOCONTIDO — não depende de process_batch.js. Todas as constantes e
//  funções auxiliares vivem neste arquivo. O namespace VAL_CONFIG evita
//  conflito de nomes caso ambos .gs estejam no mesmo projeto do Apps Script.
//
// ───────────────────────────────────────────────────────────────────────────────
//  COMO USAR
// ───────────────────────────────────────────────────────────────────────────────
//
//  Abra o arquivo DESTINO no GSlides ( o FIX_ID_* gerado pelo process_batch.js )
//  e execute no editor de scripts:
//
//    validacaoGeral()       — roda todas as validações e emite veredito
//
//  Ou rode qualquer valida_* individualmente para inspectionar um aspecto.
//  As funções ref-dependentes esperam receber o modo; chame _detectarModo()
//  antes, ou simplesmente use validacaoGeral() que orquestra tudo.
//
//  Antes de rodar, ajuste VAL_CONFIG.TEMPLATE_ID no topo do arquivo com o ID
//  do template usado na transformação ( parte da URL após /d/ e antes de /edit ).
//
// ───────────────────────────────────────────────────────────────────────────────
//  MODOS DE OPERAÇÃO ( detectados automaticamente )
// ───────────────────────────────────────────────────────────────────────────────
//
//  com_ref     — refs ocultas intactas: roda TODAS as validações
//  sem_ref     — colega já deletou as refs ocultas: 4 validações puladas,
//                4 rodam em fallback ( versão reduzida cruzando com as regras ),
//                6 independentes rodam iguais
//  parcial     — algumas refs faltam: roda tudo, valida_slides sinaliza os
//                pares quebrados e as dependentes pulam apenas esses pares
//
//  ┌──────────────────────┬──────────┬───────────┬─────────┐
//  │ validação            │ sem_ref  │ parcial   │ com_ref │
//  ├──────────────────────┼──────────┼───────────┼─────────┤
//  │ valida_dimensoes     │ roda     │ roda      │ roda    │
//  │ valida_slides        │ pulada   │ roda      │ roda    │
//  │ valida_fonte         │ roda     │ roda      │ roda    │
//  │ valida_cores_texto   │ fallback │ roda      │ roda    │
//  │ valida_imagens       │ fallback │ roda      │ roda    │
//  │ valida_classificacao │ fallback │ roda      │ roda    │
//  │ valida_corpo_lista   │ roda     │ roda      │ roda    │
//  │ valida_molde_removido│ roda     │ roda      │ roda    │
//  │ valida_notas         │ fallback │ roda      │ roda    │
//  │ valida_espacamento   │ roda     │ roda      │ roda    │
//  │ valida_bordas        │ roda     │ roda      │ roda    │
//  └──────────────────────┴──────────┴───────────┴─────────┘
//
// ───────────────────────────────────────────────────────────────────────────────
//  LIMITAÇÃO RESIDUAL EM sem_ref
// ───────────────────────────────────────────────────────────────────────────────
//
//  Sem as refs ocultas, duas classes de problema são fundamentalmente
//  indetectáveis — qualquer validador que prometa o contrário sem o original
//  está sendo desonesto:
//
//  1. CONTEÚDO PERDIDO     — shapes / parágrafos / imagens que sumiram na cópia.
//                            não há linha de base para comparar.
//  2. CLASSIFICAÇÃO ERRADA-LEGÍVEL — ex: slide de conteúdo virou seção, mas
//                            ainda é legível. o fallback só pega casos onde
//                            a classificação resultou em texto invisível
//                            ( fundo claro com texto claro, ou vice-versa ).
//
//  LEITURA vs ESCRITA:
//    valida_*     — SÓ LEITURA. get/is/as + Logger.log. Seguras em slides
//                  de terceiros sem risco de modificação.
//    corrige_*    — ESCRITA. Modificam o arquivo ativo. Sempre oferecem
//                  dryRun=true para preview sem alterar nada. Rode com
//                  dryRun primeiro para conferir antes de aplicar.
//
//  BATCH ÚNICO: por serem leitura ou escrita leve ( uma chamada por slide ),
//  não há throttle nem triggers agendados. validacaoGeral() roda tudo em
//  uma única execução. Para apresentações muito grandes ( 150+ slides )
//  pode estourar o limite de ~6min do Apps Script — nesse caso, rode as
//  valida_* / corrige_* individualmente.
//
// ═══════════════════════════════════════════════════════════════════════════════


// ─── CONFIGURAÇÃO ────────────────────────────────────────
// Namespace próprio para evitar conflito com constantes globais de
// process_batch.js caso ambos .gs estejam no mesmo projeto do Apps Script.
// Ajuste TEMPLATE_ID antes de rodar — deve ser o mesmo ID usado na
// transformação original.
const VAL_CONFIG = {
  // ID do template ( parte da URL após /d/ e antes de /edit )
  TEMPLATE_ID: " ",

  // Fonte aplicada a todos os runs durante a transformação
  FONTE_NOVA: "Inter",

  // Índices dos slides no template ( 0 = primeiro )
  IDX_CAPA:    1,
  IDX_SECAO:   8,
  IDX_PADRAO:  9,
  IDX_WHITE:   5,
  IDX_THANKS:  10,

  // Nomes de layout do template — mais estáveis que índices
  LAYOUTS_CAPA:  ["BLANK_1_2"],
  LAYOUTS_SECAO: ["TITLE_5", "TITLE_10"],
  LAYOUTS_WHITE: ["BLANK_8"],

  // Cores de texto: branco sobre fundo escuro, escuro sobre fundo claro
  COR_TEXTO_ESCURO: "#FFFFFF",
  COR_TEXTO_CLARO:  "#1c1d2f",

  // Font size forçado em caixas de corpo/lista ( multi-para ≤22pt )
  FONTE_CORPO_LISTA: 14,
  LIMIAR_CORPO_FONTSIZE: 22,

  // Limiares RGB para classificação de cores
  LIMIAR_CLARO_RGB: 240,
  LIMIAR_ESCURO_RGB: 128,

  // Percentual de runs escuros
  PERCENTUAL_RUNS_ESCUROS: 0.6,

  // Line spacing: máximo permitido (cap) e padrão fallback
  LINE_SPACING_MAX: 130,
  LINE_SPACING_DEFAULT: 115,

  // Limiar de pares ok (80%)
  LIMIAR_PARES_OK: 0.8,

  // Logos e ícones do template antigo que não devem migrar.
  // → Mantenha sincronizado com IMAGENS_BLOQUEADAS_TITULO em process_batch.js.
  IMAGENS_BLOQUEADAS_TITULO: [
    "Prancheta 6@2x.png",
    "Prancheta 6@1.5x.png",
    "Prancheta 7 cópia@2x.png",
    "Prancheta 7@1.5x.png",
    "Brand@2x.png",
    "Artboard 1@2x.png",
    "Prancheta 6.png",
  ],

  // Bloqueio por dimensão: fallback para imagens sem título
  IMAGENS_BLOQUEADAS_DIMENSOES: [
    { w: 38,  h: 38  },
    { w: 44,  h: 44  },
    { w: 80,  h: 28  },
  ],
};
// ─────────────────────────────────────────────────────────


// ─── FUNÇÕES AUXILIARES LOCAIS ───────────────────────────
// Cópias read-only das funções equivalentes em process_batch.js, usando
// VAL_CONFIG em vez de globais. Prefixo _ evita colisão se ambos .gs
// estiverem no mesmo projeto.


// Retorna true se a imagem deve ser omitida do arquivo destino.
// Bloqueio por título primeiro ( mais preciso ), dimensão como fallback
// para imagens sem metadado.
function _deveBloquear(image) {
  const title = image.getTitle() || "";
  const w = Math.round(image.getWidth());
  const h = Math.round(image.getHeight());

  if (VAL_CONFIG.IMAGENS_BLOQUEADAS_TITULO.some(t => title.includes(t))) return true;

  if (!title) {
    return VAL_CONFIG.IMAGENS_BLOQUEADAS_DIMENSOES.some(d => d.w === w && d.h === h);
  }

  return false;
}


// Determina se o fundo do slide é claro o suficiente para texto escuro.
// RGB direto: limiar >LIMIAR_CLARO_RGB. THEME: LIGHT1/BACKGROUND1/LIGHT2/DARK1 ( tema
// customizado ) → claro. Caso não conclusivo, infere pela cor do texto.
function _slideEhClaro(slide) {
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
      return r > VAL_CONFIG.LIMIAR_CLARO_RGB && g > VAL_CONFIG.LIMIAR_CLARO_RGB && b > VAL_CONFIG.LIMIAR_CLARO_RGB;
    }

    if (colorType === "THEME") {
      try {
        const themeColor = color.asThemeColor();
        const themeType = themeColor.getThemeColorType().toString();
        if (themeType === "LIGHT1" || themeType === "BACKGROUND1" || themeType === "LIGHT2") return true;
        if (themeType === "DARK1") return true;
        if (themeType === "DARK2") return false;
      } catch(e) {}

      return _textoEhEscuro(slide);
    }

    return false;
  } catch (e) {
    return false;
  }
}


// Heurística de último recurso: se a maioria do texto é escuro, o fundo
// provavelmente é claro. Exige mínimo de 2 runs.
function _textoEhEscuro(slide) {
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
              if (r < VAL_CONFIG.LIMIAR_ESCURO_RGB && g < VAL_CONFIG.LIMIAR_ESCURO_RGB && b < VAL_CONFIG.LIMIAR_ESCURO_RGB) runsEscuros++;
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

    if (runsChecados >= 2 && runsEscuros >= Math.ceil(runsChecados * VAL_CONFIG.PERCENTUAL_RUNS_ESCUROS)) {
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}


// Gradiente linear é assinatura dos slides de seção no template antigo.
function _slideEhGradient(slide) {
  try {
    return slide.getBackground().getType().toString() === "LINEAR_GRADIENT";
  } catch (e) {
    return false;
  }
}


// ─── DETECÇÃO DE MODO ────────────────────────────────────


// Identifica o modo de operação pela presença e padrão das refs ocultas.
// com_ref     ≥LIMIAR_PARES_OK% dos hidden formam par [hidden, visible] com o seguinte
// sem_ref     0 hidden no arquivo
// parcial     entre 0 e LIMIAR_PARES_OK% dos hidden formam par
function _detectarModo() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const total  = slides.length;
  if (total === 0) return 'sem_ref';

  let hidden = 0;
  slides.forEach(s => { if (s.isSkipped()) hidden++; });
  if (hidden === 0) return 'sem_ref';

  const pares    = _identificarPares();
  const paresOk  = pares.filter(p => p.ok).length;
  if (paresOk >= hidden * VAL_CONFIG.LIMIAR_PARES_OK) return 'com_ref';
  return 'parcial';
}


// Percorre os slides reconhecendo o padrão [oculto, visivel] deixado pelo
// process_batch.js. Retorna array de entradas; em cada uma, pelo menos um
// de hidden / visible é nulo quando o par está quebrado.
function _identificarPares() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const pares  = [];
  let i = 0;
  while (i < slides.length) {
    const cur  = slides[i];
    const next = slides[i + 1];
    if (cur.isSkipped()) {
      if (next && !next.isSkipped()) {
        pares.push({
          hidden: cur, visible: next,
          idxHidden: i, idxVisible: i + 1, ok: true
        });
        i += 2;
      } else {
        pares.push({
          hidden: cur, visible: null,
          idxHidden: i, idxVisible: null, ok: false
        });
        i += 1;
      }
    } else {
      pares.push({
        hidden: null, visible: cur,
        idxHidden: null, idxVisible: i, ok: false
      });
      i += 1;
    }
  }
  return pares;
}


// Lista apenas os slides visíveis, anexando a hidden ref quando existir.
// Base para todas as validações que iteram o destino.
function _visiveis() {
  return _identificarPares()
    .filter(p => p.visible)
    .map(p => ({
      slide:  p.visible,
      idx:    p.idxVisible,
      hidden: p.hidden
    }));
}


// Assinatura de texto de um slide: concatenação em maiúsculas dos shapes
// com texto. Usada para detectar moldes não removidos ( conteúdo idêntico
// a um slide do template ) e para distinguir o slide "Thank You" intencional.
function _assinaturaTexto(slide) {
  try {
    return slide.getShapes().map(sh => {
      try { return sh.getText().asString().trim(); } catch(e) { return ""; }
    }).filter(t => t).join(" | ").toUpperCase();
  } catch(e) { return ""; }
}


// Assinatura de background: "TIPO:HEX" para SOLID RGB, "TIPO:THEMETYPE" para
// tema, "TIPO" para os demais. Usada para comparar visible com o molde do
// template ( valida_classificacao e valida_molde_removido ).
function _assinaturaBg(slide) {
  try {
    const bg   = slide.getBackground();
    const tipo = bg.getType().toString();
    if (tipo === "SOLID") {
      try {
        const hex = bg.getSolidFill().getColor().asRgbColor().asHexString().toUpperCase();
        return `${tipo}:${hex}`;
      } catch(e) {
        try {
          const ct = bg.getSolidFill().getColor().getColorType().toString();
          return `${tipo}:${ct}`;
        } catch(e2) { return `${tipo}:?`; }
      }
    }
    return tipo;
  } catch(e) { return "?"; }
}


// Primeira cor de texto RGB encontrada no slide, em HEX maiúsculo.
// Retorna null se não houver cor RGB resolvível. Usada no fallback de
// valida_classificacao para checar coerência fundo×cor.
function _primeiraCorTexto(slide) {
  try {
    for (const shape of slide.getShapes()) {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) continue;
        for (const run of text.getRuns()) {
          try {
            if (!run.asString().trim()) continue;
            const fg = run.getTextStyle().getForegroundColor();
            if (!fg) continue;
            if (fg.getColorType().toString() === "RGB") {
              return fg.asRgbColor().asHexString().toUpperCase();
            }
          } catch(e) {}
        }
      } catch(e) {}
    }
  } catch(e) {}
  return null;
}


// Replica a árvore de classificação do process_batch.js sobre um slide
// ( tipicamente a ref oculta ) para inferir qual molde deveria ter sido
// usado: 'thanks' | 'secao' | 'capa' | 'white' | 'padrao'.
function _classificarHidden(slide) {
  const texto = slide.getShapes().map(s => {
    try { return s.getText().asString(); } catch(e) { return ""; }
  }).join(" ").toUpperCase();

  if (texto.includes("OBRIGADO") || texto.includes("THANK YOU")) return 'thanks';
  if (texto.includes("JUNTE-SE")) return 'secao';

  const shapesComTexto = slide.getShapes().filter(s => {
    try { return s.getText().asString().trim().length > 0; } catch(e) { return false; }
  }).length;

  const isSecaoModulo = shapesComTexto <= VAL_CONFIG.LIMIAR_SHAPES_SECAO && (
    texto.includes("MÓDULO") || texto.includes("MODULE") || texto.includes("MODULO")
  );
  if (isSecaoModulo) return 'secao';

  // isSecaoSimples: 1 shape com texto e fonte ≥LIMIAR_TITULO_FONTSIZE
  let isSecaoSimples = false;
  if (shapesComTexto === 1) {
    for (const sh of slide.getShapes()) {
      try {
        const t = sh.getText();
        if (!t.asString().trim()) continue;
        for (const run of t.getRuns()) {
          try {
            const sz = run.getTextStyle().getFontSize();
            if (sz && sz >= VAL_CONFIG.LIMIAR_TITULO_FONTSIZE) { isSecaoSimples = true; break; }
          } catch(e) {}
        }
        if (isSecaoSimples) break;
      } catch(e) {}
    }
  }
  if (isSecaoSimples) return 'secao';

  let layoutNome = "";
  try { layoutNome = slide.getLayout().getLayoutName(); } catch(e) {}
  if (VAL_CONFIG.LAYOUTS_CAPA.includes(layoutNome))  return 'capa';
  if (VAL_CONFIG.LAYOUTS_SECAO.includes(layoutNome)) return 'secao';
  if (_slideEhGradient(slide))                        return 'secao';
  if (_slideEhClaro(slide))                           return 'white';
  return 'padrao';
}


// Infere qual molde o visible efetivamente usou, comparando sua assinatura
// de background com os slides do TEMPLATE_ID. Retorna 'white' | 'escuro'
// ( capa, secao, padrao, thanks são todos escuros — indistinguíveis por bg )
// ou null se não casar com nenhum molde.
function _classificarVisible(visible) {
  try {
    const tpl         = SlidesApp.openById(VAL_CONFIG.TEMPLATE_ID);
    const tplSlides   = tpl.getSlides();
    const vBg         = _assinaturaBg(visible);
    if (vBg === "?") return null;

    const candidatos = [
      { tipo: 'capa',   idx: VAL_CONFIG.IDX_CAPA   },
      { tipo: 'secao',  idx: VAL_CONFIG.IDX_SECAO  },
      { tipo: 'padrao', idx: VAL_CONFIG.IDX_PADRAO },
      { tipo: 'white',  idx: VAL_CONFIG.IDX_WHITE  },
      { tipo: 'thanks', idx: VAL_CONFIG.IDX_THANKS }
    ];
    for (const c of candidatos) {
      if (c.idx >= tplSlides.length) continue;
      if (_assinaturaBg(tplSlides[c.idx]) === vBg) return c.tipo;
    }
  } catch(e) {}
  return null;
}


// ─── VALIDAÇÕES ──────────────────────────────────────────


// Confirma que destino e template têm as mesmas dimensões de página.
// Diferença aqui significa desalinhamento de todos os elementos copiados.
function valida_dimensoes() {
  try {
    const destino  = SlidesApp.getActivePresentation();
    const template = SlidesApp.openById(VAL_CONFIG.TEMPLATE_ID);
    const dw = destino.getPageWidth(),  dh = destino.getPageHeight();
    const tw = template.getPageWidth(), th = template.getPageHeight();
    if (dw !== tw || dh !== th) {
      Logger.log(`❌ Dimensões divergem: destino ${dw}x${dh}pt | template ${tw}x${th}pt`);
      return 1;
    }
    Logger.log(`✅ Dimensões ok: ${dw}x${dh}pt`);
    return 0;
  } catch(e) {
    Logger.log(`⚠️ valida_dimensoes: TEMPLATE_ID não acessível (${e.message})`);
    return 0;
  }
}


// Estrutura de pares [hidden, visible]. Sinaliza pares quebrados.
// Pulada em sem_ref — não há par para checar.
function valida_slides(modo) {
  if (modo === 'sem_ref') {
    Logger.log("⏭️ valida_slides: pulada (sem ref oculto)");
    return 0;
  }
  const pares   = _identificarPares();
  let ok = 0, erros = 0;
  pares.forEach(p => {
    if (p.ok) { ok++; return; }
    erros++;
    if (p.hidden && !p.visible) {
      Logger.log(`❌ Slide ${p.idxHidden + 1}: ref oculta sem visible seguinte`);
    } else if (!p.hidden && p.visible) {
      Logger.log(`❌ Slide ${p.idxVisible + 1}: visible sem ref oculta precedente`);
    }
  });
  Logger.log(`📊 valida_slides: ${ok} pares ok, ${erros} quebrados`);
  return erros;
}


// Todo run de texto em slides visíveis deveria estar na fonte FONTE_NOVA.
// Independente de refs — checa só o destino.
function valida_fonte() {
  let erros = 0;
  _visiveis().forEach(v => {
    v.slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;
        text.getRuns().forEach((run, ri) => {
          try {
            if (!run.asString().trim()) return;
            const f = run.getTextStyle().getFontFamily();
            if (f && f !== VAL_CONFIG.FONTE_NOVA) {
              erros++;
              Logger.log(`❌ Slide ${v.idx + 1} shape ${si} run ${ri}: fonte="${f}" (esperada "${VAL_CONFIG.FONTE_NOVA}")`);
            }
          } catch(e) {}
        });
      } catch(e) {}
    });
  });
  Logger.log(`📊 valida_fonte: ${erros} runs com fonte incorreta`);
  return erros;
}


// Cores de texto deveriam ser apenas COR_TEXTO_ESCURO ou COR_TEXTO_CLARO.
// com_ref : também checa que a cor escolhida é coerente com o bg do hidden.
// sem_ref : só checa que a cor é uma das duas permitidas ( fallback ).
function valida_cores_texto(modo) {
  const ALLOWED = [
    VAL_CONFIG.COR_TEXTO_ESCURO.toUpperCase(),
    VAL_CONFIG.COR_TEXTO_CLARO.toUpperCase()
  ];
  if (modo === 'sem_ref') {
    Logger.log("⚠️ valida_cores_texto: rodando em fallback (cor residual apenas)");
  }
  let erros = 0;
  _visiveis().forEach(v => {
    const esperada = (modo !== 'sem_ref' && v.hidden)
      ? (_slideEhClaro(v.hidden) ? VAL_CONFIG.COR_TEXTO_CLARO : VAL_CONFIG.COR_TEXTO_ESCURO).toUpperCase()
      : null;
    v.slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;
        text.getRuns().forEach((run, ri) => {
          try {
            if (!run.asString().trim()) return;
            const fg = run.getTextStyle().getForegroundColor();
            if (!fg) return;
            let hex;
            try { hex = fg.asRgbColor().asHexString().toUpperCase(); }
            catch(e) { return; }
            if (!ALLOWED.includes(hex)) {
              erros++;
              Logger.log(`❌ Slide ${v.idx + 1} shape ${si} run ${ri}: cor=${hex} (não permitida)`);
            } else if (esperada && hex !== esperada) {
              erros++;
              Logger.log(`❌ Slide ${v.idx + 1} shape ${si} run ${ri}: cor=${hex} (esperada ${esperada})`);
            }
          } catch(e) {}
        });
      } catch(e) {}
    });
  });
  Logger.log(`📊 valida_cores_texto: ${erros} runs com cor inconsistente`);
  return erros;
}


// Nenhuma imagem bloqueada deveria ter vazado para o destino.
// com_ref : também compara contagem visible vs hidden ( duplicação ).
// sem_ref : só checa bloqueadas vazadas ( fallback ).
function valida_imagens(modo) {
  if (modo === 'sem_ref') {
    Logger.log("⚠️ valida_imagens: rodando em fallback (bloqueadas vazadas apenas)");
  }
  let erros = 0;
  _visiveis().forEach(v => {
    const imgs = v.slide.getImages();
    imgs.forEach(img => {
      const title = img.getTitle() || "";
      const w = Math.round(img.getWidth());
      const h = Math.round(img.getHeight());
      if (_deveBloquear(img)) {
        erros++;
        const motivo = title ? `título "${title}"` : `dimensão ${w}x${h}`;
        Logger.log(`❌ Slide ${v.idx + 1}: imagem bloqueada vazou (${motivo})`);
      }
    });
    if (modo !== 'sem_ref' && v.hidden) {
      const hiddenValidas = v.hidden.getImages().filter(i => !_deveBloquear(i));
      const visibleValidas = imgs.filter(i => !_deveBloquear(i));
      if (visibleValidas.length > hiddenValidas.length) {
        erros++;
        Logger.log(`❌ Slide ${v.idx + 1}: visible tem ${visibleValidas.length} imgs, hidden tinha ${hiddenValidas.length} (possível duplicação)`);
      }
    }
  });
  Logger.log(`📊 valida_imagens: ${erros} inconsistências`);
  return erros;
}


// com_ref : reclassifica o hidden e compara com o molde que o visible usou.
// sem_ref : fallback de auto-consistência — fundo claro deve ter texto
//           COR_TEXTO_CLARO, fundo escuro COR_TEXTO_ESCURO. Pega apenas
//           classificações erradas que resultaram em texto invisível.
function valida_classificacao(modo) {
  if (modo === 'sem_ref') {
    Logger.log("⚠️ valida_classificacao: rodando em fallback (auto-consistência fundo×cor)");
    let erros = 0;
    _visiveis().forEach(v => {
      const claro        = _slideEhClaro(v.slide);
      const corEncontrada = _primeiraCorTexto(v.slide);
      if (!corEncontrada) return;
      const esperada = claro ? VAL_CONFIG.COR_TEXTO_CLARO.toUpperCase() : VAL_CONFIG.COR_TEXTO_ESCURO.toUpperCase();
      if (corEncontrada !== esperada) {
        erros++;
        Logger.log(`❌ Slide ${v.idx + 1}: fundo ${claro ? 'claro' : 'escuro'} mas texto=${corEncontrada} (esperada ${esperada})`);
      }
    });
    Logger.log(`📊 valida_classificacao: ${erros} inconsistências (fallback)`);
    return erros;
  }

  let erros = 0;
  _visiveis().forEach(v => {
    if (!v.hidden) return;
    const esperado = _classificarHidden(v.hidden);
    const usado    = _classificarVisible(v.slide);
    if (usado === null) {
      // não foi possível identificar o molde por assinatura de bg — skip
      return;
    }
    // white ↔ escuro é a distinção mais importante. cap/secao/padrao/thanks
    // são todos escuros e indistinguíveis por bg — aceitamos como compatíveis.
    if (esperado === 'white' && usado !== 'white') {
      erros++;
      Logger.log(`❌ Slide ${v.idx + 1}: hidden classificou white mas visible usou ${usado}`);
    } else if (esperado !== 'white' && usado === 'white') {
      erros++;
      Logger.log(`❌ Slide ${v.idx + 1}: hidden classificou ${esperado} mas visible usou white`);
    } else if (esperado === 'thanks' && usado !== 'thanks') {
      // thanks é distinguível por conteúdo — checagem extra
      const txtVisible = _assinaturaTexto(v.slide);
      if (!txtVisible.includes("THANK YOU") && !txtVisible.includes("OBRIGADO")) {
        erros++;
        Logger.log(`❌ Slide ${v.idx + 1}: hidden era obrigado mas visible não é thanks`);
      }
    }
  });
  Logger.log(`📊 valida_classificacao: ${erros} inconsistências`);
  return erros;
}


// Heurística: caixas multi-parágrafo com maxFontSize ≤ LIMIAR_CORPO_FONTSIZE deveriam estar
// em FONTE_CORPO_LISTA (14pt). Só sinaliza suspeitos — sem o font size
// original não há como confirmar miss com certeza.
function valida_corpo_lista() {
  let erros = 0;
  _visiveis().forEach(v => {
    v.slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;
        const paras = text.getParagraphs();
        const linhasComTexto = paras.filter(p => p.getRange().asString().trim()).length;
        if (linhasComTexto <= 1) return;
        let max = 0;
        text.getRuns().forEach(run => {
          try {
            const sz = run.getTextStyle().getFontSize();
            if (sz && sz > max) max = sz;
          } catch(e) {}
        });
        if (max > 0 && max <= VAL_CONFIG.LIMIAR_CORPO_FONTSIZE) {
          // deveria estar em FONTE_CORPO_LISTA — checa se algum run difere
          let foraDoPadrao = false;
          text.getRuns().forEach(run => {
            try {
              if (!run.asString().trim()) return;
              const sz = run.getTextStyle().getFontSize();
              if (sz && sz !== VAL_CONFIG.FONTE_CORPO_LISTA) { foraDoPadrao = true; }
            } catch(e) {}
          });
          if (foraDoPadrao) {
            erros++;
            Logger.log(`❌ Slide ${v.idx + 1} shape ${si}: corpo/lista (multi-para, max=${max}pt) com runs fora de ${VAL_CONFIG.FONTE_CORPO_LISTA}pt`);
          }
        }
      } catch(e) {}
    });
  });
  Logger.log(`📊 valida_corpo_lista: ${erros} suspeitos`);
  return erros;
}


// com_ref / parcial : conta slides vs esperado ( 2 × pares + singles ).
// sem_ref : compara texto de cada visible com texto dos slides do template —
//           coincidência 1:1 ( exceto thanks intencional ) sugere molde leftover.
function valida_molde_removido(modo) {
  if (modo === 'sem_ref') {
    let tplTexts = [];
    try {
      const tpl = SlidesApp.openById(VAL_CONFIG.TEMPLATE_ID);
      tplTexts = tpl.getSlides().map(s => _assinaturaTexto(s));
    } catch(e) {
      Logger.log(`⚠️ valida_molde_removido: TEMPLATE_ID não acessível (${e.message})`);
      return 0;
    }
    let erros = 0;
    SlidesApp.getActivePresentation().getSlides().forEach((s, i) => {
      if (s.isSkipped()) return;
      const txt = _assinaturaTexto(s);
      if (txt.length < 10) return;
      // thanks é intencionalmente uma cópia do template — não é leftover
      if (txt.includes("THANK YOU") || txt.includes("OBRIGADO")) return;
      if (tplTexts.includes(txt)) {
        erros++;
        Logger.log(`❌ Slide ${i + 1}: conteúdo idêntico a slide do template (possível molde não removido)`);
      }
    });
    Logger.log(`📊 valida_molde_removido: ${erros} suspeitos`);
    return erros;
  }

  const pares    = _identificarPares();
  const paresOk  = pares.filter(p => p.ok).length;
  const singles  = pares.filter(p => !p.ok && p.hidden && !p.visible).length;
  const esperado = paresOk * 2 + singles;
  const atual    = SlidesApp.getActivePresentation().getSlides().length;
  const diff     = atual - esperado;
  if (diff > 0) {
    Logger.log(`❌ ${diff} slides a mais que o esperado (possível molde não removido). Esperado ~${esperado}, atual ${atual}`);
    return diff;
  }
  Logger.log(`✅ valida_molde_removido: ${atual} slides (esperado ~${esperado})`);
  return 0;
}


// com_ref : compara conteúdo das notas par-a-par ( hidden vs visible ).
// sem_ref : fallback de presença — apenas conta visíveis sem notas.
function valida_notas(modo) {
  if (modo === 'sem_ref') {
    Logger.log("⚠️ valida_notas: rodando em fallback (presença apenas)");
    let semNotas = 0;
    _visiveis().forEach(v => {
      try {
        const n = v.slide.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
        if (!n) semNotas++;
      } catch(e) { semNotas++; }
    });
    Logger.log(`📊 valida_notas: ${semNotas} visíveis sem notas (não há como saber se perdeu ou se original também não tinha)`);
    return 0;
  }

  let erros = 0;
  _visiveis().forEach(v => {
    if (!v.hidden) return;
    try {
      const nHidden  = v.hidden.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
      const nVisible = v.slide.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
      if (nHidden && !nVisible) {
        erros++;
        Logger.log(`❌ Slide ${v.idx + 1}: hidden tinha notas mas visible não`);
      } else if (nHidden && nVisible && nHidden !== nVisible) {
        erros++;
        Logger.log(`❌ Slide ${v.idx + 1}: notas divergentes entre hidden e visible`);
      }
    } catch(e) {}
  });
  Logger.log(`📊 valida_notas: ${erros} inconsistências`);
  return erros;
}


// ─── CORREÇÕES ───────────────────────────────────────────
// Funções corrige_* MODIFICAM o arquivo ativo. Sempre oferecem dryRun=true
// para preview sem alterar nada. Rode com dryRun=true primeiro.


// Copia notas do apresentador do hidden ref para o visible quando o visible
// está vazio e o hidden tem conteúdo. NÃO sobrescreve notas existentes no
// visible — apenas preenche lacunas.
//
// Casos:
//   hidden tem, visible não tem  → COPIA do hidden para o visible
//   hidden tem, visible tem      → pula ( visible já tem conteúdo próprio )
//   hidden não tem               → pula ( nada a copiar )
//
// ⚠️ FUNÇÃO MUTADORA: altera o arquivo ativo.
//
// Uso:
//   corrige_notas(true)   — DRY-RUN: apenas loga o que faria, sem modificar
//   corrige_notas()       — aplica as correções de fato
//
// Retorna o número de slides corrigidos ( ou que seriam, em dryRun ).
function corrige_notas(dryRun = false) {
  const modo = _detectarModo();
  if (modo === 'sem_ref') {
    Logger.log("⏭️ corrige_notas: pulada (sem ref oculto — não há fonte para copiar)");
    return 0;
  }

  Logger.log(dryRun
    ? "🔍 corrige_notas: DRY-RUN (preview apenas, sem modificações)"
    : "🔧 corrige_notas: aplicando correções");

  let corrigidos = 0;
  let puladosVisibleTemNotas = 0;
  let puladosHiddenSemNotas = 0;
  let puladosSemHidden = 0;
  let erros = 0;

  _visiveis().forEach(v => {
    if (!v.hidden) {
      puladosSemHidden++;
      return;
    }

    try {
      const nHidden  = v.hidden.getNotesPage().getSpeakerNotesShape().getText().asString().trim();
      const nVisible = v.slide.getNotesPage().getSpeakerNotesShape().getText().asString().trim();

      if (!nHidden) {
        puladosHiddenSemNotas++;
        return;
      }

      if (nVisible) {
        // visible já tem notas — não sobrescrever, preservar conteúdo próprio
        puladosVisibleTemNotas++;
        return;
      }

      // Caso de correção: hidden tem notas, visible não tem
      Logger.log(`📝 Slide ${v.idx + 1}: ${dryRun ? "seria corrigido" : "copiando notas"} (${nHidden.length} chars)`);

      if (!dryRun) {
        try {
          v.slide.getNotesPage().getSpeakerNotesShape().getText().setText(nHidden);
          corrigidos++;
        } catch(e) {
          erros++;
          Logger.log(`❌ Slide ${v.idx + 1}: erro ao copiar — ${e.message}`);
        }
      } else {
        corrigidos++;
      }
    } catch(e) {
      erros++;
      Logger.log(`❌ Slide ${v.idx + 1}: erro ao ler notas — ${e.message}`);
    }
  });

  Logger.log(
    `📊 corrige_notas: ${corrigidos} ${dryRun ? "seriam corrigidos" : "corrigidos"}` +
    ` | ${puladosVisibleTemNotas} já tinham notas` +
    ` | ${puladosHiddenSemNotas} hidden sem notas` +
    ` | ${puladosSemHidden} sem hidden` +
    ` | ${erros} erros`
  );
  return corrigidos;
}


// Remove todos os slides marcados como ocultos ( isSkipped() === true ).
// No fluxo do process_batch.js, esses são os slides originais mantidos
// como referência de auditoria. Depois de validado, podem ser removidos
// para limpar o arquivo final.
//
// ⚠️ FUNÇÃO MUTADORA: altera o arquivo ativo ( remove slides ).
//
// Uso:
//   corrige_slides_ocultos(true)   — DRY-RUN: apenas lista/conta, sem remover
//   corrige_slides_ocultos()       — remove os slides ocultos de fato
//
// Retorna o número de slides removidos ( ou que seriam, em dryRun ).
function corrige_slides_ocultos(dryRun = true) {
  const slides = SlidesApp.getActivePresentation().getSlides();

  // Coleta referências diretas aos slides ocultos ANTES de qualquer remoção.
  // Referência ao objeto slide é estável — não depende de índice, então
  // remover em qualquer ordem é seguro. Índice só é usado para log.
  const ocultos = [];
  slides.forEach((s, i) => {
    if (s.isSkipped()) ocultos.push({ slide: s, idx: i });
  });

  if (ocultos.length === 0) {
    Logger.log("✅ corrige_slides_ocultos: nenhum slide oculto encontrado");
    return 0;
  }

  Logger.log(dryRun
    ? `🔍 corrige_slides_ocultos: DRY-RUN — ${ocultos.length} slides ocultos seriam removidos`
    : `🔧 corrige_slides_ocultos: removendo ${ocultos.length} slides ocultos`);

  // Lista todos no dry-run para inspeção visual
  if (dryRun) {
    ocultos.forEach(o => {
      let previa = "";
      try {
        previa = o.slide.getShapes().map(sh => {
          try { return sh.getText().asString().trim(); } catch(e) { return ""; }
        }).filter(t => t).join(" | ").substring(0, 50);
      } catch(e) {}
      // Logger.log(`   Slide ${o.idx + 1}: ${previa || (sem texto)}`);
       Logger.log(`Slide ${o.idx + 1}: Seria removido`);
    });
    Logger.log(`📊 corrige_slides_ocultos: ${ocultos.length} seriam removidos (dryRun)`);
    return ocultos.length;
  }

  // Aplica: remove de trás pra frente para preservar índices no log
  // caso algo precise referenciar a posição original durante o loop.
  let removidos = 0;
  let erros = 0;
  for (let i = ocultos.length - 1; i >= 0; i--) {
    const o = ocultos[i];
    try {
      o.slide.remove();
      removidos++;
      Logger.log(`🗑️ Slide ${o.idx + 1} removido`);
    } catch(e) {
      erros++;
      Logger.log(`❌ Slide ${o.idx + 1}: erro ao remover — ${e.message}`);
    }
  }

  Logger.log(`📊 corrige_slides_ocultos: ${removidos} removidos | ${erros} erros`);
  return removidos;
}


// Line spacing deveria ser ≤LINE_SPACING_MAX ( cap imposto no process_batch.js ).
function valida_espacamento() {
  let erros = 0;
  _visiveis().forEach(v => {
    v.slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;
        text.getParagraphs().forEach((para, pi) => {
          try {
            if (!para.getRange().asString().trim()) return;
            const ls = para.getRange().getParagraphStyle().getLineSpacing();
            if (ls && ls > VAL_CONFIG.LINE_SPACING_MAX) {
              erros++;
              Logger.log(`❌ Slide ${v.idx + 1} shape ${si} para ${pi}: line spacing=${ls} (cap ${VAL_CONFIG.LINE_SPACING_MAX})`);
            }
          } catch(e) {}
        });
      } catch(e) {}
    });
  });
  Logger.log(`📊 valida_espacamento: ${erros} parágrafos com line spacing > ${VAL_CONFIG.LINE_SPACING_MAX}`);
  return erros;
}


// insertTextBox cria borda por default — o process_batch.js remove quando o
// original não tinha. Sinaliza visíveis com borda aparente em shapes de texto.
function valida_bordas() {
  let erros = 0;
  _visiveis().forEach(v => {
    v.slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;
        const border = shape.getBorder();
        if (border && border.isVisible()) {
          erros++;
          Logger.log(`❌ Slide ${v.idx + 1} shape ${si}: borda aparente (insertTextBox deveria ter removido)`);
        }
      } catch(e) {}
    });
  });
  Logger.log(`📊 valida_bordas: ${erros} shapes com borda aparente`);
  return erros;
}


// ─── ORQUESTRADOR ────────────────────────────────────────


function validacaoGeral() {
  const modo = _detectarModo();
  Logger.log("══════════════════════════════════════════════════════");
  Logger.log(`🔍 Validação pós-transformação | Modo: ${modo}`);
  Logger.log("══════════════════════════════════════════════════════\n");

  const validacoes = [
    ['valida_dimensoes',      () => valida_dimensoes()],
    ['valida_slides',         () => valida_slides(modo)],
    ['valida_fonte',          () => valida_fonte()],
    ['valida_cores_texto',    () => valida_cores_texto(modo)],
    ['valida_imagens',        () => valida_imagens(modo)],
    ['valida_classificacao',  () => valida_classificacao(modo)],
    ['valida_corpo_lista',    () => valida_corpo_lista()],
    ['valida_molde_removido', () => valida_molde_removido(modo)],
    ['valida_notas',          () => valida_notas(modo)],
    ['valida_espacamento',    () => valida_espacamento()],
    ['valida_bordas',         () => valida_bordas()]
  ];

  let totalErros = 0;
  const contagem = {};

  validacoes.forEach(([nome, fn]) => {
    Logger.log(`\n── ${nome} ──`);
    try {
      const n = fn();
      contagem[nome] = n;
      totalErros    += n;
    } catch(e) {
      Logger.log(`⚠️ ${nome}: erro na execução — ${e.message}`);
      contagem[nome] = -1;
    }
  });

  Logger.log("\n══════════════════════════════════════════════════════");
  Logger.log("📊 RESUMO");
  Logger.log("──────────────────────────────────────────────────────");
  Object.entries(contagem).forEach(([nome, n]) => {
    const status = n === 0 ? "✅" : (n < 0 ? "⚠️" : "❌");
    Logger.log(`  ${status} ${nome.padEnd(24)} ${n < 0 ? "erro" : n}`);
  });
  Logger.log("──────────────────────────────────────────────────────");
  Logger.log(`Total: ${totalErros} inconsistência(s) em ${validacoes.length} validações\n`);
  if (totalErros === 0) {
    Logger.log("✅ Arquivo consistente com as regras");
  } else {
    Logger.log(`❌ ${totalErros} inconsistência(s) — revisar log acima`);
  }
}
