// ═══════════════════════════════════════════════════════════════════════════════
//  PPT AUTOMATION — Troca de fonte para apresentação ativa
// ═══════════════════════════════════════════════════════════════════════════════
//  O script percorre todos os slides visíveis (ocultos são ignorados) e substitui
//  todas as fontes para a fonte especificada em FONTE_NOVA, preservando estilos
//  (bold, italic, underline) de cada run de texto.
//
//  Útil para padronizar fontes após mudanças no template ou corrigir inconsistências.
//
// ───────────────────────────────────────────────────────────────────────────────
//  COMO USAR
// ───────────────────────────────────────────────────────────────────────────────
//
//  Abra a apresentação no GSlides e execute no editor de scripts:
//
//    trocarFonte()               — substitui todas as fontes para FONTE_NOVA
//    diagnosticarFontes()        — lista as fontes atualmente em uso
//
// ═══════════════════════════════════════════════════════════════════════════════

// ─── CONFIGURAÇÃO ───────────────────────────────────────
const FONTE_NOVA = "Inter";
// ─────────────────────────────────────────────────────────


function trocarFonte() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  let total = 0;
  let trocadas = 0;

  Logger.log(`🔤 Alterando fontes para "${FONTE_NOVA}"...\n`);

  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;

    slide.getShapes().forEach(shape => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;

        text.getRuns().forEach(run => {
          try {
            if (!run.asString().trim()) return;
            const fonteAtual = run.getTextStyle().getFontFamily();
            total++;
            if (fonteAtual && fonteAtual !== FONTE_NOVA) {
              run.getTextStyle().setFontFamily(FONTE_NOVA);
              trocadas++;
            }
          } catch(e) {}
        });
      } catch(e) { Logger.log(`⚠️ Slide ${i + 1}: ${e.message}`); }
    });
  });

  Logger.log(`✅ Pronto!`);
  Logger.log(`📊 ${trocadas} runs trocados de ${total} processados`);
}


function diagnosticarFontes() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const contagem = {};

  Logger.log("🔍 DIAGNÓSTICO DE FONTES\n");

  slides.forEach((slide, i) => {
    if (slide.isSkipped()) return;

    slide.getShapes().forEach((shape, si) => {
      try {
        const text = shape.getText();
        if (!text.asString().trim()) return;

        text.getRuns().forEach((run, ri) => {
          try {
            if (!run.asString().trim()) return;
            const f = run.getTextStyle().getFontFamily() || "(sem fonte)";
            if (!contagem[f]) contagem[f] = 0;
            contagem[f]++;
          } catch(e) {}
        });
      } catch(e) {}
    });
  });

  Logger.log("Distribuição de fontes:");
  Object.entries(contagem)
    .sort((a, b) => b[1] - a[1])
    .forEach(([fonte, count]) => {
      Logger.log(`  ${count}x | ${fonte}`);
    });
}