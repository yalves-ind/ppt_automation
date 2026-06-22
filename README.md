# PPT Automation — Migração de Template para Google Slides

Scripts de Google Apps Script para migrar apresentações do Google Slides de um template antigo para um novo, preservando conteúdo, notas e imagens e aplicando automaticamente a identidade visual do novo template.

---

## Visão Geral

O fluxo completo tem três etapas:

1. **`process_batch.js`** — Lê a apresentação de origem, classifica cada slide por tipo e recria todos os slides no novo template. Roda em lotes agendados para contornar o limite de execução de ~6 minutos do Apps Script.
2. **`validacoes.js`** — Inspeciona o arquivo gerado e sinaliza inconsistências: fonte errada, cor de texto incorreta, imagens bloqueadas que vazaram, notas perdidas, espaçamento fora do cap, bordas indevidas etc.
3. **`troca_fonte.js`** — Utilitário independente para padronizar a fonte de qualquer apresentação aberta. Útil como correção pontual sem rodar a migração completa.

---

## Pré-requisitos

- Conta Google com acesso ao Google Drive e ao Google Slides que será migrado
- Acesso de edição ao template de destino
- Google Apps Script vinculado à apresentação de **origem** (não ao template)

---

## Instalação

1. Abra a apresentação de **origem** no Google Slides
2. Acesse **Extensões > Apps Script**
3. Crie três arquivos `.gs` e cole o conteúdo de cada script:
   - `process_batch.gs` ← conteúdo de `process_batch.js`
   - `validacoes.gs` ← conteúdo de `validacoes.js`
   - `troca_fonte.gs` ← conteúdo de `troca_fonte.js`
4. Configure as constantes em `process_batch.gs` e em `VAL_CONFIG` dentro de `validacoes.gs` (ver seção [Configuração](#configuração))
5. Salve o projeto

---

## Configuração

> **Ordem obrigatória:** preencha `TEMPLATE_ID` → rode as funções de mapeamento (Passos 1–3 do [Fluxo de Uso](#fluxo-de-uso)) → ajuste os índices e listas de bloqueio → só então execute `transformarComTemplate()`. Rodar a transformação sem passar pelo mapeamento é o principal motivo de resultados inesperados.

### `process_batch.js` — Constantes principais

| Constante | O que é | Padrão |
|---|---|---|
| `TEMPLATE_ID` | ID do arquivo de template no Drive (parte da URL entre `/d/` e `/edit`) | `" "` ← **obrigatório preencher** |
| `FONTE_NOVA` | Fonte aplicada a todos os runs de texto | `"Inter"` |
| `IDX_CAPA` | Índice (0-based) do slide de capa no template | `1` |
| `IDX_SECAO` | Índice do slide de seção no template | `8` |
| `IDX_PADRAO` | Índice do slide padrão escuro | `9` |
| `IDX_WHITE` | Índice do slide claro | `5` |
| `IDX_THANKS` | Índice do slide "Thank You" | `10` |
| `LAYOUTS_CAPA` | Nomes de layout que mapeiam para capa | `["BLANK_1_2"]` |
| `LAYOUTS_SECAO` | Nomes de layout que mapeiam para seção | `["TITLE_5", "TITLE_10"]` |
| `LAYOUTS_WHITE` | Nomes de layout que podem ser claros ou escuros (verificação visual) | `["BLANK_8"]` |
| `COR_TEXTO_ESCURO` | Cor do texto em slides escuros | `"#FFFFFF"` |
| `COR_TEXTO_CLARO` | Cor do texto em slides claros | `"#1c1d2f"` |
| `IMAGENS_BLOQUEADAS_TITULO` | Títulos de imagem que não devem migrar (logos do template antigo). **Preencha com os títulos retornados por `identificarImagensTemplate()`** — candidatos são os marcados com 🚩 (3+ slides) | `[]` por padrão |
| `IMAGENS_BLOQUEADAS_DIMENSOES` | Dimensões (px) de imagens sem título que não devem migrar — fallback para assets sem metadado. Preencha com os tamanhos identificados pelo mesmo mapeamento | ver arquivo |
| `PADDING` | Compensação de padding para `insertTextBox` (pt) | `7.2` |
| `FONTE_CORPO_LISTA` | Font size forçado em caixas multi-parágrafo de corpo/lista | `14` |
| `BATCH_SIZE` | Slides processados por execução | `50` |

#### Limiares de classificação

| Constante | Descrição | Padrão |
|---|---|---|
| `LIMIAR_CLARO_RGB` | Canal RGB mínimo para fundo ser considerado "claro" | `240` |
| `LIMIAR_ESCURO_RGB` | Canal RGB máximo para texto ser considerado "escuro" | `128` |
| `LIMIAR_TITULO_FONTSIZE` | Font size mínimo para slide ser classificado como seção simples (pt) | `40` |
| `LIMIAR_CORPO_FONTSIZE` | Font size máximo para caixas multi-parágrafo serem tratadas como corpo | `22` |
| `PERCENTUAL_RUNS_ESCUROS` | Percentual mínimo de runs escuros para inferir fundo claro | `0.6` |
| `LIMIAR_SHAPES_SECAO` | Máximo de shapes com texto para o slide ser candidato a seção | `3` |
| `LINE_SPACING_MAX` | Cap de line spacing aplicado na cópia | `130` |
| `LINE_SPACING_DEFAULT` | Line spacing padrão quando o original não define | `115` |

#### Delays e throttles

| Constante | Descrição | Padrão |
|---|---|---|
| `DELAY_APOS_MAKECOPY_MS` | Espera após `makeCopy` (eventual consistency do Drive) | `3000` ms |
| `THROTTLE_POR_SLIDE_MS` | Throttle entre slides (evita quota da Slides API) | `800` ms |
| `THROTTLE_AO_REMOVER_MS` | Throttle ao remover slides do molde | `500` ms |
| `DELAY_ERRO_REMOCAO_MS` | Espera extra em caso de erro na remoção | `2000` ms |
| `DELAY_PROXIMO_BATCH_S` | Intervalo entre batches agendados | `15` s |

> **`validacoes.js`** espelha as mesmas constantes dentro do objeto `VAL_CONFIG`. Mantenha os valores sincronizados.

---

## Fluxo de Uso

> **Ponto de partida:** antes de qualquer coisa, preencha `TEMPLATE_ID` em `process_batch.gs` (e o campo equivalente em `VAL_CONFIG.TEMPLATE_ID` em `validacoes.gs`) com o ID do seu template. Sem isso, as funções de mapeamento não conseguem abrir o template para comparação.

### Pré-migração (leitura, sem alterações)

Execute estas funções na ordem abaixo a partir da apresentação de **origem**. Nenhuma delas altera arquivos.

**Passo 1 — Verificar dimensões**
```
verificarDimensoes()
```
Confirma que origem e template têm as mesmas dimensões de página (ex: 960×540 pt). Diferenças causam desalinhamento de todos os elementos copiados.

**Passo 2 — Identificar imagens do template antigo**
```
identificarImagensTemplate()
```
Lista todas as imagens com frequência de aparição. Imagens marcadas com 🚩 (3+ slides) são candidatas a logos/ícones do template antigo. Revise e atualize `IMAGENS_BLOQUEADAS_TITULO` ou `IMAGENS_BLOQUEADAS_DIMENSOES` conforme necessário.

**Passo 3 — Estatísticas de background**
```
estatisticasBackground()
```
Mostra a distribuição de tipos de fundo (RGB, THEME, GRADIENT) por contagem. Útil para calibrar `LAYOUTS_WHITE` e identificar cores que possam ser mal classificadas.

> Se encontrar fundos ambíguos, use `debugWhiteSlides()` ou `debugBlank8()` para investigar slide a slide antes de prosseguir.

### Migração

**Passo 4 — Iniciar o job**
```
transformarComTemplate()
```
Cria o arquivo de destino baseado no template e inicia o processamento. O primeiro batch roda imediatamente; os seguintes são agendados automaticamente a cada ~15 s.

- Não feche o editor de scripts durante a execução.
- Para cancelar um job em andamento: `limparJob()`
- Acompanhe o progresso em **Execuções > Ver execuções do projeto**

Ao final, o log exibe a URL do arquivo criado.

### Pós-migração (validação)

Abra o arquivo **destino** (`FIX_ID_*`) no editor de scripts e execute:

```
validacaoGeral()
```

Roda todas as validações e emite um veredito consolidado. Para rodar validações individualmente, veja a seção [Referência de Validações](#referência-de-validações).

Se as notas do apresentador estiverem faltando nos slides visíveis:
```
corrige_notas(true)   // dry-run: mostra o que faria
corrige_notas()       // aplica as correções
```

Para remover os slides ocultos de referência após validação:
```
corrige_slides_ocultos(true)  // dry-run
corrige_slides_ocultos()      // aplica
```

---

## Como os Slides São Classificados

O `process_batch.js` usa a seguinte árvore de prioridade para mapear cada slide da origem para o molde correto do template:

```
1. Texto contém "OBRIGADO" / "THANK YOU"  → slide Thanks (fixo, sem cópia de conteúdo)
2. Texto contém "JUNTE-SE"                → slide Seção
3. ≤ LIMIAR_SHAPES_SECAO shapes + keyword → slide Seção (seção de módulo)
4. 1 shape + font size ≥ LIMIAR_TITULO    → slide Seção (seção simples)
5. Layout em LAYOUTS_CAPA                 → slide Capa
6. Layout em LAYOUTS_SECAO               → slide Seção (por layout)
7. Fundo LINEAR_GRADIENT                 → slide Seção (herança visual)
8. Fundo claro (RGB > LIMIAR_CLARO_RGB)  → slide White
9. Padrão                                → slide Padrão (escuro)
```

Cada slide da origem é inserido como **oculto** logo antes do slide transformado, servindo de referência de auditoria. O arquivo destino fica com o par `[oculto original | visível transformado]` para cada slide.

---

## Referência de Diagnóstico (`process_batch.js`)

| Função | Descrição |
|---|---|
| `verificarDimensoes()` | Compara dimensões de página origem vs template |
| `identificarImagensTemplate()` | Lista imagens por frequência; 🚩 = candidata a bloqueio |
| `estatisticasBackground()` | Distribuição de tipos de fundo por contagem |
| `listarLayoutsTemplate()` | Lista nomes de layout do template (para configurar `LAYOUTS_*`) |
| `listarSlidesTemplate()` | Lista slides do template com prévia de texto (para conferir `IDX_*`) |
| `debugWhiteSlides()` | Mostra classificação visual de cada slide (white / escuro / gradient) |
| `debugBlank8()` | Foca em slides `BLANK_8` para distinguir white vs black |
| `debugImagensBloqueadas()` | Mostra quais imagens seriam bloqueadas e por quê |
| `debugSlideEspecifico(N)` | Dump completo do slide N: fundo, layout, runs, imagens |

---

## Referência de Validações (`validacoes.js`)

### Modos de operação

O `validacoes.js` detecta automaticamente o estado do arquivo destino:

| Modo | Descrição |
|---|---|
| `com_ref` | Refs ocultas intactas — todas as validações rodam em modo completo |
| `sem_ref` | Refs ocultas já removidas — 4 validações puladas, 4 rodam em fallback |
| `parcial` | Algumas refs faltam — validações ignoram apenas os pares quebrados |

### Validações disponíveis

| Função | O que verifica | `sem_ref` |
|---|---|---|
| `valida_dimensoes()` | Destino e template têm mesmas dimensões | roda |
| `valida_slides()` | Estrutura de pares `[oculto, visível]` | pulada |
| `valida_fonte()` | Todos os runs usam `FONTE_NOVA` | roda |
| `valida_cores_texto()` | Cores são apenas `COR_TEXTO_ESCURO` ou `COR_TEXTO_CLARO` | fallback |
| `valida_imagens()` | Nenhuma imagem bloqueada vazou; sem duplicações | fallback |
| `valida_classificacao()` | Slide foi mapeado para o molde correto | fallback |
| `valida_corpo_lista()` | Caixas multi-parágrafo estão em `FONTE_CORPO_LISTA` | roda |
| `valida_molde_removido()` | Slides do molde foram removidos do destino | roda |
| `valida_notas()` | Notas do apresentador foram preservadas | fallback |
| `valida_espacamento()` | Line spacing ≤ `LINE_SPACING_MAX` | roda |
| `valida_bordas()` | Shapes sem borda no original não têm borda no destino | roda |

### Correções disponíveis

| Função | O que faz | Destrutiva? |
|---|---|---|
| `corrige_notas(dryRun)` | Copia notas do oculto para o visível quando o visível está vazio | Sim — use `dryRun=true` primeiro |
| `corrige_slides_ocultos(dryRun)` | Remove todos os slides marcados como ocultos | Sim — use `dryRun=true` primeiro |

---

## Utilitário de Fonte (`troca_fonte.js`)

Script independente para padronizar fonte em qualquer apresentação, sem precisar rodar a migração completa.

| Função | Descrição |
|---|---|
| `trocarFonte()` | Substitui todas as fontes para `FONTE_NOVA` em slides visíveis |
| `diagnosticarFontes()` | Lista fontes em uso e contagem de runs por fonte |

Configure `FONTE_NOVA` no topo do arquivo antes de rodar.

---

## Limitações Conhecidas

- **Heurísticas calibradas**: a classificação de slides foi calibrada para um conjunto específico de apresentações. Apresentações com layouts muito distintos podem precisar de ajuste nas constantes e limiares antes de rodar.
- **Sem refs, sem base de comparação**: com as refs ocultas removidas, é impossível detectar conteúdo perdido na cópia ou classificações erradas que ainda resultem em texto legível.
- **Gradientes não replicados**: fundos com `LINEAR_GRADIENT` não têm equivalente no novo template e são mapeados para o slide de seção padrão.
- **Cores originais descartadas**: todas as cores de texto são substituídas por `COR_TEXTO_ESCURO` ou `COR_TEXTO_CLARO` para garantir contraste. Destaques coloridos no original são perdidos.
- **Apresentações grandes**: acima de ~150 slides, `validacaoGeral()` pode estourar o limite de execução de 6 min do Apps Script. Nesse caso, rode as funções `valida_*` individualmente.

---

## Estrutura de Arquivos

```
.
├── process_batch.js   # Migração principal — classificação e cópia de conteúdo
├── validacoes.js      # Validação pós-migração do arquivo gerado
└── troca_fonte.js     # Utilitário independente de padronização de fonte
```
