---
name: informacao-tecnica-iat
description: Monta uma Informação Técnica (IT) no formato institucional do IAT/ERLON a partir do Requerimento de licenciamento e dos PDFs do processo, gerando um .docx parcialmente preenchido sobre um modelo embutido. Use sempre que Leonardo pedir para "montar/elaborar/produzir a informação técnica desse processo", "preencher a IT", "fazer a IT de apoio", "gerar a Informação Técnica para os arquivos enviados", ou quando ele entregar o Requerimento (com PCA/PGRS/PAE/portarias de outorga/Análise Documental) sinalizando que quer a IT. Acionar mesmo se o pedido vier curto ou implícito — qualquer combinação de "informação técnica" / "IT" + processo de licenciamento do IAT é gatilho válido. Não confundir com a Análise Documental (skill analise-documental-iat): a IT é o documento técnico de apoio, com seções DESCRIÇÃO DA ATIVIDADE, ÁGUA UTILIZADA, EFLUENTES LÍQUIDOS, EMISSÕES, RESÍDUOS SÓLIDOS e CONCLUSÕES.
---

# Informação Técnica — IAT/ERLON

## Visão geral

Produz uma **Informação Técnica parcialmente preenchida (.docx)** que Leonardo termina manualmente. A skill **edita o template embutido** (`assets/INFORMAÇÃO_TÉCNICA_MODELO.docx`) preenchendo os `[[placeholders]]` com os dados extraídos dos documentos do processo. É uma **IT de apoio**: os relatos de vistoria ainda não ocorreram, então ficam em branco.

**Regra de ouro — fonte primária:** o **Requerimento é sempre a fonte primária**. PCA, PGRS/PGRSS, PAE, memoriais, portarias de outorga, licença anterior e a Análise Documental são **complementos**. Em qualquer conflito, prevalece o Requerimento.

**Princípio de não-surpresa:** preencher o máximo que os documentos permitirem; **tudo que não for localizado fica com o placeholder original intacto** — não apagar, não substituir por espaço em branco, não inventar. Não recriar o documento do zero: **editar sempre o .docx original**, preservando cabeçalho, rodapé, imagens e formatação.

## Inputs esperados

| Input | Obrigatório | Papel |
|---|---|---|
| Requerimento (PDF) | Sim | **Fonte primária** — formulário oficial do SGA preenchido pelo empreendedor |
| Análise Documental (a única .docx no processo) | Não, mas usual | Origem da **lista de documentos analisados** (itens atendidos) |
| PCA / MCE / memorial | Não | Caracterização do processo produtivo, fluxograma, matérias-primas |
| PGRS / PGRSS ou resumo | Não | Resíduos sólidos e destinadores |
| Portarias de Outorga | Não | Outorgas de água/efluentes |
| Licença anterior | Não | Tipo, número e protocolo da licença vigente (para renovações) |

Os arquivos do processo frequentemente são **ZIP com extensão `.pdf`** (usar `zipfile`) ou **texto puro com extensão `.pdf`**. A Análise Documental é a **única `.docx`** do processo — não analisá-la, apenas extrair dela a lista de documentos.

Se o Requerimento estiver ausente, **pedir antes de prosseguir**. Diante de divergências entre fontes (ex.: endereço diferente no Requerimento e no PCA), **sinalizar para Leonardo decidir** — não resolver em silêncio.

## Workflow

1. **Ler o Requerimento** (fonte primária) e extrair identificação + caracterização.
2. **Complementar** com PCA/PGRS/PAE/portarias/licença anterior; extrair a **lista de documentos** da Análise Documental (itens atendidos).
3. **Montar o `dados.json`** (contrato no docstring de `scripts/fill_template.py`): `valores`, `condicoes` e `docs_analisados`.
4. **Apresentar prévia textual** ao usuário (ver "Checkpoint") para validação.
5. **Gerar o .docx** editando o template embutido (unpack → `fill_template.py` → pack `--original`).
6. **Validar**: converter para PDF (LibreOffice) e renderizar as páginas antes de entregar.
7. **Salvar em `/mnt/user-data/outputs/`** e apresentar via `present_files`.

## Detalhamento por seção

### Tabela de identificação
Quatro linhas: **PROTOCOLO, INTERESSADO, ATIVIDADE, MUNICÍPIO** — placeholders `[[nº_protocolo]]`, `[[razão_social]]`, `[[atvidades_específicas]]`, `[[município]], [[UF]]`. Cabeçalho fica `___ /2026` e o título `INFORMAÇÃO TÉCNICA ___/2026 – IAT ERLON` (placeholder numérico mantido — Leo numera depois).

### DESCRIÇÃO DA ATIVIDADE
Modalidade por extenso + sigla (ex.: `Renovação de Licença de Operação (RLO)`), razão social, atividade principal, descrição de produtos, endereço completo com CEP, área construída, nº de funcionários, porte. Em seguida: data do protocolo e legislação(ões) aplicável(is) (padrão atual: **Instrução Normativa IAT nº 65/2025**; muitos casos combinam com Resolução CEMA nº 70/2009 ou 107/2020). `[[descrição_fluxograma]]` é preenchido a partir do MCE/PCA quando houver material claro; caso contrário fica em branco.

### Documentos analisados
Lista em **numeral romano minúsculo** (`i.`, `ii.`, `iii.` …), **um item por parágrafo, sem marcadores (bullets)**. Montada a partir dos itens **atendidos** da Análise Documental. O script já converte o estilo de lista em parágrafo de corpo e aplica a numeração.

### MATÉRIA PRIMA / ÁGUA UTILIZADA / EFLUENTES LÍQUIDOS / EMISSÕES ATMOSFÉRICAS / RESÍDUOS SÓLIDOS
Reproduzir vazões, quantidades, números de outorga e unidades **exatamente** como no Requerimento. Em **RESÍDUOS SÓLIDOS**, usar **apenas a descrição do resíduo — nunca o código IBAMA**.

### CONCLUSÕES e "Relatório fotográfico"
Ficam **em branco** por padrão. Só redigir conclusão **quando Leonardo pedir explicitamente**: ~50 palavras, favorável à concessão, recomendando o **deferimento**.

### Assinatura
**Não tocar.** O template já traz `Leonardo Parreira / Engenheiro Ambiental e Sanitarista / Residente – IAT/ERLON`. Manter como está.

## Blocos condicionais `{(caso ...)}`

O modelo marca trechos opcionais com `{(caso ...)}` + conteúdo entre chaves `{...}`. No `dados.json`, o objeto `condicoes` decide cada um:

| Flag | True (mantém, limpa marcador) | False (remove bloco inteiro) |
|---|---|---|
| `ampliacao` | mantém "que foi aumentada para …" | remove (e a vírgula) |
| `licenca_anterior` | mantém referência à licença vigente | remove |
| `poco` | mantém `[[descrição_abastecimento_água]]` | remove o trecho do poço |
| `portaria_outorga_agua` | mantém a frase + a lista `[[nº_outorga_agua]]` | remove ambas |
| `tratamento_efluentes` | mantém o detalhamento do tratamento | remove |
| `pgrs` | "declarados e apresentados no PGRS" | "declarados," |
| `inventario_residuos` | mantém o parágrafo do inventário | remove o parágrafo |

Regra geral: **bloco que se aplica → remover só o marcador/chaves e manter o texto** (placeholders internos são preenchidos por `valores`); **bloco que não se aplica → remover marcador + conteúdo**.

## Parágrafos de vistoria

Os parágrafos "Durante a vistoria, …" são **mantidos**, apenas limpando o marcador `{(caso haja relatório de vistoria)}` e as chaves; os placeholders `[[relatório_vistoria_*]]` ficam **em branco** para Leo preencher após a visita. O script faz isso automaticamente. **Quirk do modelo:** esses placeholders são reaproveitados de forma inconsistente entre seções (`[[relatório_vistoria_mat_primas]]` aparece em água/efluentes; `[[relatório_vistoria_matérias_primas]]` em emissões). Mantê-los verbatim e mencionar a Leo, se relevante.

## Geração do .docx

Editar o template via unpack → `fill_template.py` → pack:

```bash
# 1. copiar o modelo embutido
cp <skill-dir>/assets/INFORMAÇÃO_TÉCNICA_MODELO.docx /home/claude/saida.docx
# 2. desempacotar
python /mnt/skills/public/docx/scripts/office/unpack.py /home/claude/saida.docx /home/claude/unpacked
# 3. preencher (contrato do JSON no docstring do script)
python <skill-dir>/scripts/fill_template.py /home/claude/unpacked /home/claude/dados.json
# 4. reempacotar preservando relações/imagens/estilos
python /mnt/skills/public/docx/scripts/office/pack.py /home/claude/unpacked \
   "/mnt/user-data/outputs/INFORMAÇÃO TÉCNICA - {EMPRESA}.docx" --original /home/claude/saida.docx
```

O `fill_template.py` faz: resolução dos blocos condicionais, limpeza dos parágrafos de vistoria, expansão da lista de documentos (numeral romano, sem bullet), preenchimento dos `[[placeholders]]` presentes em `valores` (deixando os ausentes intactos), remoção do highlight amarelo de scaffolding, e **auditoria final** dos placeholders remanescentes. Ao final, conferir no PDF que (a) cabeçalho/rodapé/imagens estão íntegros, (b) CONCLUSÕES e Relatório fotográfico estão em branco, (c) nenhum bloco condicional não-aplicável sobrou.

**Substituição segura:** o script substitui placeholders mais longos antes dos mais curtos para evitar corrupção parcial. Smart quotes, quando presentes, ficam como entidades XML (`&#x201C;` / `&#x201D;`) — usar a forma de entidade em substituições manuais.

## Output

- Nome do arquivo: **`INFORMAÇÃO TÉCNICA - {EMPRESA}.docx`** (ex.: `INFORMAÇÃO TÉCNICA - LABORATÓRIO LOGOS.docx`). A razão social como aparece no Requerimento.
- Salvar em `/mnt/user-data/outputs/` e apresentar via `present_files`.

## Checkpoint com o usuário

Antes de gerar, mostrar **prévia textual** de: identificação (4 campos), DESCRIÇÃO DA ATIVIDADE, lista de documentos, e os blocos preenchidos de água/efluentes/emissões/resíduos — destacando o que ficou **em branco** (vistoria, fluxograma, campos não localizados) e qualquer **divergência entre fontes**. Esperar confirmação ou correções. Leonardo trabalha de forma iterativa e directiva, com listas numeradas de correção e fidelidade textual exata ao conteúdo citado.

## Casos atípicos (não automatizar no fluxo geral)

Alterações estruturais feitas para casos específicos **não entram no template geral**: ex.: renomear MATÉRIA PRIMA → PRODUTO ARMAZENADO, fundir ÁGUA + EFLUENTES, adicionar seção ALTERAÇÃO DE RAZÃO SOCIAL. Para **LO-A**, incluir uma seção dedicada **OBRA DE AMPLIAÇÃO** — mas como passo manual após a geração, não como parte do preenchimento padrão. Diante de pedido desse tipo, tratá-lo como ajuste pontual sobre o .docx gerado.

## Referências

- `scripts/fill_template.py` — script de preenchimento; o **contrato completo do `dados.json`** está no docstring.
- `references/diretrizes.md` — regras de domínio consolidadas e mapa de placeholders do modelo.
- `assets/INFORMAÇÃO_TÉCNICA_MODELO.docx` — modelo oficial em branco (com cabeçalho, rodapé e assinatura).
- `/mnt/skills/public/docx/SKILL.md` — referência técnica para geração de .docx.
