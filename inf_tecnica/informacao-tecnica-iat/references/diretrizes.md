# Diretrizes de domínio — Informação Técnica IAT/ERLON

Consolidado das decisões de Leonardo ao longo dos casos do projeto. Em conflito, vale a regra mais recente.

## Mapa de placeholders do modelo

| Placeholder | Origem | Observação |
|---|---|---|
| `[[nº_protocolo]]` | Requerimento | formato `XX.XXX.XXX-X` |
| `[[razão_social]]` | Requerimento | grafia como no Requerimento (geralmente maiúsculas) |
| `[[atvidades_específicas]]` | Requerimento | linha ATIVIDADE da tabela (sic: "atvidades") |
| `[[município]]`, `[[UF]]` | Requerimento | ex.: Londrina, PR |
| `[[modalidade]]` | Requerimento | por extenso + sigla (RLO, LAS, RLAS, LP, LO-A…) |
| `[[atividade_principal]]` | Requerimento | atividade principal descrita |
| `[[descrição_produtos]]` | Requerimento/PCA | produtos/insumos/volumes |
| `[[endereço]]` | Requerimento | endereço completo com CEP |
| `[[área_const]]` | Requerimento | área construída em m² |
| `[[área_ampliação]]` | Requerimento | só em ampliações (LO-A) |
| `[[nº_funcionários]]` | Requerimento | número de funcionários |
| `[[porte]]` | Requerimento | Pequeno / Médio / Grande |
| `[[data_protocolo]]` | Requerimento | dd/mm/aaaa |
| `[[legislações_aplicáveis]]` | norma | IN IAT 65/2025 (padrão); + CEMA 70/2009 ou 107/2020 |
| `[[licença_anterior]]` / `[[nº_licença_anterior]]` / `[[nº_protocolo_licença_anterior]]` | licença anexa | identificar sempre que houver licença vigente |
| `[[docs_analisados]]` | Análise Documental | itens atendidos → lista romana sem bullet |
| `[[matérias_primas]]` | Requerimento/PCA | reproduzir quantidades exatas |
| `[[tipo_abastecimento_água]]` / `[[descrição_abastecimento_água]]` | Requerimento | rede pública, poço… |
| `[[nº_outorga_agua]]` | Portaria | número(s) de outorga de água |
| `[[descrição_efluentes]]` / `[[tratamento_efluentes]]` / `[[destinação_efluentes]]` / `[[vazão_efluentes]]` | Requerimento/PCA | efluentes |
| `[[nº_outorga_efluentes]]` / `[[data_validade_portaria_efluentes]]` | Portaria | outorga de efluentes |
| `[[descrição_emissoes_atm]]` | Requerimento/PCA | emissões; "Nada consta" se não houver |
| `[[lista_residuos_destinadores]]` | PGRS/Requerimento | descrição + destinador, **sem código IBAMA** |
| `[[ano_inventário_resíduos]]` / `[[nº_inventário_resíduos]]` | SGA | inventário de resíduos |
| `[[relatório_vistoria_*]]` | vistoria | **sempre em branco** (visita ocorre depois) |
| `[[descrição_fluxograma]]` | MCE/PCA | em branco se não houver material claro |

## Regras de conteúdo

- **Requerimento é fonte primária**; demais documentos complementam; conflito resolve a favor do Requerimento.
- **Nunca** incluir código IBAMA em resíduos — só descrição.
- **CONCLUSÕES em branco** salvo pedido explícito. Quando pedido: ~50 palavras, favorável, recomendar **deferimento**.
- Lista de documentos em **numeral romano, um por parágrafo, sem bullet**.
- Informação não localizada → **placeholder intacto** (não apagar, não inventar).
- Identificar **licença anterior e seu protocolo** sempre que os documentos permitirem.
- **Divergências entre fontes** (ex.: endereços) → **sinalizar para Leo**, não resolver sozinho.
- Blocos condicionais não aplicáveis → **remover marcador e conteúdo**.
- Parágrafos de vistoria → **manter**, placeholder em branco (decisão atual; antes alguns casos os removiam — a opção vigente é manter).

## Formatação

- Cabeçalho: `___ /2026`. Título: `INFORMAÇÃO TÉCNICA ___/2026 – IAT ERLON`.
- Tabela de identificação: PROTOCOLO, INTERESSADO, ATIVIDADE, MUNICÍPIO.
- Títulos de seção em CAIXA ALTA, negrito, sublinhado (estilo do modelo).
- Assinatura: manter o bloco do modelo (Leonardo Parreira / Engenheiro Ambiental e Sanitarista / Residente – IAT/ERLON).
- **Sempre** editar o DOCX original; preservar cabeçalho, rodapé e imagens.
- Saída: `INFORMAÇÃO TÉCNICA - {EMPRESA}.docx`.

## Casos atípicos (pontuais, fora do fluxo padrão)

- Renomear MATÉRIA PRIMA → PRODUTO ARMAZENADO (empreendimentos de serviço/armazenagem).
- Fundir ÁGUA UTILIZADA + EFLUENTES LÍQUIDOS em "ÁGUA E EFLUENTES" quando os textos forem curtos.
- Seção ALTERAÇÃO DE RAZÃO SOCIAL após DESCRIÇÃO DA ATIVIDADE (parágrafo breve, ≤150 palavras).
- LO-A: incluir seção OBRA DE AMPLIAÇÃO (passo manual pós-geração).

## Quirks técnicos do modelo

- A **data** ("Londrina, … de 2026.") é campo dinâmico do Word: renderiza a data atual. O `data_documento` do JSON é best-effort.
- Os placeholders de vistoria têm nomes inconsistentes entre seções (ver SKILL.md).
- Arquivos do processo às vezes são ZIP/texto com extensão `.pdf`.
- Auditar placeholders remanescentes após o pack: `grep -o '\[\[[^]]*\]\]' word/document.xml | sort -u`.
