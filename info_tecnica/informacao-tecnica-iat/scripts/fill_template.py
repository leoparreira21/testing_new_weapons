#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_template.py — Preenche a Informação Técnica (IAT/ERLON) sobre o modelo embutido.

Opera sobre o `word/document.xml` de um modelo JÁ DESEMPACOTADO
(`assets/INFORMAÇÃO_TÉCNICA_MODELO.docx` → unpack). Faz substituições
byte-a-byte precisas: o modelo tem alguns marcadores irregulares, então a
abordagem é por substituição literal (não por parser genérico), que é a que
se mostrou confiável no histórico de casos do projeto.

USO
----
    python fill_template.py <pasta_unpacked> <dados.json>

ESTRUTURA DO dados.json
-----------------------
{
  "valores": {
      // qualquer [[placeholder]] do modelo -> texto. Os ausentes ficam INTACTOS.
      "nº_protocolo": "25.391.278-2",
      "razão_social": "EMPRESA EXEMPLO LTDA",
      "atvidades_específicas": "Laboratórios clínicos",
      "município": "Londrina", "UF": "PR",
      "modalidade": "Renovação de Licença de Operação (RLO)",
      "atividade_principal": "...", "descrição_produtos": "...",
      "endereço": "Rua X, nº 0, Bairro, Londrina/PR, CEP 00.000-000",
      "área_const": "499,48", "nº_funcionários": "40", "porte": "Pequeno",
      "data_protocolo": "10/05/2025",
      "legislações_aplicáveis": "Resolução CEMA nº 107/2020 e Instrução Normativa IAT nº 65/2025",
      "licença_anterior": "Licença de Operação (LO)",
      "nº_licença_anterior": "225195",
      "nº_protocolo_licença_anterior": "16.704.730-0",
      "matérias_primas": "...",
      "tipo_abastecimento_água": "rede pública",
      "descrição_abastecimento_água": "...",
      "nº_outorga_agua": "...",
      "descrição_efluentes": "...", "tratamento_efluentes": "...",
      "destinação_efluentes": "...", "vazão_efluentes": "...",
      "nº_outorga_efluentes": "...", "data_validade_portaria_efluentes": "...",
      "descrição_emissoes_atm": "...",
      "lista_residuos_destinadores": "...",
      "ano_inventário_resíduos": "...", "nº_inventário_resíduos": "...",
      // opcional: data do documento por extenso (substitui a do modelo se presente)
      "data_documento": "Londrina, 10 de junho de 2026."
  },

  "condicoes": {
      // True  -> mantém o bloco, removendo apenas marcador {(caso ...)} e chaves
      // False -> remove o bloco inteiro (marcador + conteúdo)
      "ampliacao": false,
      "licenca_anterior": true,
      "poco": false,
      "portaria_outorga_agua": false,
      "tratamento_efluentes": false,
      "pgrs": true,
      "inventario_residuos": false
  },

  // lista de documentos analisados; cada item vira um parágrafo em numeral
  // romano minúsculo (i., ii., iii. ...). Se vazia/ausente, [[docs_analisados]]
  // permanece intacto.
  "docs_analisados": ["Requerimento ...", "PCA ...", "..."]
}

REGRAS FIXAS (não dependem de flag)
-----------------------------------
- Placeholders sem valor em `valores` ficam INTACTOS (Leo preenche depois).
- Parágrafos "Durante a vistoria, ..." são MANTIDOS, apenas limpando o marcador
  {(caso haja relatório de vistoria)} e as chaves; os placeholders
  [[relatório_vistoria_*]] ficam EM BRANCO (vistoria ocorre depois).
- Bloco de assinatura (Leonardo Parreira / Engenheiro Ambiental e Sanitarista /
  Residente – IAT/ERLON) e cabeçalho/rodapé/imagens NÃO são tocados.
- CONCLUSÕES e "Relatório fotográfico" ficam em branco.
- Highlight amarelo (scaffolding de "preencher aqui") é removido do corpo.
"""
import json
import re
import sys


# --------------------------------------------------------------------------- #
# Constantes literais extraídas do modelo (byte-a-byte)
# --------------------------------------------------------------------------- #
AMPL = " {(caso seja licença ambiental de ampliação)} {que foi aumentada para  [[área_ampliação]] m²}"
AMPL_KEEP = " que foi aumentada para  [[área_ampliação]] m²"
AMPL_FALSE_REMOVE = "," + AMPL  # remove também a vírgula -> "m² e contando"

LICANT = " {(caso haja licença anterior)} {e em atendimento às condicionantes da Licença Ambiental vigente do empreendimento, referente à [[licença_anterior]] nº [[nº_licença_anterior]], registrada sob protocolo [[nº_protocolo_licença_anterior]]}"
LICANT_KEEP = " e em atendimento às condicionantes da Licença Ambiental vigente do empreendimento, referente à [[licença_anterior]] nº [[nº_licença_anterior]], registrada sob protocolo [[nº_protocolo_licença_anterior]]"

POCO_MARK = "{(caso use poço)} "
POCO_FULL_REMOVE = ", {(caso use poço)} [[descrição_abastecimento_água]]"

OUTORGA_OPEN = " {(caso haja portaria de outorga)} {O uso dos recursos hídricos está devidamente regularizado por meio das seguintes Portarias de Outorga de Direito de Uso da Água:"
OUTORGA_OPEN_KEEP = " O uso dos recursos hídricos está devidamente regularizado por meio das seguintes Portarias de Outorga de Direito de Uso da Água:"
OUTORGA_CLOSE = "[[nº_outorga_agua]]}"
OUTORGA_CLOSE_KEEP = "[[nº_outorga_agua]]"

EFL = " {(caso use tratamento de efluentes)} {O tratamento é realizado em [[tratamento_efluentes]], após o qual o efluente tratado é [[destinação_efluentes]], com vazão de [[vazão_efluentes]], conforme Portaria de Outorga de direito nº [[nº_outorga_efluentes]], com validade até a data [[data_validade_portaria_efluentes]].}"
EFL_KEEP = " O tratamento é realizado em [[tratamento_efluentes]], após o qual o efluente tratado é [[destinação_efluentes]], com vazão de [[vazão_efluentes]], conforme Portaria de Outorga de direito nº [[nº_outorga_efluentes]], com validade até a data [[data_validade_portaria_efluentes]]."

INV = "{(caso exista Inventário de Resíduos)} {Cumpre destacar que a empresa entregou o Inventário de Resíduos referente ao ano de [[ano_inventário_resíduos]], registrado sob o nº [[nº_inventário_resíduos]].}"
INV_KEEP = "Cumpre destacar que a empresa entregou o Inventário de Resíduos referente ao ano de [[ano_inventário_resíduos]], registrado sob o nº [[nº_inventário_resíduos]]."

PGRS_RAW = "declarados {e apresentados no PGRS], verificou-se"
PGRS_YES = "declarados e apresentados no PGRS, verificou-se"
PGRS_NO = "declarados, verificou-se"

# parágrafos de vistoria (paraId -> texto final desejado, placeholder mantido)
VISTORIA = {
    "0720426D": "Durante a vistoria, [[relatório_vistoria_mat_primas]].",
    "2D2067FF": "Durante a vistoria, [[relatório_vistoria_mat_primas]].",
    "3498D2E7": "Durante a vistoria, [[relatório_vistoria_mat_primas]].",
    "65DFB757": "Durante a vistoria, [[relatório_vistoria_matérias_primas]].",
    # 25F3D964 já está limpo no modelo
}

DOCS_PARA_ID = "4D07DB48"          # parágrafo style="Listas" com [[docs_analisados]]
OUTORGA_LIST_PARA_ID = "6D3AF23D"  # parágrafo style="Listas" com [[nº_outorga_agua]]}
INV_PARA_ID = "69417F9B"           # parágrafo do inventário

ROMANOS = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x",
           "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii", "xviii", "xix", "xx",
           "xxi", "xxii", "xxiii", "xxiv", "xxv", "xxvi", "xxvii", "xxviii"]


def romano(n):
    return ROMANOS[n] if n < len(ROMANOS) else str(n + 1)


def remover_paragrafo(xml, para_id):
    """Remove inteiramente o <w:p> de paraId informado."""
    pat = re.compile(
        rf'<w:p w14:paraId="{para_id}"[^>]*>(?:(?!<w:p\b).)*?</w:p>\s*',
        re.DOTALL,
    )
    novo, n = pat.subn("", xml, count=1)
    if n != 1:
        raise RuntimeError(f"Falha ao remover parágrafo {para_id} (n={n})")
    return novo


def main():
    if len(sys.argv) != 3:
        print("uso: fill_template.py <pasta_unpacked> <dados.json>", file=sys.stderr)
        sys.exit(2)

    unpacked, dados_path = sys.argv[1], sys.argv[2]
    doc_path = f"{unpacked}/word/document.xml"
    with open(doc_path, encoding="utf-8") as f:
        xml = f.read()
    with open(dados_path, encoding="utf-8") as f:
        dados = json.load(f)

    valores = dados.get("valores", {}) or {}
    cond = dados.get("condicoes", {}) or {}
    docs = dados.get("docs_analisados", []) or []

    # ------------------------------------------------------------------ #
    # 1) BLOCOS CONDICIONAIS (antes de preencher placeholders simples)
    # ------------------------------------------------------------------ #
    # Ampliação
    if cond.get("ampliacao"):
        xml = xml.replace(AMPL, AMPL_KEEP)
    else:
        xml = xml.replace(AMPL_FALSE_REMOVE, "")

    # Licença anterior
    if cond.get("licenca_anterior"):
        xml = xml.replace(LICANT, LICANT_KEEP)
    else:
        xml = xml.replace(LICANT, "")

    # Poço (apenas remove o marcador, ou remove o trecho de descrição)
    if cond.get("poco"):
        xml = xml.replace(POCO_MARK, "")
    else:
        xml = xml.replace(POCO_FULL_REMOVE, "")

    # Portaria de outorga (água) — bloco cruza dois parágrafos
    if cond.get("portaria_outorga_agua"):
        xml = xml.replace(OUTORGA_OPEN, OUTORGA_OPEN_KEEP)
        xml = xml.replace(OUTORGA_CLOSE, OUTORGA_CLOSE_KEEP)
    else:
        xml = xml.replace(OUTORGA_OPEN, "")
        xml = remover_paragrafo(xml, OUTORGA_LIST_PARA_ID)

    # Tratamento de efluentes
    if cond.get("tratamento_efluentes"):
        xml = xml.replace(EFL, EFL_KEEP)
    else:
        xml = xml.replace(EFL, "")

    # PGRS (inline na seção de resíduos)
    if cond.get("pgrs"):
        xml = xml.replace(PGRS_RAW, PGRS_YES)
    else:
        xml = xml.replace(PGRS_RAW, PGRS_NO)

    # Inventário de resíduos
    if cond.get("inventario_residuos"):
        xml = xml.replace(INV, INV_KEEP)
    else:
        xml = remover_paragrafo(xml, INV_PARA_ID)

    # ------------------------------------------------------------------ #
    # 2) PARÁGRAFOS DE VISTORIA — manter, limpando marcador (placeholder fica)
    # ------------------------------------------------------------------ #
    for pid, texto_final in VISTORIA.items():
        m = re.search(rf'(<w:p w14:paraId="{pid}"[^>]*>.*?</w:p>)', xml, re.DOTALL)
        if not m:
            continue
        bloco = m.group(1)
        # substitui o texto de TODOS os <w:t> do parágrafo: 1º recebe o texto
        # limpo, os demais ficam vazios (preservando a estrutura de runs).
        ts = list(re.finditer(r'(<w:t[^>]*>)(.*?)(</w:t>)', bloco, re.DOTALL))
        if ts:
            novo_bloco = bloco
            for i, t in enumerate(ts):
                conteudo = texto_final if i == 0 else ""
                novo_bloco = novo_bloco.replace(t.group(0),
                                                t.group(1) + conteudo + t.group(3), 1)
            xml = xml.replace(bloco, novo_bloco, 1)

    # ------------------------------------------------------------------ #
    # 3) LISTA DE DOCUMENTOS ANALISADOS (numeral romano, 1 por parágrafo)
    # ------------------------------------------------------------------ #
    if docs:
        m = re.search(rf'(<w:p w14:paraId="{DOCS_PARA_ID}"[^>]*>.*?</w:p>)', xml, re.DOTALL)
        if m:
            modelo_par = m.group(1)
            # texto do primeiro documento entra no parágrafo existente
            paragrafos = []
            for i, doc_txt in enumerate(docs):
                texto = f"{romano(i)}. {doc_txt}"
                par = modelo_par
                # sem bullet: troca o estilo de lista por parágrafo de corpo
                par = par.replace('<w:pStyle w:val="Listas"/>',
                                  '<w:pStyle w:val="Textogeral"/>')
                ts = list(re.finditer(r'(<w:t[^>]*>)(.*?)(</w:t>)', par, re.DOTALL))
                for j, t in enumerate(ts):
                    conteudo = _xml_escape(texto) if j == 0 else ""
                    par = par.replace(t.group(0),
                                      t.group(1) + conteudo + t.group(3), 1)
                paragrafos.append(par)
            xml = xml.replace(modelo_par, "\n      ".join(paragrafos), 1)

    # ------------------------------------------------------------------ #
    # 4) DATA DO DOCUMENTO (opcional)
    # ------------------------------------------------------------------ #
    if valores.get("data_documento"):
        xml = xml.replace("Londrina, 26 de maio de 2026.",
                          _xml_escape(valores["data_documento"]))

    # ------------------------------------------------------------------ #
    # 5) PLACEHOLDERS SIMPLES [[campo]] -> valor (ausentes ficam intactos)
    # ------------------------------------------------------------------ #
    # NÃO preencher os placeholders de vistoria (ficam em branco por decisão).
    bloqueados = {"relatório_vistoria_mat_primas", "relatório_vistoria_matérias_primas",
                  "relatório_vistoria_resíduos"}
    # ordenar por comprimento desc evita corromper placeholders que se sobrepõem
    for chave in sorted(valores.keys(), key=len, reverse=True):
        if chave in bloqueados or chave == "data_documento":
            continue
        xml = xml.replace(f"[[{chave}]]", _xml_escape(str(valores[chave])))

    # ------------------------------------------------------------------ #
    # 6) Remover highlight amarelo (scaffolding de "preencher aqui")
    # ------------------------------------------------------------------ #
    xml = xml.replace('<w:highlight w:val="yellow"/>', "")

    with open(doc_path, "w", encoding="utf-8") as f:
        f.write(xml)

    # auditoria de placeholders remanescentes
    restantes = sorted(set(re.findall(r'\[\[[^\]]+\]\]', xml)))
    print("OK. Placeholders remanescentes (mantidos intactos para Leo):")
    for r in restantes:
        print("   ", r)


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


if __name__ == "__main__":
    main()
