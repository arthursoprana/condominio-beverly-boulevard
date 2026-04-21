"""
Extrai dados de demonstrativo financeiro usando Gemini 2.5 Flash.

Uso:
    pixi run python scan_gemini.py                    # processa todas as imagens de ./imagens/
    pixi run python scan_gemini.py 2026-03.jpg        # processa imagens específicas
    pixi run python scan_gemini.py --pasta fotos/     # processa imagens de outra pasta

Variáveis de ambiente:
    GEMINI_API_KEY  — chave da Google AI Studio (https://aistudio.google.com/apikey)

Resultados são cacheados em .cache_gemini/ por hash do arquivo, então re-rodar é grátis.
Saída: planilha demonstrativos.xlsx com evolução mensal e gráficos.
"""

import argparse
import base64
import hashlib
import json
import os
import sys
import time
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
from google import genai
from google.genai import types

# ── Configuração ─────────────────────────────────────────────────
MODELO = "gemini-2.5-flash"
CACHE_DIR = Path(".cache_gemini")
PASTA_IMAGENS = Path("imagens")
EXTENSOES_IMG = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".pdf"}

PROMPT = """Você está olhando para um demonstrativo financeiro mensal de um CONDOMÍNIO brasileiro
(Demonstrativo de Receitas e Despesas). Sua tarefa: extrair TODAS as linhas que contêm um valor
monetário, preservando a ordem em que aparecem na página.

⚠️ CRÍTICO — ALINHAMENTO DESCRIÇÃO↔VALOR:
  - Cada linha horízontal da imagem tem UMA descrição (à esquerda) e UM valor (à direita).
  - NUNCA pule uma linha. Se você vê um valor mas não consegue ler bem a descrição,
    inclua mesmo assim com a melhor leitura possível — nunca puxe o valor da próxima linha.
  - Linhas pequenas/curtas (ex: "Multas e Juros" com valor 63,57) são fáceis de pular: NO ATTENTION,
    leia linha-por-linha de cima para baixo.
  - Conte as linhas com valor antes de retornar e confirme que descrições e valores têm o mesmo número.

Para cada linha, retorne:
- "descricao": texto descritivo da linha, limpo (sem pontos de líder "....", sem referências de página tipo "9 a 20")
- "valor": número decimal usando PONTO como separador decimal (ex: 1308.10, -39664.48). Negativos com sinal.
- "tipo": classifique como uma destas categorias:
    * "saldo"     — Saldo Anterior, Saldo Atual
    * "total"     — RECEITAS, DESPESAS, Receitas - Despesas (totais em CAIXA ALTA ou de seção)
    * "subtotal"  — Receitas Operacionais, Receitas Financeiras, Despesas com Pessoal,
                    Despesas Administrativas, Despesas Financeiras, Conta Transitória
    * "item"      — todas as demais linhas individuais

Regras adicionais:
- Inclua linhas com valor 0,00 (Conta Transitória pode ter zero).
- Preserve nomes próprios e números de NF/protocolo (ex: "NF 153", "NFCe 60624").
- Use grafia portuguesa correta (Síndico, Refeição, etc).

Retorne APENAS um array JSON, sem texto extra.
"""


def extrair_via_gemini(image_path: Path, client: genai.Client) -> list[dict]:
    """Extrai linhas do demonstrativo. Cacheia resultado por hash do arquivo."""
    img_bytes = image_path.read_bytes()
    img_hash = hashlib.sha256(img_bytes).hexdigest()[:16]
    cache_file = CACHE_DIR / f"{image_path.stem}_{img_hash}.json"

    if cache_file.exists():
        print(f"   📦 cache hit ({cache_file.name})")
        return json.loads(cache_file.read_text(encoding="utf-8"))

    print(f"   🌐 chamando {MODELO}...")
    mime = "image/jpeg" if image_path.suffix.lower() in (".jpg", ".jpeg") else "image/png"
    if image_path.suffix.lower() == ".pdf":
        mime = "application/pdf"

    ultimo_erro = None
    for tentativa in range(1, 5):
        try:
            resp = client.models.generate_content(
                model=MODELO,
                contents=[
                    types.Part.from_bytes(data=img_bytes, mime_type=mime),
                    PROMPT,
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0,
                    thinking_config=types.ThinkingConfig(thinking_budget=8000),
                ),
            )
            break
        except Exception as e:
            ultimo_erro = e
            msg = str(e)
            # 503/429 são transitórios → backoff exponencial
            if "503" in msg or "429" in msg or "UNAVAILABLE" in msg or "RESOURCE_EXHAUSTED" in msg:
                espera = 2 ** tentativa
                print(f"   ⏳ tentativa {tentativa} falhou ({msg[:80]}...), aguardando {espera}s")
                time.sleep(espera)
                continue
            raise
    else:
        raise RuntimeError(f"Falha após 4 tentativas: {ultimo_erro}")

    rows = json.loads(resp.text)
    CACHE_DIR.mkdir(exist_ok=True)
    cache_file.write_text(
        json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"   💾 cache salvo em {cache_file.name}")
    return rows


# ── Estilos Excel ────────────────────────────────────────────────
COR_TIPO = {
    "saldo":    PatternFill("solid", fgColor="FFE699"),  # amarelo
    "total":    PatternFill("solid", fgColor="B4C7E7"),  # azul
    "subtotal": PatternFill("solid", fgColor="D9E1F2"),  # azul claro
    "item":     None,
}
FONTE_DESTAQUE = Font(bold=True)


def escrever_aba(ws, rows: list[dict], titulo: str):
    ws.title = titulo[:31]  # limite Excel
    ws.append(["Descrição", "Valor", "Tipo"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="305496")
        c.font = Font(bold=True, color="FFFFFF")

    for row in rows:
        ws.append([row.get("descricao", ""), row.get("valor"), row.get("tipo", "item")])
        excel_row = ws.max_row
        tipo = row.get("tipo", "item")
        fill = COR_TIPO.get(tipo)
        if fill:
            for cell in ws[excel_row]:
                cell.fill = fill
        if tipo in ("saldo", "total", "subtotal"):
            for cell in ws[excel_row]:
                cell.font = FONTE_DESTAQUE
        # formato número BR
        ws.cell(row=excel_row, column=2).number_format = '#,##0.00;-#,##0.00'

    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.freeze_panes = "A2"


def descobrir_imagens(pasta: Path) -> list[Path]:
    """Encontra imagens/PDFs na pasta, ordenadas por nome (espera-se YYYY-MM)."""
    imgs = [f for f in sorted(pasta.iterdir()) if f.suffix.lower() in EXTENSOES_IMG]
    return imgs


def construir_evolucao(wb: openpyxl.Workbook, todos: dict[str, list[dict]]):
    """Cria aba 'Evolução' com pivot mensal + gráficos embutidos."""
    meses = list(todos.keys())  # já ordenados
    if not meses:
        return

    # Coletar todas as descrições (manter ordem de aparição)
    todas_desc = []
    desc_set = set()
    for mes in meses:
        for r in todos[mes]:
            d = r.get("descricao", "")
            if d and d not in desc_set:
                todas_desc.append(d)
                desc_set.add(d)

    # Montar lookup: desc → tipo (pega o primeiro encontrado)
    tipo_de = {}
    for mes in meses:
        for r in todos[mes]:
            d = r.get("descricao", "")
            if d and d not in tipo_de:
                tipo_de[d] = r.get("tipo", "item")

    # Montar pivot: {desc: {mes: valor}}
    pivot = {}
    for mes in meses:
        for r in todos[mes]:
            d = r.get("descricao", "")
            if d:
                pivot.setdefault(d, {})[mes] = r.get("valor")

    # Escrever aba
    ws = wb.create_sheet("Evolução", 0)
    # Cabeçalho
    ws.append(["Tipo", "Descrição"] + meses)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496")
    ws.freeze_panes = "C2"

    for desc in todas_desc:
        tipo = tipo_de.get(desc, "item")
        row_data = [tipo, desc]
        for mes in meses:
            row_data.append(pivot.get(desc, {}).get(mes))
        ws.append(row_data)

        # Estilo
        excel_row = ws.max_row
        fill = COR_TIPO.get(tipo)
        if fill:
            for cell in ws[excel_row]:
                cell.fill = fill
        if tipo in ("saldo", "total", "subtotal"):
            for cell in ws[excel_row]:
                cell.font = FONTE_DESTAQUE
        for col in range(3, 3 + len(meses)):
            ws.cell(row=excel_row, column=col).number_format = '#,##0.00;[Red]-#,##0.00'

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    for i in range(len(meses)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14

    if len(meses) < 2:
        return  # gráficos só fazem sentido com 2+ meses


def gerar_html(todos: dict[str, list[dict]], img_paths: dict[str, Path], output: Path):
    """Gera página HTML com gráficos interativos Plotly + visualizador de imagens."""
    meses = list(todos.keys())
    if len(meses) < 2:
        print("   ⚠️  Menos de 2 meses — HTML não gerado (precisa de evolução).")
        return

    # Montar pivot: {desc: {mes: valor}}
    pivot = {}
    for mes in meses:
        for r in todos[mes]:
            d = r.get("descricao", "")
            if d:
                pivot.setdefault(d, {})[mes] = r.get("valor")

    def serie(nome):
        return [pivot.get(nome, {}).get(m) for m in meses]

    COMMON = dict(template="plotly_white", hovermode="x unified",
                  margin=dict(l=60, r=30, t=50, b=40), dragmode=False)

    # ── Chart 1: Receitas × Despesas × Saldo Atual (line) ──
    fig1 = go.Figure()
    for nome, cor in [("RECEITAS", "#2ecc71"), ("DESPESAS", "#e74c3c"), ("Saldo Atual", "#3498db")]:
        fig1.add_trace(go.Scatter(
            x=meses, y=serie(nome), name=nome, mode="lines+markers",
            line=dict(color=cor, width=3),
            hovertemplate="R$ %{y:,.2f}<extra>" + nome + "</extra>",
        ))
    fig1.update_layout(
        title="Receitas × Despesas × Saldo Atual", height=420,
        yaxis=dict(title="R$", tickformat=",.0f"),
        xaxis=dict(showspikes=True, spikemode="across", spikesnap="cursor",
                   spikethickness=1, spikecolor="#888", spikedash="dot"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        **COMMON,
    )

    # ── Chart 2: Composição das Receitas (stacked bar) ──
    fig2 = go.Figure()
    for nome, cor in [("Receitas Operacionais", "#27ae60"), ("Receitas Financeiras", "#82e0aa")]:
        fig2.add_trace(go.Bar(
            x=meses, y=serie(nome), name=nome, marker_color=cor,
            hovertemplate="R$ %{y:,.2f}<extra>" + nome + "</extra>",
        ))
    fig2.update_layout(
        title="Composição das Receitas", barmode="stack", height=380,
        yaxis=dict(title="R$", tickformat=",.0f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        **COMMON,
    )

    # ── Chart 3: Composição das Despesas (stacked bar) ──
    fig3 = go.Figure()
    cores_desp = [("Despesas com Pessoal", "#e74c3c"), ("Despesas Administrativas", "#e67e22"), ("Despesas Financeiras", "#9b59b6")]
    for nome, cor in cores_desp:
        fig3.add_trace(go.Bar(
            x=meses, y=serie(nome), name=nome, marker_color=cor,
            hovertemplate="R$ %{y:,.2f}<extra>" + nome + "</extra>",
        ))
    fig3.update_layout(
        title="Composição das Despesas", barmode="stack", height=380,
        yaxis=dict(title="R$", tickformat=",.0f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        **COMMON,
    )

    # ── Montar HTML completo ──
    pcfg = {"scrollZoom": False, "displayModeBar": False, "responsive": True}
    div1 = fig1.to_html(include_plotlyjs="cdn", full_html=False, div_id="chart1", config=pcfg)
    div2 = fig2.to_html(include_plotlyjs=False, full_html=False, div_id="chart2", config=pcfg)
    div3 = fig3.to_html(include_plotlyjs=False, full_html=False, div_id="chart3", config=pcfg)

    # ── Encode imagens como base64 ──
    img_data_js = {}
    for mes in meses:
        p = img_paths.get(mes)
        if p and p.exists():
            b64 = base64.b64encode(p.read_bytes()).decode("ascii")
            suffix = p.suffix.lower()
            mime = "image/jpeg" if suffix in (".jpg", ".jpeg") else "image/png"
            img_data_js[mes] = f"data:{mime};base64,{b64}"

    options_html = "\n".join(
        f'        <option value="{m}">{m}</option>' for m in meses
    )
    img_map_js = json.dumps(img_data_js)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<title>Condomínio Beverly Boulevard — Evolução Financeira</title>
<style>
  body {{ font-family: system-ui, sans-serif; background: #f8f9fa; margin: 0; padding: 20px; }}
  h1 {{ text-align: center; color: #2c3e50; margin-bottom: 30px; }}
  .chart {{ max-width: 1000px; margin: 0 auto 30px; background: white;
            border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); padding: 10px;
            min-height: 380px; overflow-x: auto; }}
  .viewer {{ max-width: 1000px; margin: 0 auto 30px; background: white;
             border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); padding: 20px; }}
  .viewer h2 {{ margin: 0 0 15px; color: #2c3e50; }}
  .viewer select {{ font-size: 16px; padding: 6px 12px; border-radius: 4px;
                    border: 1px solid #ccc; margin-bottom: 15px; }}
  .viewer img {{ max-width: 100%; border: 1px solid #e0e0e0; border-radius: 4px; }}
  .viewer .placeholder {{ color: #999; font-style: italic; }}
</style>
</head>
<body>
<h1>Condomínio Beverly Boulevard — Evolução Financeira</h1>
<div class="chart">{div1}</div>
<div class="chart">{div2}</div>
<div class="chart">{div3}</div>
<div class="viewer">
  <h2>Demonstrativo Original</h2>
  <select id="mesSelect" onchange="showImage()">
    <option value="">Selecione o mês…</option>
{options_html}
  </select>
  <div id="imgContainer"><p class="placeholder">Selecione um mês acima para visualizar o demonstrativo.</p></div>
</div>
<script>
var imgMap = {img_map_js};
function showImage() {{
  var mes = document.getElementById('mesSelect').value;
  var container = document.getElementById('imgContainer');
  if (mes && imgMap[mes]) {{
    container.innerHTML = '<img src="' + imgMap[mes] + '" alt="Demonstrativo ' + mes + '">';
  }} else {{
    container.innerHTML = '<p class="placeholder">Imagem não disponível para este mês.</p>';
  }}
}}
</script>
</body>
</html>"""

    output.write_text(html, encoding="utf-8")
    print(f"   📊 HTML salvo: {output}")


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("imagens", nargs="*", help="Imagens específicas (opcional)")
    parser.add_argument("--pasta", type=Path, default=PASTA_IMAGENS,
                        help=f"Pasta com imagens (default: {PASTA_IMAGENS}/)")
    parser.add_argument("--saida", default="demonstrativos.xlsx", help="Excel de saída")
    parser.add_argument("--no-cache", action="store_true", help="Ignora cache")
    args = parser.parse_args()

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("❌ Defina a variável GEMINI_API_KEY (https://aistudio.google.com/apikey)")
        sys.exit(1)

    client = genai.Client(api_key=api_key)

    if args.no_cache and CACHE_DIR.exists():
        for f in CACHE_DIR.glob("*.json"):
            f.unlink()

    # Determinar imagens: args explícitos ou auto-scan da pasta
    if args.imagens:
        img_paths = [Path(p) for p in args.imagens]
    else:
        if not args.pasta.exists():
            args.pasta.mkdir(parents=True)
            print(f"📁 Pasta '{args.pasta}/' criada. Coloque as imagens lá e rode novamente.")
            sys.exit(0)
        img_paths = descobrir_imagens(args.pasta)
        if not img_paths:
            print(f"📁 Nenhuma imagem encontrada em '{args.pasta}/'. "
                  f"Extensões aceitas: {', '.join(sorted(EXTENSOES_IMG))}")
            sys.exit(0)

    print(f"📋 {len(img_paths)} imagem(ns) para processar")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    todos = {}  # {mes_label: rows}

    for path in sorted(img_paths):
        if not path.exists():
            print(f"⚠️  pulando {path} (não existe)")
            continue
        print(f"\n📄 {path.name}")
        try:
            rows = extrair_via_gemini(path, client)
        except Exception as e:
            print(f"   ❌ erro: {e}")
            continue
        print(f"   ✅ {len(rows)} linhas extraídas")

        mes_label = path.stem  # ex: "2025-12", "2026-02"
        todos[mes_label] = rows

        ws = wb.create_sheet()
        escrever_aba(ws, rows, mes_label)

    # Aba Evolução (pivot, sem gráficos — plots vão no HTML)
    construir_evolucao(wb, todos)

    # Aba Consolidado
    consolidado = []
    for mes, rows in todos.items():
        for r in rows:
            consolidado.append({"arquivo": mes, **r})

    if consolidado:
        ws = wb.create_sheet("Consolidado")
        ws.append(["Arquivo", "Descrição", "Valor", "Tipo"])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
        for r in consolidado:
            ws.append([r["arquivo"], r.get("descricao", ""), r.get("valor"), r.get("tipo", "item")])
            ws.cell(row=ws.max_row, column=3).number_format = '#,##0.00;-#,##0.00'
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 65
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 12
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    wb.save(args.saida)
    print(f"\n✅ Excel salvo: {args.saida}")

    # HTML com gráficos interativos
    html_path = Path(args.saida).with_suffix(".html")
    img_map = {}
    for path in sorted(img_paths):
        mes_label = path.stem
        if mes_label in todos:
            img_map[mes_label] = path
    gerar_html(todos, img_map, html_path)


if __name__ == "__main__":
    main()
