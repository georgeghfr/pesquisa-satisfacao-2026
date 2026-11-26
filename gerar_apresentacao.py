#!/usr/bin/env python3
"""Gera a apresentação HTML a partir das planilhas de pesquisa de satisfação.

Uso:  python gerar_apresentacao.py [ENTRADA] [SAIDA.html] [--rascunho]
  ENTRADA = um .xlsx com uma aba por curso, OU uma pasta com os .xlsx baixados
            do Moodle (um arquivo por curso). Sem argumento, usa a pasta 'entrada'.
  --rascunho = mantém os slides vazios com o aviso "A preencher". Sem a opção, a
            apresentação sai limpa (slides sem conteúdo são omitidos) e as
            pendências aparecem apenas no terminal.
Requer: pip install openpyxl matplotlib     (modelo_apresentacao.html deve estar nesta pasta)
"""

import io
import logging
import math
import statistics
import sys, re, json, base64, mimetypes, unicodedata
from datetime import date
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit(
        "Erro: A biblioteca 'openpyxl' não está instalada.\n"
        "Instale-a executando no terminal: pip install openpyxl"
    )

try:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    # Texto do gráfico continua texto (nítido em qualquer zoom) e herda a fonte do slide.
    plt.rcParams["svg.fonttype"] = "none"
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Segoe UI", "Arial", "DejaVu Sans"]
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

AQUI = Path(__file__).resolve().parent
IGN = {
    "nil",
    "nada a acrescentar",
    "sem comentário",
    "sem comentários",
    "não há",
    "não houve",
    "nenhuma sugestão a acrescentar",
    "não tenho comentários à acrescentar",
    "sem sugestão",
    "nada a comentar",
    "nada a registrar",
    "nada a declarar",
    "nada a sugerir",
    "nada a relatar",
    "nada a dizer",
    "nada a relatar, pois todos os pontos foram bem abordados",
    "não tenho nada a sugerir, tendo em vista que o conteúdo foi bastante satisfatório para o meu aprendizado",
}
ESPERADO = {"P": 30, "E": 14}  # Itens de escala 1-10 avaliados (Presencial / EAD)

NOTA_CORTE = 7.0  # nota abaixo da qual a faixa/aluno é destacado
DELTA_MIN = 0.10  # variação menor que isto é tratada como "estável" (cinza)

# Siglas que aparecem escritas de mais de um jeito; a da direita é a oficial.
ALIAS_SIGLA = {"CDDAE": "CDDAAE"}

NOTAS_PADRAO = [
    {"titulo": "Desempenho dos Discentes", "imagem": ""},
    {"titulo": "Frequência de Notas da Turma", "imagem": ""},
    {
        "titulo": "Histograma do Desempenho da Turma em relação aos 3 últimos anos",
        "imagem": "",
    },
]

SKELETON_BASE = {
    "nome": "",
    "sigla": "",
    "edicao": "",
    "periodo_local": "",
    "coordenacao": {"Coordenação Técnica": ""},
    "parecer": "",
    "acoes": [],
}

AVISOS: list[str] = []  # problemas de dados encontrados (impressos no final)


def skeleton_por_tipo(tipo: str) -> dict:
    base = json.loads(json.dumps(SKELETON_BASE))
    if tipo == "P":
        base["notas"] = json.loads(json.dumps(NOTAS_PADRAO))
    return base


def limpo(x):
    if x is None:
        return ""
    t = re.sub(r"<br\s*/?>", " ", str(x)).strip()
    return "" if len(t) < 2 or t.lower().strip(" .") in IGN else t


def norm_txt(v):
    if v is None:
        return ""
    s = re.sub(r"\s+", " ", str(v)).strip()
    return "" if s.lower() in {"none", "nan"} else s


def parse_num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        val = float(v)
        return val if math.isfinite(val) else None
    s = str(v).strip().replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        val = float(s)
        return val if math.isfinite(val) else None
    except ValueError:
        return None


def fmt_pt(val, dec=2):
    if val is None or (isinstance(val, float) and not math.isfinite(val)):
        return "—"
    return f"{val:.{dec}f}".replace(".", ",")


def norm_key(k):
    s = re.sub(r"\s+", " ", str(k or "")).strip()
    s = re.sub(r"\s*_\s*(\d+)$", r"_\1", s)
    return s


def norm_col(c):
    s = unicodedata.normalize("NFKD", str(c or ""))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).strip().lower()


def limpar_periodo(p):
    s = norm_txt(p)
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1].strip()
    return re.sub(r"\s+", " ", s)


# ----------------------------------------------------------------------------
# Padronização dos textos que vão para a capa (erros de digitação, travessões,
# patentes) e validação do período.
# ----------------------------------------------------------------------------
MESES = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12,
}
TYPOS = [(re.compile(r"\bjuh+o\b", re.I), "julho")]

PATENTES = {
    "cap": "Cap", "maj": "Maj", "cel": "Cel", "ten": "Ten", "brig": "Brig",
    "tc": "TC", "cp": "CP", "1t": "1T", "2t": "2T", "1s": "1S", "2s": "2S",
    "3s": "3S", "sgt": "Sgt",
}
QUADROS = {"AV", "INF", "MRM", "BCT", "QOINF", "QOENG", "QOCON", "QOEA", "QOAV", "INT", "ENG", "SAR"}
PARTICULAS = {"de", "da", "do", "dos", "das", "e"}


def corrigir_texto(s):
    s = norm_txt(s)
    for rx, rep in TYPOS:
        s = rx.sub(rep, s)
    return re.sub(r"\s+[-–—]\s+", " – ", s)


def norm_pessoa(txt):
    """'CAP AV Domingues' -> 'Cap AV Domingues'; 'Maj MARTINS' -> 'Maj Martins'."""
    partes = []
    for tok in norm_txt(txt).split(" "):
        if not tok:
            continue
        chave = tok.lower()
        if chave in PATENTES:
            partes.append(PATENTES[chave])
        elif tok.upper() in QUADROS:
            partes.append(tok.upper())
        elif any(ch.isdigit() for ch in tok):
            partes.append(tok)
        elif chave in PARTICULAS and partes:
            partes.append(chave)
        else:
            partes.append(tok[:1].upper() + tok[1:].lower())
    return " ".join(partes)


def normalizar_item(item: dict) -> dict:
    item["nome"] = corrigir_texto(item.get("nome"))
    item["edicao"] = corrigir_texto(item.get("edicao"))
    item["periodo_local"] = corrigir_texto(item.get("periodo_local"))
    sig = norm_txt(item.get("sigla"))
    item["sigla"] = ALIAS_SIGLA.get(sig, sig)
    coord = item.get("coordenacao")
    if isinstance(coord, dict):
        for k_, v in list(coord.items()):
            coord[k_] = norm_pessoa(v)
    return item


def validar_periodo(chave: str, periodo: str):
    if not periodo:
        return f"{chave}: período vazio"
    m = re.search(
        r"(\d{1,2})\s+([A-Za-zçÇãÃ]+)\s+à\s+(\d{1,2})\s+([A-Za-zçÇãÃ]+)", periodo
    )
    if not m:
        return f"{chave}: período fora do padrão ('{periodo}')"
    d1, m1, d2, m2 = m.groups()
    mes1, mes2 = MESES.get(m1.lower()), MESES.get(m2.lower())
    if not mes1 or not mes2:
        return f"{chave}: mês não reconhecido em '{periodo}'"
    try:
        ini, fim = date(2024, mes1, int(d1)), date(2024, mes2, int(d2))
    except ValueError:
        return f"{chave}: data inválida em '{periodo}'"
    if fim < ini:
        return f"{chave}: o período termina antes de começar ('{periodo}'); confira a planilha"
    return None


def ler_folha(ws, nome):
    R = [tuple(r) + (None,) * 4 for r in ws.iter_rows(values_only=True)]
    if len(R) < 6:
        return None
    try:
        n = int(re.search(r"(\d+)", str(R[1][0])).group(1))
        nq = int(re.search(r"(\d+)", str(R[2][0])).group(1))
    except Exception:
        print(f"  [ignorada] {nome}: formato inesperado")
        return None
    if n == 0:
        return None
    qs = [
        i
        for i, r in enumerate(R)
        if i >= 5 and isinstance(r[0], (int, float)) and isinstance(r[1], str)
    ]
    q, sc, cm, mult, qo = {}, [], {}, [], {}
    for a, i in enumerate(qs):
        r, fim = R[i], (qs[a + 1] if a + 1 < len(qs) else len(R))
        nx = R[i + 1] if i + 1 < len(R) else (None,) * 6
        chave = r[1].split()[0].rstrip(".")
        if r[2] is not None and nx[0] is None and isinstance(nx[2], (int, float)):
            ops = [x for x in r[2:] if x is not None]
            c = [float(x or 0) for x in nx[2 : 2 + len(ops)]]
            if ops[0] == "Não observado.":
                sc.append([r[1], c, int(r[0])])
            else:
                q[chave] = c
                qo[chave] = int(r[0])
                if chave == "4.2" and nq >= 50:
                    mult = [
                        [re.sub(r"[\s.]+$", "", str(o)), int(v)] for o, v in zip(ops, c)
                    ]
        else:
            tx = [t for t in (limpo(row[2]) for row in R[i:fim]) if t]
            if tx:
                cm[chave] = [r[1].strip(), tx, int(r[0])]
    t = "P" if nq >= 50 else "E"

    if nome in {"COSR", "CADS"} and not qs:
        for idx in range(3, min(len(R), 15)):
            r0, r1 = R[idx][0], R[idx][1]
            print(
                f"    [DEBUG {nome}] linha {idx}: "
                f"r[0]={r0!r} ({type(r0).__name__}), "
                f"r[1]={r1!r} ({type(r1).__name__})"
            )

    if len(sc) != ESPERADO[t]:
        print(
            f"  [ERRO] {nome}: {len(sc)} itens de escala (esperado {ESPERADO[t]}). Curso ignorado."
        )
        return None
    return {"t": t, "n": n, "q": q, "s": sc, "c": cm, "m": mult, "qo": qo}


def localizar_arquivo_por_nomes(ent: Path, nomes_base: list[str]):
    candidatos = []
    for base in nomes_base:
        candidatos.extend(
            [
                AQUI / f"{base}.xlsx",
                ent / f"{base}.xlsx" if ent.is_dir() else None,
            ]
        )
    for p in candidatos:
        if p and p.is_file():
            return p
    for p in AQUI.glob("*.xlsx"):
        if not p.name.startswith("~$") and any(
            norm_col(p.stem).startswith(norm_col(b)) for b in nomes_base
        ):
            return p
    return None


def localizar_planilha_info(ent: Path):
    return localizar_arquivo_por_nomes(ent, ["informaçoes", "informacoes", "info"])


def localizar_planilha_desempenho(ent: Path):
    return localizar_arquivo_por_nomes(
        ent, ["desempenho", "desempenho_cursos", "notas", "notas_cursos"]
    )


def ler_informacoes_xlsx(caminho: Path):
    if not caminho or not caminho.is_file():
        return {}
    wb = load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}

    header_idx = 0
    col_map = {}
    for idx, r in enumerate(rows[:5]):
        temp = {norm_col(v): j for j, v in enumerate(r) if norm_txt(v)}
        if any("identificador" in k or "curso" in k for k in temp):
            header_idx = idx
            col_map = temp
            break
    if not col_map:
        col_map = {norm_col(v): j for j, v in enumerate(rows[0]) if norm_txt(v)}

    def get_col(row, *prefixes):
        for k, j in col_map.items():
            if any(k.startswith(p) or p in k for p in prefixes):
                if j < len(row):
                    return row[j]
        return None

    dados = {}
    for r in rows[header_idx + 1 :]:
        raw_id = get_col(r, "identificador")
        chave = norm_key(raw_id)
        if not chave:
            continue
        nome = norm_txt(get_col(r, "nome completo", "nome"))
        edicao = norm_txt(get_col(r, "edicao", "versao"))
        periodo = limpar_periodo(get_col(r, "periodo", "local"))
        coord = norm_txt(get_col(r, "coordenador"))
        sigla = chave.split("_")[0]

        dados[chave] = {
            "nome": nome,
            "sigla": sigla,
            "edicao": edicao,
            "periodo_local": periodo,
            "coordenacao": {"Coordenação Técnica": coord},
        }
    return dados


# ----------------------------------------------------------------------------
# Gráficos (SVG, fundo transparente, cores do tema claro do modelo)
# ----------------------------------------------------------------------------
_NAVY, _TX, _MU, _AC, _GRID = "#00205F", "#0F1F3A", "#5B6B82", "#E0A21B", "#DDE3EC"
_STEEL, _PREV, _UP, _DOWN = "#2F6FB5", "#9AA6B5", "#1F8A54", "#C93C3C"
_FONTE_SVG = "'Segoe UI', Roboto, Arial, sans-serif"


def _fig_para_uri(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.15, transparent=True)
    plt.close(fig)
    svg = buf.getvalue().decode("utf-8")
    # O SVG fica dentro de <img>: sem fonte de reserva o navegador cairia em serifa.
    svg = re.sub(r"'(?:DejaVu Sans|Segoe UI|Arial)'", _FONTE_SVG, svg)
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode("utf-8")).decode()


def _estilo_base(fig, ax):
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color(_GRID)
    ax.tick_params(axis="x", colors=_MU, length=0, labelsize=14)
    ax.tick_params(axis="y", colors=_MU, length=0, labelsize=14)


def _grafico_frequencia_uri(notas: list[float], titulo_curso: str) -> str:
    if not HAS_MPL or not notas:
        return ""
    n_alunos = len(notas)
    k_bins = max(1, math.ceil(math.sqrt(n_alunos)))
    min_n, max_n = min(notas), max(notas)
    media = sum(notas) / n_alunos

    fig, ax = plt.subplots(figsize=(11, 5))
    _estilo_base(fig, ax)

    x7 = None
    if math.isclose(min_n, max_n):
        counts, x = np.array([n_alunos]), np.arange(1)
        labels = [fmt_pt(min_n, 2)]
        x_media = 0
        abaixo = [min_n < NOTA_CORTE]
    else:
        counts, edges = np.histogram(notas, bins=k_bins, range=(min_n, max_n))
        labels = [
            f"{fmt_pt(edges[i], 2)}\n{fmt_pt(edges[i + 1], 2)}"
            for i in range(len(edges) - 1)
        ]
        x = np.arange(len(counts))
        larg = edges[1] - edges[0]
        x_media = (media - edges[0]) / larg - 0.5
        # só pinta a faixa que está inteira abaixo do corte (as que o cruzam ficam neutras)
        abaixo = [edges[i + 1] <= NOTA_CORTE + 1e-9 for i in range(len(counts))]
        if min_n < NOTA_CORTE < max_n:
            x7 = (NOTA_CORTE - edges[0]) / larg - 0.5

    max_c = int(max(counts)) if len(counts) else 1
    cores = [_AC if a else _STEEL for a in abaixo]
    bars = ax.bar(x, counts, width=0.72, color=cores, zorder=3)
    for b in bars:
        b.set_linewidth(0)
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_yticks([])
    ax.set_ylim(0, max_c * 1.32)
    ax.set_xlim(-0.6, len(counts) - 0.4)

    for b, c in zip(bars, counts):
        ax.text(
            b.get_x() + b.get_width() / 2,
            b.get_height() + max_c * 0.03,
            str(int(c)),
            ha="center", va="bottom", fontsize=18, fontweight="bold", color=_TX,
        )
    if x7 is not None:
        ax.axvline(x7, color=_AC, linestyle=(0, (2, 3)), linewidth=2, zorder=4)
    ax.axvline(x_media, color=_TX, linestyle=(0, (4, 4)), linewidth=1.4, zorder=4)
    ax.text(
        x_media, max_c * 1.27, f"Média {fmt_pt(media, 2)}",
        ha="center", va="bottom", fontsize=14, fontweight="bold", color="#FFFFFF",
        bbox=dict(boxstyle="round,pad=0.35", fc=_NAVY, ec="none"),
    )
    rotulo = f"Intervalo de notas · {n_alunos} alunos"
    if x7 is not None or any(abaixo):
        rotulo += f" · âmbar: abaixo de {fmt_pt(NOTA_CORTE, 1)}"
    ax.set_xlabel(rotulo, fontsize=13, color=_MU, labelpad=12)
    return _fig_para_uri(fig)


def _serie_historico(historico_anos, media_atual: float, edicao_txt: str):
    """Até 2 períodos anteriores + o atual, na ordem em que o gráfico os mostra."""
    ano_atual = "Atual"
    m_ano = re.search(r"(20\d{2})", norm_txt(edicao_txt))
    if m_ano:
        ano_atual = m_ano.group(1)

    anteriores = []
    vistos = set()
    for rotulo, val in historico_anos:
        chave_a = rotulo.strip()
        if not chave_a or chave_a == ano_atual or chave_a in vistos:
            continue
        vistos.add(chave_a)
        anteriores.append((chave_a, float(val)))

    anteriores = anteriores[-2:]
    return anteriores + [(ano_atual, float(media_atual))]


def _grafico_historico_uri(serie: list[tuple[str, float]]) -> str:
    if not HAS_MPL:
        return ""
    rotulos = [r for r, _ in serie]
    valores = [v for _, v in serie]

    fig, ax = plt.subplots(figsize=(10, 5))
    _estilo_base(fig, ax)
    cores = [_PREV] * (len(valores) - 1) + [_STEEL]
    x = np.arange(len(valores))
    # pontos ligados por linha: não exigem eixo começando em zero (evita exagerar diferenças)
    ax.plot(x, valores, color=_PREV, linewidth=3, zorder=2)
    ax.scatter(x, valores, s=420, c=cores, zorder=3, edgecolors="#FFFFFF", linewidths=3)

    min_v, max_v = min(valores), max(valores)
    y_min = max(0.0, math.floor((min_v - 0.5) * 2) / 2)
    y_max = min(10.0, max(max_v + 0.6, y_min + 1.0))
    ax.set_ylim(y_min, y_max)
    ax.set_xlim(-0.5, len(valores) - 0.5)
    ax.yaxis.set_major_formatter(lambda v, _: fmt_pt(v, 1))
    ax.set_xticks(x)
    ax.set_xticklabels(rotulos, fontsize=16, fontweight="bold", color=_TX)
    ax.yaxis.grid(True, color=_GRID, linewidth=1, zorder=0)
    ax.set_axisbelow(True)
    ax.set_xlabel(
        "Média da turma por período (o eixo não começa em zero)",
        fontsize=13, color=_MU, labelpad=12,
    )

    for xi, v in zip(x, valores):
        ax.text(
            xi,
            v + (y_max - y_min) * 0.06,
            fmt_pt(v, 2),
            ha="center", va="bottom", fontsize=20, fontweight="bold", color=_TX,
        )
    if len(valores) > 1:  # variação em relação ao período anterior
        # mesma conta do título do slide: diferença entre os valores já arredondados
        d = round(round(valores[-1], 2) - round(valores[-2], 2), 2)
        sinal = "+" if d >= 0 else "−"
        if abs(d) < DELTA_MIN:
            texto, cor = f"≈ estável ({sinal}{fmt_pt(abs(d), 2)} vs. {rotulos[-2]})", _MU
        else:
            seta = "▲" if d >= 0 else "▼"
            texto = f"{seta} {sinal}{fmt_pt(abs(d), 2)} vs. {rotulos[-2]}"
            cor = _UP if d >= 0 else _DOWN
        ax.text(
            x[-1], valores[-1] + (y_max - y_min) * 0.17, texto,
            ha="center", va="bottom", fontsize=14, fontweight="bold", color=cor,
        )
    return _fig_para_uri(fig)


def _moda_texto(notas: list[float]) -> str:
    if not notas:
        return "Sem moda"
    contagens: dict[float, int] = {}
    for n in notas:
        k = round(float(n), 6)
        contagens[k] = contagens.get(k, 0) + 1
    max_freq = max(contagens.values())
    if max_freq <= 1:
        return "Sem moda"
    return ", ".join(fmt_pt(k, 2) for k, f in contagens.items() if f == max_freq)


def ler_aba_desempenho(ws, chave_curso: str, edicao_txt: str, nome_curso: str):
    notas: list[float] = []
    historico_raw: list[tuple[int | None, str, float]] = []

    # Coluna A (idx 0): aluno | Coluna B (idx 1): média final | Coluna C (idx 2): ano | Coluna D (idx 3): média do ano
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = list(row) + [None] * 6
        aluno = norm_txt(r[0])
        nota = parse_num(r[1])
        if aluno and nota is not None:
            notas.append(nota)

        ano_val = parse_num(r[2])
        media_hist = parse_num(r[3])
        if ano_val is not None and media_hist is not None:
            ano_int = int(round(ano_val))
            rotulo_ano = (
                str(ano_int) if math.isclose(ano_val, ano_int) else norm_txt(r[2])
            )
            historico_raw.append(
                (ano_int if math.isclose(ano_val, ano_int) else None, rotulo_ano, media_hist)
            )

    if not notas:
        return None

    media = statistics.mean(notas)
    mediana = statistics.median(notas)
    maior_grau = max(notas)
    menor_grau = min(notas)
    amplitude = maior_grau - menor_grau
    desvio_padrao = statistics.stdev(notas) if len(notas) >= 2 else 0.0
    acima_media = sum(1 for x in notas if x > media)
    abaixo_media = sum(1 for x in notas if x < media)
    abaixo_corte = sum(1 for x in notas if x < NOTA_CORTE)
    moda_str = _moda_texto(notas)

    historico_ordenado = [
        (rot, med)
        for _, rot, med in sorted(
            historico_raw,
            key=lambda item: (item[0] is None, item[0] if item[0] is not None else 0),
        )
    ]

    rotulo_exibicao = nome_curso or chave_curso
    serie = _serie_historico(historico_ordenado, media, edicao_txt)
    img_freq = _grafico_frequencia_uri(notas, rotulo_exibicao)
    img_hist = _grafico_historico_uri(serie)

    return {
        "total_alunos": len(notas),
        "media_num": media,
        "abaixo_corte_num": abaixo_corte,
        "corte": NOTA_CORTE,
        "hist": [[r, v] for r, v in serie],
        "estatisticas": [
            {"rotulo": "MÉDIA", "valor": fmt_pt(media, 2)},
            {"rotulo": "MODA", "valor": moda_str},
            {"rotulo": "MEDIANA", "valor": fmt_pt(mediana, 2)},
            {"rotulo": "AMPLITUDE", "valor": fmt_pt(amplitude, 2)},
            {"rotulo": "DESVIO PADRÃO", "valor": fmt_pt(desvio_padrao, 2)},
            {"rotulo": "NOTAS ACIMA DA MÉDIA", "valor": str(acima_media)},
            {"rotulo": "NOTAS ABAIXO DA MÉDIA", "valor": str(abaixo_media)},
            {
                "rotulo": f"NOTAS ABAIXO DE {fmt_pt(NOTA_CORTE, 1)}",
                "valor": str(abaixo_corte),
            },
            {"rotulo": "MAIOR GRAU", "valor": fmt_pt(maior_grau, 2)},
            {"rotulo": "MENOR GRAU", "valor": fmt_pt(menor_grau, 2)},
        ],
        "img_frequencia": img_freq,
        "img_historico": img_hist,
    }


def ler_desempenho_xlsx(caminho: Path, info_map: dict):
    if not caminho or not caminho.is_file():
        return {}
    wb = load_workbook(caminho, read_only=True, data_only=True)
    resultado = {}
    for ws in wb.worksheets:
        chave = norm_key(ws.title)
        while chave in resultado:
            chave += "_"
        meta = info_map.get(chave, {})
        dados_aba = ler_aba_desempenho(
            ws,
            chave,
            meta.get("edicao", ""),
            meta.get("nome", ""),
        )
        if dados_aba:
            resultado[chave] = dados_aba
    return resultado


def mesclar_info(base_json: dict, planilha_item: dict, tipo: str = "P"):
    item = skeleton_por_tipo(tipo)
    if isinstance(base_json, dict):
        for k, v in base_json.items():
            if k in {"notas", "desempenho"} and tipo != "P":
                continue
            if k == "coordenacao" and isinstance(v, dict):
                item["coordenacao"].update(v)
            elif v is not None:
                item[k] = v

    if tipo == "P" and not item.get("notas"):
        item["notas"] = json.loads(json.dumps(NOTAS_PADRAO))

    if planilha_item:
        for campo in ("nome", "sigla", "edicao", "periodo_local"):
            val = norm_txt(planilha_item.get(campo))
            if val:
                item[campo] = val
        coord_plan = norm_txt(
            (planilha_item.get("coordenacao") or {}).get("Coordenação Técnica")
        )
        if coord_plan:
            item.setdefault("coordenacao", {})["Coordenação Técnica"] = coord_plan
    return normalizar_item(item)


def uri(caminho):
    if not caminho:
        return ""
    f = AQUI / caminho
    if not f.is_file():
        print(f"  [AVISO] imagem não encontrada: {caminho}")
        return ""
    mt = mimetypes.guess_type(f.name)[0] or "image/png"
    return f"data:{mt};base64," + base64.b64encode(f.read_bytes()).decode()


def _vazio(v) -> bool:
    if isinstance(v, (list, tuple)):
        return not any(norm_txt(x) for x in v)
    return not norm_txt(v)


def listar_pendencias(cursos: dict, info: dict, desemp: dict) -> list[str]:
    linhas = []
    for k, d in cursos.items():
        i = info.get(k, {})
        faltas = []
        if not norm_txt(i.get("nome")):
            faltas.append("nome")
        if not norm_txt(i.get("periodo_local")):
            faltas.append("período")
        if not any(norm_txt(v) for v in (i.get("coordenacao") or {}).values()):
            faltas.append("coordenação")
        if d["t"] == "P" and k not in desemp:
            faltas.append("notas da turma (planilha de desempenho)")
        if _vazio(i.get("parecer")):
            faltas.append("parecer")
        if _vazio(i.get("acoes")):
            faltas.append("ações demandadas")
        if faltas:
            linhas.append(f"  {k}: faltam {', '.join(faltas)}")
    return linhas


def main():
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    args = [arg for arg in sys.argv[1:] if not arg.startswith("--")]
    rascunho = "--rascunho" in flags
    ent = Path(args[0]) if len(args) > 0 else AQUI / "entrada"
    saida = (
        Path(args[1]) if len(args) > 1 else AQUI / "apresentacao_satisfacao_cursos.html"
    )

    plan_info = localizar_planilha_info(ent)
    info_planilha = ler_informacoes_xlsx(plan_info)
    if plan_info:
        print(
            f"Informações carregadas de: {plan_info.name} ({len(info_planilha)} cursos)"
        )

    plan_desemp = localizar_planilha_desempenho(ent)
    desemp_planilha = ler_desempenho_xlsx(plan_desemp, info_planilha)
    if plan_desemp:
        print(
            f"Desempenho carregado de: {plan_desemp.name} ({len(desemp_planilha)} abas)"
        )
    elif not HAS_MPL:
        print(
            "  [AVISO] matplotlib não instalado; gráficos de notas não serão gerados."
        )

    arqs = (
        [ent]
        if ent.is_file()
        else sorted(
            p
            for p in ent.rglob("*.xlsx")
            if not p.name.startswith("~$")
            and (not plan_info or p.resolve() != plan_info.resolve())
            and (not plan_desemp or p.resolve() != plan_desemp.resolve())
            and not norm_col(p.stem).startswith(("informac", "desempenh", "notas"))
        )
    )
    if not arqs:
        sys.exit(f"Nenhum .xlsx de respostas encontrado em {ent}")

    cursos = {}
    for f in arqs:
        wb = load_workbook(f, read_only=True, data_only=True)
        for ws in wb.worksheets:
            nome_raw = ws.title if len(wb.worksheets) > 1 else f.stem
            nome = norm_key(nome_raw)
            while nome in cursos:
                nome += "_"
            d = ler_folha(ws, nome)
            if d:
                cursos[nome] = d
                print(
                    f'{nome}: {"presencial" if d["t"] == "P" else "EAD"}, {d["n"]} respostas'
                )

    if not cursos:
        sys.exit("Nenhum curso com respostas foi lido.")

    jf = AQUI / "informacoes_cursos.json"
    info_salvo = json.loads(jf.read_text(encoding="utf-8")) if jf.is_file() else {}
    info_salvo_norm = {norm_key(k): v for k, v in info_salvo.items() if k != "_LEIA-ME"}

    todas_chaves = list(dict.fromkeys(list(cursos.keys()) + list(info_planilha.keys())))
    novo_info = {}
    if "_LEIA-ME" in info_salvo:
        novo_info["_LEIA-ME"] = info_salvo["_LEIA-ME"]

    for k in todas_chaves:
        tipo_curso = cursos[k]["t"] if k in cursos else "P"
        novo_info[k] = mesclar_info(
            info_salvo_norm.get(k), info_planilha.get(k), tipo=tipo_curso
        )

    jf.write_text(json.dumps(novo_info, ensure_ascii=False, indent=2), encoding="utf-8")

    for k in cursos:
        aviso = validar_periodo(k, novo_info[k].get("periodo_local", ""))
        if aviso:
            AVISOS.append(aviso)

    for k, d in cursos.items():
        i = json.loads(json.dumps(novo_info[k]))
        if d["t"] == "P":
            dp = desemp_planilha.get(k)
            if dp:
                i["desempenho"] = dp
                imgs = [dp.get("img_frequencia", ""), dp.get("img_historico", "")]
                for idx_n, n in enumerate(i.get("notas", [])):
                    if idx_n == 1 and imgs[0]:
                        n["imagem"] = imgs[0]
                    elif idx_n == 2 and imgs[1]:
                        n["imagem"] = imgs[1]
                    elif n.get("imagem"):
                        n["imagem"] = uri(n.get("imagem", ""))
            else:
                for n in i.get("notas", []):
                    n["imagem"] = uri(n.get("imagem", ""))
        else:
            i.pop("notas", None)
            i.pop("desempenho", None)
        d["i"] = i

    modelo = (AQUI / "modelo_apresentacao.html").read_text(encoding="utf-8")
    if "__DATA__" not in modelo:
        sys.exit("modelo_apresentacao.html não contém o marcador __DATA__.")
    dados = json.dumps(cursos, ensure_ascii=False).replace("</", "<\\/")
    saida.write_text(
        modelo.replace("__DATA__", dados)
        .replace("__LOGO__", uri("logo.png"))
        .replace("__DRAFT__", "true" if rascunho else "false"),
        encoding="utf-8",
    )
    print(f"\nPronto: {saida}  ({len(cursos)} cursos)")

    if AVISOS:
        print("\nAvisos de dados (confira nas planilhas):")
        for a in AVISOS:
            print(f"  {a}")
    pend = listar_pendencias(cursos, novo_info, desemp_planilha)
    if pend:
        modo = "mantidos como 'A preencher'" if rascunho else "omitidos da apresentação"
        print(f"\nPendências (slides vazios {modo}; use --rascunho para vê-los):")
        print("\n".join(pend))


if __name__ == "__main__":
    main()
