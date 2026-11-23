#!/usr/bin/env python3
"""Gera a apresentação HTML a partir das planilhas de pesquisa de satisfação.

Uso:  python gerar_apresentacao.py [ENTRADA] [SAIDA.html]
  ENTRADA = um .xlsx com uma aba por curso, OU uma pasta com os .xlsx baixados
            do Moodle (um arquivo por curso). Sem argumento, usa a pasta 'entrada'.
Requer: pip install openpyxl     (modelo_apresentacao.html deve estar nesta pasta)
"""

import sys, re, json, base64, mimetypes, unicodedata
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit(
        "Erro: A biblioteca 'openpyxl' não está instalada.\n"
        "Instale-a executando no terminal: pip install openpyxl"
    )

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


def localizar_planilha_info(ent: Path):
    candidatos = [
        AQUI / "informaçoes.xlsx",
        AQUI / "informacoes.xlsx",
        ent / "informaçoes.xlsx" if ent.is_dir() else None,
        ent / "informacoes.xlsx" if ent.is_dir() else None,
    ]
    for p in candidatos:
        if p and p.is_file():
            return p
    for p in AQUI.glob("*.xlsx"):
        if not p.name.startswith("~$") and norm_col(p.stem).startswith("informac"):
            return p
    return None


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


def mesclar_info(base_json: dict, planilha_item: dict, tipo: str = "P"):
    item = skeleton_por_tipo(tipo)
    if isinstance(base_json, dict):
        for k, v in base_json.items():
            if k == "notas" and tipo != "P":
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
    return item


def uri(caminho):
    if not caminho:
        return ""
    f = AQUI / caminho
    if not f.is_file():
        print(f"  [AVISO] imagem não encontrada: {caminho}")
        return ""
    mt = mimetypes.guess_type(f.name)[0] or "image/png"
    return f"data:{mt};base64," + base64.b64encode(f.read_bytes()).decode()


def main():
    args = [arg for arg in sys.argv[1:] if not arg.startswith("--")]
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

    arqs = (
        [ent]
        if ent.is_file()
        else sorted(
            p
            for p in ent.rglob("*.xlsx")
            if not p.name.startswith("~$")
            and (not plan_info or p.resolve() != plan_info.resolve())
            and not norm_col(p.stem).startswith("informac")
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

    for k, d in cursos.items():
        i = json.loads(json.dumps(novo_info[k]))
        if d["t"] == "P":
            for n in i.get("notas", []):
                n["imagem"] = uri(n.get("imagem", ""))
        else:
            i.pop("notas", None)
        d["i"] = i

    modelo = (AQUI / "modelo_apresentacao.html").read_text(encoding="utf-8")
    dados = json.dumps(cursos, ensure_ascii=False).replace("</", "<\\/")
    saida.write_text(
        modelo.replace("__DATA__", dados).replace("__LOGO__", uri("logo.png")),
        encoding="utf-8",
    )
    print(f"\nPronto: {saida}  ({len(cursos)} cursos)")


if __name__ == "__main__":
    main()
