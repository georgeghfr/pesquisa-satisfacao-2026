#!/usr/bin/env python3
"""Gera a apresentação HTML a partir das planilhas de pesquisa de satisfação.

Uso:  python gerar_apresentacao.py [ENTRADA] [SAIDA.html]
  ENTRADA = um .xlsx com uma aba por curso, OU uma pasta com os .xlsx baixados
            do Moodle (um arquivo por curso). Sem argumento, usa a pasta 'entrada'.
Requer: pip install openpyxl     (modelo_apresentacao.html deve estar nesta pasta)
"""
import sys, re, json, base64, mimetypes
from pathlib import Path
from openpyxl import load_workbook

AQUI = Path(__file__).resolve().parent
# Comentários sem conteúdo (comparados em minúsculas, sem pontos/espaços nas pontas)
IGN = {'nil', 'nada a acrescentar', 'sem comentário', 'sem comentários', 'não há', 'não houve',
       'nenhuma sugestão a acrescentar', 'não tenho comentários à acrescentar', 'sem sugestão',
       'nada a comentar', 'nada a registrar', 'nada a declarar', 'nada a sugerir', 'nada a relatar',
       'nada a dizer', 'nada a relatar, pois todos os pontos foram bem abordados',
       'não tenho nada a sugerir, tendo em vista que o conteúdo foi bastante satisfatório para o meu aprendizado'}
ESPERADO = {'P': 30, 'E': 14}  # itens avaliados de 1 a 6 (presencial / EAD)


def limpo(x):
    if x is None:
        return ''
    t = re.sub(r'<br\s*/?>', ' ', str(x)).strip()
    return '' if len(t) < 2 or t.lower().strip(' .') in IGN else t


def ler_folha(ws, nome):
    R = [tuple(r) + (None,) * 4 for r in ws.iter_rows(values_only=True)]
    if len(R) < 6:  # aba vazia
        return None
    try:
        n = int(re.search(r'(\d+)', str(R[1][0])).group(1))
        nq = int(re.search(r'(\d+)', str(R[2][0])).group(1))
    except Exception:
        print(f'  [ignorada] {nome}: formato inesperado'); return None
    if n == 0:
        return None
    qs = [i for i, r in enumerate(R) if i >= 5 and isinstance(r[0], (int, float)) and isinstance(r[1], str)]
    q, sc, cm, mult, qo = {}, [], {}, [], {}
    for a, i in enumerate(qs):
        r, fim = R[i], (qs[a + 1] if a + 1 < len(qs) else len(R))
        nx = R[i + 1] if i + 1 < len(R) else (None,) * 6
        chave = r[1].split()[0].rstrip('.')
        if r[2] is not None and nx[0] is None and isinstance(nx[2], (int, float)):  # pergunta fechada
            ops = [x for x in r[2:] if x is not None]
            c = [float(x or 0) for x in nx[2:2 + len(ops)]]
            if ops[0] == 'Não observado.':
                sc.append([r[1], c, int(r[0])])  # int(r[0]) = ordem na planilha
            else:
                q[chave] = c
                qo[chave] = int(r[0])
                if chave == '4.2':  # múltipla escolha: [alternativa, quantidade]
                    mult = [[re.sub(r'[\s.]+$', '', str(o)), int(v)] for o, v in zip(ops, c)]
        else:  # pergunta aberta: comentários na coluna C
            tx = [t for t in (limpo(row[2]) for row in R[i:fim]) if t]
            if tx:
                cm[chave] = [r[1].strip(), tx, int(r[0])]  # int(r[0]) = ordem na planilha
    t = 'P' if nq >= 50 else 'E'
    if len(sc) != ESPERADO[t]:
        print(f'  [ERRO] {nome}: {len(sc)} itens de escala 1-6 (esperado {ESPERADO[t]}). Curso ignorado.'); return None
    return {'t': t, 'n': n, 'q': q, 's': sc, 'c': cm, 'm': mult, 'qo': qo}


SKELETON = {'nome': '', 'sigla': '', 'edicao': '', 'finalidade': '', 'objetivos_intro': '', 'objetivos': [],
            'perfil_titulo': 'Perfil Profissional ou Perfil de Egresso', 'perfil_intro': '', 'perfil': [],
            'coordenacao': {'Coordenação Técnica': ''},
            'parecer': '',
            'notas': [{'titulo': 'Desempenho dos Discentes', 'imagem': ''},
                      {'titulo': 'Frequência de Notas da Turma', 'imagem': ''},
                      {'titulo': 'Histograma do Desempenho da Turma em relação aos 3 últimos anos', 'imagem': ''}],
            'acoes': []}


def uri(caminho):
    """Converte um arquivo de imagem em data URI (vazio se não existir)."""
    if not caminho:
        return ''
    f = AQUI / caminho
    if not f.is_file():
        print(f'  [AVISO] imagem não encontrada: {caminho}'); return ''
    mt = mimetypes.guess_type(f.name)[0] or 'image/png'
    return f'data:{mt};base64,' + base64.b64encode(f.read_bytes()).decode()


def main():
    ent = Path(sys.argv[1]) if len(sys.argv) > 1 else AQUI / 'entrada'
    saida = Path(sys.argv[2]) if len(sys.argv) > 2 else AQUI / 'apresentacao_satisfacao_cursos.html'
    arqs = [ent] if ent.is_file() else sorted(p for p in ent.rglob('*.xlsx') if not p.name.startswith('~$'))
    if not arqs:
        sys.exit(f'Nenhum .xlsx encontrado em {ent}')
    cursos = {}
    for f in arqs:
        wb = load_workbook(f, read_only=True, data_only=True)
        for ws in wb.worksheets:
            nome = ws.title if len(wb.worksheets) > 1 else f.stem
            while nome in cursos:
                nome += '_'
            d = ler_folha(ws, nome)
            if d:
                cursos[nome] = d
                print(f'{nome}: {"presencial" if d["t"] == "P" else "EAD"}, {d["n"]} respostas, '
                      f'{sum(len(v[1]) for v in d["c"].values())} comentários')
    if not cursos:
        sys.exit('Nenhum curso com respostas foi lido.')
    jf = AQUI / 'informacoes_cursos.json'
    info = json.loads(jf.read_text(encoding='utf-8')) if jf.is_file() else {}
    novos = [k for k in cursos if k not in info]
    for k in novos:  # cursos novos entram em branco no arquivo para você preencher
        info[k] = json.loads(json.dumps(SKELETON))
    if novos or not jf.is_file():
        jf.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f'Adicionados em branco a informacoes_cursos.json: {", ".join(novos)}')
    for k, d in cursos.items():
        i = json.loads(json.dumps(info[k]))
        for n in i.get('notas', []):
            n['imagem'] = uri(n.get('imagem', ''))
        d['i'] = i
    modelo = (AQUI / 'modelo_apresentacao.html').read_text(encoding='utf-8')
    dados = json.dumps(cursos, ensure_ascii=False).replace('</', '<\\/')
    saida.write_text(modelo.replace('__DATA__', dados).replace('__LOGO__', uri('logo.png')), encoding='utf-8')
    print(f'\nPronto: {saida}  ({len(cursos)} cursos)')


if __name__ == '__main__':
    main()
