#!/usr/bin/env python3
"""Gera a apresentação HTML a partir das planilhas de pesquisa de satisfação.

Uso:  python gerador_apresentacao.py [ENTRADA] [SAIDA.html]
"""
import sys, re, json, base64, mimetypes
from pathlib import Path
from openpyxl import load_workbook

AQUI = Path(__file__).resolve().parent
IGN = {'nil', 'nada a acrescentar', 'sem comentário', 'sem comentários', 'não há', 'não houve',
       'nenhuma sugestão a acrescentar', 'não tenho comentários à acrescentar', 'sem sugestão',
       'nada a comentar', 'nada a registrar', 'nada a declarar', 'nada a sugerir', 'nada a relatar',
       'nada a dizer', 'nada a relatar, pois todos os pontos foram bem abordados',
       'não tenho nada a sugerir, tendo em vista que o conteúdo foi bastante satisfatório para o meu aprendizado'}
ESPERADO = {'P': 30, 'E': 14}


def limpo(x):
    if x is None:
        return ''
    t = re.sub(r'<br\s*/?>', ' ', str(x)).strip()
    return '' if len(t) < 2 or t.lower().strip(' .') in IGN else t


def ler_folha(ws, nome):
    R = [tuple(r) + (None,) * 4 for r in ws.iter_rows(values_only=True)]
    if len(R) < 6:
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
        if r[2] is not None and nx[0] is None and isinstance(nx[2], (int, float)):
            ops = [x for x in r[2:] if x is not None]
            c = [float(x or 0) for x in nx[2:2 + len(ops)]]
            if ops[0] == 'Não observado.':
                sc.append([r[1], c, int(r[0])])
            else:
                q[chave] = c
                qo[chave] = int(r[0])
                if chave == '4.2':
                    mult = [[re.sub(r'[\s.]+$', '', str(o)), int(v)] for o, v in zip(ops, c)]
        else:
            tx = [t for t in (limpo(row[2]) for row in R[i:fim]) if t]
            if tx:
                cm[chave] = [r[1].strip(), tx, int(r[0])]
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


def ler_planilha_info(caminho_info):
    """Lê a segunda planilha para preencher dados textuais e metadados dos cursos no JSON."""
    if not caminho_info.is_file():
        return {}
    
    wb = load_workbook(caminho_info, data_only=True)
    ws = wb.active # Pega a primeira aba (ou ajuste conforme necessário)
    
    dados_extras = {}
    # Exemplo: Assume que a planilha tem cabeçalhos na linha 1 e as linhas seguintes trazem os dados.
    # Colunas esperadas comuns: Sigla, Nome, Finalidade, etc. Ajuste conforme a estrutura real da sua planilha de info.
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]: 
            continue
        sigla = str(row[0]).strip() # Ex: Sigla do curso na coluna A
        dados_extras[sigla] = {
            'sigla': sigla,
            'nome': str(row[1]).strip() if len(row) > 1 and row[1] else '',
            'finalidade': str(row[2]).strip() if len(row) > 2 and row[2] else '',
        }
    return dados_extras


def uri(caminho):
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
    
    # Identifica os arquivos separadamente na pasta de entrada
    arqs_xlsx = list(ent.rglob('*.xlsx')) if ent.is_dir() else [ent]
    
    # Separa qual arquivo é de respostas e qual é o de informações com base no nome do arquivo
    arq_respostas = next((p for p in arqs_xlsx if 'info' not in p.name.lower() and 'curso' not in p.name.lower()), arqs_xlsx[0] if arqs_xlsx else None)
    arq_info = next((p for p in arqs_xlsx if 'info' in p.name.lower() or 'curso' in p.name.lower()), None)

    if not arq_respostas:
        sys.exit(f'Nenhuma planilha de respostas encontrada em {ent}')

    print(f'Lendo dados estatísticos de: {arq_respostas.name}')
    cursos = {}
    wb = load_workbook(arq_respostas, read_only=True, data_only=True)
    for ws in wb.worksheets:
        nome = ws.title if len(wb.worksheets) > 1 else arq_respostas.stem
        while nome in cursos:
            nome += '_'
        d = ler_folha(ws, nome)
        if d:
            cursos[nome] = d
            print(f'  -> {nome}: {"presencial" if d["t"] == "P" else "EAD"}, {d["n"]} respostas')

    if not cursos:
        sys.exit('Nenhum curso com respostas foi lido.')

    # Lê a planilha de informações se ela existir
    info_externa = ler_planilha_info(arq_info) if arq_info else {}
    if arq_info:
        print(f'Lendo metadados de texto de: {arq_info.name}')

    jf = AQUI / 'informacoes_cursos.json'
    info = json.loads(jf.read_text(encoding='utf-8')) if jf.is_file() else {}
    
    novos = []
    for k in cursos:
        if k not in info:
            # Cria a estrutura base
            info[k] = json.loads(json.dumps(SKELETON))
            # Se encontrou dados correspondentes na planilha de info, já injeta aqui!
            if k in info_externa:
                info[k]['nome'] = info_externa[k]['nome']
                info[k]['sigla'] = info_externa[k]['sigla']
                info[k]['finalidade'] = info_externa[k]['finalidade']
            novos.append(k)

    if novos or not jf.is_file():
        jf.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f'Adicionados a informacoes_cursos.json: {", ".join(novos)}')

    for k, d in cursos.items():
        i = json.loads(json.dumps(info[k]))
        for n in i.get('notas', []):
            n['imagem'] = uri(n.get('imagem', ''))
        d['i'] = i

    modelo = (AQUI / 'modelo_apresentacao.html').read_text(encoding='utf-8')
    dados = json.dumps(cursos, ensure_ascii=False).replace('</', '<\\/')
    saida.write_text(modelo.replace('__DATA__', dados).replace('__LOGO__', uri('logo.png')), encoding='utf-8')
    print(f'\nPronto: {saida} ({len(cursos)} cursos processados)')


if __name__ == '__main__':
    main()