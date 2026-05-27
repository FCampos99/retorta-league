#!/usr/bin/env python3
"""Retorta League — servidor local.
Actualiza o Excel quando reportas um jogo e regenera o website.
Uso: python3 server.py
"""

import json, os, re, socket
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime
import openpyxl

EXCEL_PATH = os.path.expanduser('~/Desktop/Retorta-League.xlsx')
HTML_PATH  = os.path.expanduser('~/Desktop/retorta-league/index.html')
PORT       = 8765

SEASON_SHEET_MAP = {
    'UCL_SS':     'UCL SS',
    'UCL_WS':     'UCL WS',
    'UCL_SS_2ed': 'UCL SS 2ed',
    'UCL_WS_2ed': 'UCL WS 2ed',
    'UCL_SS_3ed': 'UCL SS 3ed',
}

SEASONS_META = [
    ('UCL SS',     'Época de Verão',   '1.ª Edição', '2024',    '18/03/2024', '29/07/2024', False),
    ('UCL WS',     'Época de Inverno', '1.ª Edição', '2024-25', '05/09/2024', '28/01/2025', False),
    ('UCL SS 2ed', 'Época de Verão',   '2.ª Edição', '2025',    '06/03/2025', '07/07/2025', False),
    ('UCL WS 2ed', 'Época de Inverno', '2.ª Edição', '2025-26', '08/09/2025', '19/01/2026', False),
    ('UCL SS 3ed', 'Época de Verão',   '3.ª Edição', '2026',    '10/03/2026', '14/07/2026', True),
]


# ── HTTP Handler ────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._ok_cors()
        self.end_headers()

    def do_GET(self):
        if self.path == '/ping':
            self._ok_cors()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"ok":true}')

    def do_POST(self):
        if self.path == '/report':
            length = int(self.headers.get('Content-Length', 0))
            body   = self.rfile.read(length)
            try:
                data   = json.loads(body)
                result = process_report(data)
                self._ok_cors()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(result, ensure_ascii=False).encode())
            except Exception as e:
                self.send_response(500)
                self._cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}, ensure_ascii=False).encode())

    def _ok_cors(self):
        self.send_response(200)
        self._cors_headers()

    def _cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def log_message(self, fmt, *args):
        ts = datetime.now().strftime('%H:%M:%S')
        print(f'[{ts}] {fmt % args}')


# ── Excel update ────────────────────────────────────────────────────────────

def parse_pt_date(s):
    d, m, y = s.split('/')
    return datetime(int(y), int(m), int(d))

def process_report(data):
    """
    data = {
      season_id : 'UCL_SS_3ed',
      date      : '02/06/2026',   # DD/MM/YYYY
      winners   : ['Francisco Campos', ...],
      losers    : ['João Capela', ...],
    }
    """
    season_id = data['season_id']
    date_str  = data['date']
    winners   = set(data['winners'])
    losers    = set(data['losers'])

    sheet_name  = SEASON_SHEET_MAP[season_id]
    target_date = parse_pt_date(date_str)

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[sheet_name]

    # Find the column whose header matches the target date
    date_col = None
    for col in range(12, ws.max_column + 2):
        val = ws.cell(row=2, column=col).value
        if isinstance(val, datetime) and val.date() == target_date.date():
            date_col = col
            break

    if date_col is None:
        raise ValueError(f'Jornada {date_str} não encontrada na época {season_id}')

    # Walk player rows (col 1 = rank, col 2 = name)
    updated   = []
    not_found = list(winners | losers)

    for row in range(3, ws.max_row + 1):
        rank = ws.cell(row=row, column=1).value
        name = ws.cell(row=row, column=2).value
        if not isinstance(rank, (int, float)) or not isinstance(name, str):
            continue
        if name in winners:
            ws.cell(row=row, column=date_col).value = 'W'
            updated.append(f'{name}: W')
            if name in not_found: not_found.remove(name)
        elif name in losers:
            ws.cell(row=row, column=date_col).value = 'L'
            updated.append(f'{name}: L')
            if name in not_found: not_found.remove(name)

    wb.save(EXCEL_PATH)
    print(f'  Excel guardado — {len(updated)} células actualizadas')

    regenerate_html()

    return {'ok': True, 'updated': updated, 'not_found': not_found}


# ── Website regeneration ────────────────────────────────────────────────────

def extract_seasons():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    seasons = []

    for sheet_name, tipo, edicao, ano, inicio, fim, atual in SEASONS_META:
        ws   = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        match_dates = []
        for col in rows[1][11:]:
            if isinstance(col, datetime):
                match_dates.append(col.strftime('%d/%m/%Y'))
            elif col is not None:
                match_dates.append(str(col))

        # Count played jornadas (any column with ≥1 result)
        played_jornadas = 0
        for i in range(len(match_dates)):
            for row in rows[2:]:
                if len(row) > 11 + i and row[11 + i] in ('W', 'L'):
                    played_jornadas += 1
                    break
        played_jornadas = max(played_jornadas, 1)

        players = []
        for row in rows[2:]:
            if row[0] is None or row[1] is None:
                continue
            if not isinstance(row[0], (int, float)):
                continue

            results = [row[11 + i] if 11 + i < len(row) and row[11 + i] in ('W', 'L') else None
                       for i in range(len(match_dates))]

            vitorias = results.count('W')
            derrotas = results.count('L')
            jogos    = vitorias + derrotas
            if jogos == 0:
                continue

            win_rate  = round(vitorias / jogos, 4)
            factor    = round(jogos / played_jornadas, 4)
            media     = round(win_rate * factor, 4)

            players.append({
                'rank': int(row[0]),
                'nome': row[1],
                'alcunha': row[2] or row[1],
                'jogos': jogos,
                'vitorias': vitorias,
                'derrotas': derrotas,
                'winRate': win_rate,
                'mediaPonderada': media,
                'resultados': results,
            })

        # Re-rank by mediaPonderada
        players.sort(key=lambda p: -p['mediaPonderada'])
        rank = 1
        for i, p in enumerate(players):
            if i > 0 and p['mediaPonderada'] < players[i-1]['mediaPonderada']:
                rank = i + 1
            p['rank'] = rank

        seasons.append({
            'id': sheet_name.replace(' ', '_'),
            'label': f'{tipo} · {edicao}',
            'tipo': tipo,
            'edicao': edicao,
            'ano': ano,
            'inicio': inicio,
            'fim': fim,
            'atual': atual,
            'datas': match_dates,
            'jogadores': players,
        })

    return seasons


def regenerate_html():
    seasons  = extract_seasons()
    new_json = json.dumps(seasons, ensure_ascii=False)

    html = open(HTML_PATH, encoding='utf-8').read()
    new_html = re.sub(
        r'/\* RETORTA_DATA_BEGIN \*/.*?/\* RETORTA_DATA_END \*/',
        f'/* RETORTA_DATA_BEGIN */{new_json}/* RETORTA_DATA_END */',
        html, flags=re.DOTALL
    )
    open(HTML_PATH, 'w', encoding='utf-8').write(new_html)
    print(f'  index.html regenerado — {len(seasons)} épocas')


# ── Entry point ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print('═' * 50)
    print('  Retorta League — Servidor Local')
    print(f'  Excel  : {EXCEL_PATH}')
    print(f'  Website: {HTML_PATH}')
    print('═' * 50)
    print('  A gerar dados iniciais...')
    regenerate_html()

    server = HTTPServer(('localhost', PORT), Handler)
    server.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    print(f'  Servidor a correr em http://localhost:{PORT}')
    print('  Abre o index.html e reporta jogos!')
    print('  Ctrl+C para parar.')
    print('═' * 50)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  Servidor parado.')
