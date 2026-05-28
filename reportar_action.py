#!/usr/bin/env python3
"""Corre no GitHub Actions. Suporta dois modos:
  - Texto livre: TEXTO preenchido → Claude API interpreta quem ganhou/perdeu
  - Estruturado: WINNERS_JSON + LOSERS_JSON preenchidos directamente
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault('SKIP_GIT_PUSH', '1')

from server import process_report, regenerate_html, SEASONS_META, SEASON_SHEET_MAP, EXCEL_PATH
import openpyxl
from datetime import datetime

season_id  = os.environ['SEASON_ID']
date_str   = os.environ['JORNADA_DATE']
texto      = os.environ.get('TEXTO', '').strip()
winners    = json.loads(os.environ.get('WINNERS_JSON', '[]'))
losers     = json.loads(os.environ.get('LOSERS_JSON', '[]'))

# ── Modo texto: usar Claude API para interpretar ──────────────────────────
if texto:
    import anthropic

    # Carregar lista de jogadores da época actual
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet_name = SEASON_SHEET_MAP.get(season_id, 'UCL SS 3ed')
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    players = []
    for row in rows[2:]:
        if isinstance(row[0], (int, float)) and isinstance(row[1], str):
            players.append(f"{row[1]} (alcunha: {row[2] or row[1]})")

    client = anthropic.Anthropic()
    response = client.messages.create(
        model='claude-haiku-4-5-20251001',
        max_tokens=512,
        system=f"""És o assistente da Retorta League (futebol de 5).
Época: {season_id} · Data: {date_str}

Jogadores conhecidos:
{chr(10).join(f'  • {p}' for p in players)}

Quando o utilizador descrever um resultado, usa a ferramenta para registar os nomes completos.""",
        tools=[{
            "name": "registar",
            "description": "Registar vencedores e derrotados",
            "input_schema": {
                "type": "object",
                "properties": {
                    "winners": {"type": "array", "items": {"type": "string"}, "description": "Nomes completos dos vencedores"},
                    "losers":  {"type": "array", "items": {"type": "string"}, "description": "Nomes completos dos derrotados"}
                },
                "required": ["winners", "losers"]
            }
        }],
        messages=[{"role": "user", "content": texto}]
    )

    for block in response.content:
        if block.type == 'tool_use':
            winners = block.input['winners']
            losers  = block.input['losers']
            print(f'🤖 Claude interpretou:')
            print(f'   Vitória: {winners}')
            print(f'   Derrota: {losers}')
            break
    else:
        print('❌ Claude não conseguiu interpretar o resultado.')
        sys.exit(1)

# ── Actualizar Excel ──────────────────────────────────────────────────────
if not winners or not losers:
    print('❌ Sem jogadores definidos.')
    sys.exit(1)

print(f'📋 {date_str} · {season_id}')
result = process_report({'season_id': season_id, 'date': date_str, 'winners': winners, 'losers': losers})

if result['ok']:
    print(f'✅ {len(result["updated"])} jogadores actualizados')
    if result['not_found']:
        print(f'⚠️  Não encontrados no Excel: {", ".join(result["not_found"])}')
        print('   (novos jogadores não são adicionados automaticamente)')
else:
    print(f'❌ {result.get("error")}')
    sys.exit(1)
