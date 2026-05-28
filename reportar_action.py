#!/usr/bin/env python3
"""Corre no GitHub Actions para actualizar o Excel e regenerar o site."""
import json, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.environ.setdefault('SKIP_GIT_PUSH', '1')

from server import process_report, regenerate_html

season_id = os.environ['SEASON_ID']
date_str  = os.environ['JORNADA_DATE']
winners   = json.loads(os.environ['WINNERS_JSON'])
losers    = json.loads(os.environ['LOSERS_JSON'])

print(f'📋 Jornada {date_str} · {season_id}')
print(f'   Vitória: {winners}')
print(f'   Derrota: {losers}')

result = process_report({'season_id': season_id, 'date': date_str, 'winners': winners, 'losers': losers})

if result['ok']:
    print(f'✅ {len(result["updated"])} jogadores actualizados')
    if result['not_found']:
        print(f'⚠️  Não encontrados: {", ".join(result["not_found"])}')
else:
    print(f'❌ {result.get("error")}')
    sys.exit(1)
