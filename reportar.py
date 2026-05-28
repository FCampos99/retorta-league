#!/usr/bin/env python3
"""
Retorta League — Reporte de Jogo via Claude API
Uso: python3 reportar.py "Campos, Quim e Rodrigo ganharam. Perderam: Capela, Renato e Vieira"
     python3 reportar.py          (modo interactivo)
"""

import os, sys, json
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from server import process_report, SEASONS_META, EXCEL_PATH
import openpyxl
import anthropic

client = anthropic.Anthropic()


def get_context():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    meta = next(m for m in SEASONS_META if m[6])   # atual = True
    sheet_name, tipo, edicao = meta[0], meta[1], meta[2]
    ws   = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    players = []
    for row in rows[2:]:
        if isinstance(row[0], (int, float)) and isinstance(row[1], str):
            nick = row[2] or row[1]
            players.append(f"{row[1]} (alcunha: {nick})")

    dates = []
    for col in rows[1][11:]:
        if isinstance(col, datetime):
            dates.append(col.strftime('%d/%m/%Y'))

    return {
        'season_id':    sheet_name.replace(' ', '_'),
        'season_label': f'{tipo} · {edicao}',
        'players':      players,
        'dates':        dates,
        'today':        datetime.now().strftime('%d/%m/%Y'),
    }


TOOL = {
    "name": "registar_resultado",
    "description": "Regista o resultado de um jogo da Retorta League: actualiza o Excel e publica o website.",
    "input_schema": {
        "type": "object",
        "properties": {
            "season_id": {
                "type": "string",
                "description": "ID da época (ex: UCL_SS_3ed)"
            },
            "date": {
                "type": "string",
                "description": "Data da jornada no formato DD/MM/YYYY"
            },
            "winners": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Nomes completos dos jogadores da equipa vencedora"
            },
            "losers": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Nomes completos dos jogadores da equipa perdedora"
            },
            "confirmacao": {
                "type": "string",
                "description": "Frase curta a confirmar com o utilizador antes de gravar (ex: 'Jornada 02/06/2026 — Campos, Quim e Rodrigo ganharam contra Capela, Renato e Vieira')"
            }
        },
        "required": ["season_id", "date", "winners", "losers", "confirmacao"]
    }
}


def parse_report(text, ctx):
    system = f"""És o assistente da Retorta League, uma liga de futebol de 5 entre amigos em Portugal.

Época actual: {ctx['season_label']} (ID: {ctx['season_id']})
Data de hoje: {ctx['today']}

Jogadores conhecidos nesta época:
{chr(10).join(f'  • {p}' for p in ctx['players'])}

Jornadas disponíveis nesta época:
{', '.join(ctx['dates'])}

Quando o utilizador descrever um resultado, usa a ferramenta registar_resultado.
Usa sempre os nomes completos conforme a lista acima (não as alcunhas).
Se não souberes a data exacta, usa a próxima jornada ainda sem resultado.
Se um nome não estiver na lista, inclui-o como foi mencionado (pode ser jogador novo).
Responde sempre em português de Portugal."""

    response = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=1024,
        system=system,
        tools=[TOOL],
        messages=[{"role": "user", "content": text}]
    )

    for block in response.content:
        if block.type == "tool_use":
            return block.input

    # Sem tool use — Claude pediu esclarecimentos
    for block in response.content:
        if hasattr(block, 'text'):
            print(f"\nClaude: {block.text}")
    return None


def main():
    if len(sys.argv) > 1:
        text = ' '.join(sys.argv[1:])
    else:
        print("=" * 52)
        print("  Retorta League — Reporte de Jogo")
        print("=" * 52)
        print("Descreve o resultado (ex: 'Campos, Quim e Rodrigo")
        print("ganharam. Perderam: Capela, Renato e Vieira'):")
        text = input("> ").strip()
        if not text:
            return

    print("\n⚽ A processar com Claude...")
    ctx  = get_context()
    data = parse_report(text, ctx)

    if not data:
        print("❌ Não foi possível extrair o resultado. Sê mais específico.")
        return

    print(f"\n📋 {data['confirmacao']}")
    print(f"   Vitória : {', '.join(data['winners'])}")
    print(f"   Derrota : {', '.join(data['losers'])}")
    print(f"   Jornada : {data['date']}")

    resp = input("\nConfirmar? [s/N] ").strip().lower()
    if resp != 's':
        print("Cancelado.")
        return

    print("\n📊 A actualizar Excel e website...")
    result = process_report(data)

    if result['ok']:
        print(f"✅ {len(result['updated'])} jogadores actualizados")
        if result['not_found']:
            print(f"⚠️  Não encontrados no Excel (novos jogadores?): {', '.join(result['not_found'])}")
        print("🌐 https://fcampos99.github.io/retorta-league/")
    else:
        print(f"❌ Erro: {result.get('error')}")


if __name__ == '__main__':
    main()
