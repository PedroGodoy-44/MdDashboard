# Mais Delivery — BI direto na API

Script único,  Puxa pedidos da API, guarda em cache SQLite e gera dashboard + Excel.

## Instalação (PC novo)

```bash
pip install requests pandas openpyxl
```

Coloque o arquivo **`_env`** (suas credenciais, o mesmo do projeto) na **mesma pasta** do script.

## Uso

```bash
# 1) PREVIEW na hora, sem tocar na API (dados sintéticos):
python bi_maisdelivery.py --mock

# 2) Rodada real (1ª vez puxa ~180 dias; depois o cache deixa rápido):
python bi_maisdelivery.py

# Outras opções:
python bi_maisdelivery.py --desde 2026-04-01 --ate 2026-06-28
python bi_maisdelivery.py --estabs 82229,86759      # só alguns estabelecimentos
python bi_maisdelivery.py --no-fetch                # só recalcula do cache
python bi_maisdelivery.py --ciclo 35 --ocupacao 0.90 # ajusta o modelo de motoboys
```

## Saídas (na pasta atual, ou `--out`)

| Arquivo | O que é |
|---|---|
| `dashboard.html` | Painel interativo — abra no navegador. Filtro por estabelecimento, cards, gráficos, heatmap, ranking, produtos, clientes, dimensionamento. |
| `bi_maisdelivery.xlsx` | Uma aba por métrica + base de pedidos. |
| `cache.sqlite` | Persistência local (consultável com qualquer cliente SQLite). |
| `_amostra_pedido.json` | 1 pedido cru da API — serve pra mapear campos que não foram detectados. |

## Como funciona

- **Detecção automática de campos:** o script descobre sozinho os nomes reais no JSON da API
  (data, valor, situação, cliente, produtos…). No terminal ele imprime o que achou; no dashboard,
  chips ✓/✗ mostram o que está disponível.
- **Cache incremental:** cada estabelecimento tem um *watermark*; rodadas seguintes só puxam o que
  mudou (com 3 dias de sobreposição pra pegar atualização de status).
- **Dimensionamento de motoboys = MODELO**, não dado medido: a API legada não traz o ciclo real de
  entrega por entregador. Estima a partir de volume/hora × ciclo médio ÷ ocupação. Para números
  reais, é preciso capturar aceite/retirada/entrega por motoboy (ver `BI_PLANO_E_GAPS.md`).

## Métricas cobertas

Gerais (hoje/semana/mês/ano + crescimento + ticket + receita), distribuição por hora/dia/mês,
heatmap hora×dia, horários de pico, útil×fim de semana, ranking de estabelecimentos, cancelamentos,
top produtos, clientes (recorrência, ticket, top), dimensionamento por hora.
