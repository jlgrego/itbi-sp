# Preço do m² em São Paulo, por bairro

Site com o valor declarado por m² de área construída, a partir das guias de ITBI pagas publicadas pela Secretaria Municipal da Fazenda.

## Como funciona

- As planilhas de 2025 e 2026 ficam anexadas a Releases deste repositório (`dados-2025` e `dados-2026`).
- O workflow `atualizar.yml` roda todo mês. Ele baixa as planilhas, aplica as regras, gera as páginas e publica o site no GitHub Pages.
- O script que faz tudo é `scripts/gerar_site.py`. O modelo das páginas de bairro é `scripts/modelo_bairro.html`.

## Atualizar a planilha de 2026

1. Baixe o arquivo novo pelo link oficial da Fazenda.
2. Renomeie para `itbi_2026.xlsx`.
3. Em **Releases**, edite o Release `dados-2026`, remova o anexo antigo e envie o novo.
4. Rode o workflow **Atualizar site** em **Actions**, ou espere a data agendada.

## Regras aplicadas

- Somente compra e venda (natureza 1).
- Somente imóveis ativos e prediais. Apartamentos de condomínio entram com a área da unidade.
- Um registro por cadastro: o mais recente.
- Área construída entre 20 e 1.000 m².
- Preço por m² entre R$ 500 e R$ 100 mil. Valores fora dessa faixa costumam ser erro de digitação.
- Bairros com menos de 30 transações não ganham página.

## Limitações

- Transações rurais e as pagas pelo PPI não estão na base.
- Valor declarado não é avaliação oficial.
