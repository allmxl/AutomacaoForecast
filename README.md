# AutomacaoForecast
Projeto de automação para formatação de planilhas

## Objetivo
Automatizar a atualização da planilha BD a partir de arquivos Excel recebidos mensalmente (restritos e irrestritos), adicionando novas colunas.

## Estrutura do Projeto
- `input/` → arquivos Excel recebidos
- `output/` → arquivos processados (opcional)
- `scripts/` → scripts Python
- `BD.xlsx` → planilha principal

## Como usar
1. Coloque os arquivos Excel na pasta `input/`.
2. Execute o script `scripts/automatizacao_bd.py`.
3. A planilha `BD.xlsx` será atualizada com os dados processados.

Estrutura final: quais serão as colunas definitivas (exatamente nessa ordem)?
Base | Ciclo | Classificação | Ano | Mês | Código | Descrição | Produto | Marca | Fornecimento | Categoria | Soma de Qtd | Soma de Valor
Estrutura inicial: nesta ordem:Tabela dinâmica 
(iniciado na linha 4) 
Material | Código | Marca | Submarca  | CategoriaMkt | Lançamento | Classificação 
(iniciado na linha 3) 
Prev.Vendas(Qtde) | Prev.vendas (R$)
(iniciado na linha 4)
Set.2025 | out.2025 | nov.2025 | dez.2025 | 2025 | ...próximos 12 meses (A coluna N de total 2025 deve ser ignorada) 


Origem dos dados:
Orçamento → já está no BD, não mexemos.
Fornecedimento → vem do DRE. 
Realizado → vem do DRE.
Forecast Restrito/Irrestrito → vem dos arquivos da pasta ForecastAutomacao 

A linha 3 deve concatenar com a linha 4 para formar o OBJETIVO|MÊS (Prev.Vendas(Qtde) Set.2025) -- Vai ser os dados desta coluna que será o novo Soma de Qtd 
Nestas planilhas tem do mês atual até o próximo ano (deve ser lido tudo) 

Período: 
exemplo -- 
NO BD 

Forecast Abril Irrestrito | Ciclo 3 Irrestrito (Abril) | Irrestrito | 2025 | abr | 5008 | Anti-mais | 5008\Anti-mais | Anti-Septico | Prod.Interno | Dermatológico | 146 | 190
Forecast Abril Irrestrito | Ciclo 3 Irrestrito (Abril) | Irrestrito | 2025 | mai | 5008 | Anti-mais | 5008\Anti-mais | Anti-Septico | Prod.Interno | Dermatológico | 160 | 290
Forecast Abril Irrestrito | Ciclo 3 Irrestrito (Abril) | Irrestrito | 2025 | jun | 5008 | Anti-mais | 5008\Anti-mais | Anti-Septico | Prod.Interno | Dermatológico | 500 | 900

NA PLANILHA com tabela dinâmica
linha 3                                                                             Prev.vendas (Qtde)  | Prev.vendas (Qtde) ....   Prev.vendas (R$) |   Prev.vendas (R$)
linha 4                                                                                    Set.2025     |     Out.2025       ....      Set.2025      |      Out.2025
linha 5 Anti-mais | 5008 | Anti-Septico | Anti | Dermatológico |    |  Portfólio         146                   160                         900                290

----------------------------------------------------------------------------------------------------------------------------------------------------------------

Leitura de DRE deu certo-- formatação realizada! 

Problema atual: Puxa todos os meses e anos do realizado e preciso que seja puxado que mês e ano especificos que estou rodando o sistema. 

Para segunda>> Ajuste de meses do realizado, ajuste de todo o forecast 
