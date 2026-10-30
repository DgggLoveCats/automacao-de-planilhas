# 📊 Consolidador de Planilhas de Vendas

Automação em Python que lê várias planilhas de vendas, **cada uma em um formato diferente**, limpa os dados e gera um único relatório em Excel pronto para usar, com indicadores, gráficos e uma lista de tudo o que foi descartado.

## O problema

Pequenas empresas costumam receber relatórios de vendas de várias fontes (vendedores, filiais, sistemas diferentes): colunas com nomes variados, datas em formatos diferentes, preços como `R$ 1.299,90`, linhas duplicadas e células vazias. Juntar tudo à mão leva horas e gera erros.

## A solução

Basta colocar as planilhas em uma pasta e rodar um comando. Em segundos sai um relatório consolidado.

**O que o script faz**

- Lê todos os arquivos `.xlsx` e `.csv` da pasta (todas as abas de cada arquivo).
- Reconhece nomes de colunas alternativos (`Qtd`, `QTDE`, `Quantidade`; `Item`, `Produto`...), ignorando acentos e maiúsculas.
- Detecta sozinho o separador e o encoding dos CSVs.
- Converte datas (`15/03/2026`, `2026-03-15`, datas do Excel) e valores (`R$ 1.299,90`, `89,90`, `150`).
- Padroniza textos (`"  mOUSE "` → `Mouse`) e remove duplicatas.
- Valida cada linha e registra o motivo de cada descarte na aba **Problemas**, então nada some em silêncio.
- Gera o relatório com fórmulas reais (`SUMIFS`, `SUM`), que recalculam se os dados mudarem.

**O relatório (`saida/relatorio_vendas.xlsx`)**

| Aba | Conteúdo |
|---|---|
| Resumo | Faturamento, itens vendidos, nº de vendas, ticket médio, período |
| Por Produto / Por Vendedor / Por Mês | Totais, % do total e gráfico de barras |
| Dados | Todas as vendas consolidadas, com filtro |
| Problemas | Linhas descartadas e o motivo |

## Como usar

```bash
# 1. Instale as dependências
pip install -r requirements.txt

# 2. (Opcional) gere planilhas de exemplo bagunçadas para testar
python gerar_dados_exemplo.py

# 3. Consolide
python consolidador.py
```

Opções:

```bash
python consolidador.py --entrada minha_pasta --saida resultado.xlsx
python consolidador.py --manter-duplicados   # não remove linhas idênticas
```

## Estrutura

```
consolidador.py          # script principal
gerar_dados_exemplo.py   # cria planilhas de exemplo em dados_entrada/
test_consolidador.py     # testes automatizados (pytest)
requirements.txt
```

## Adaptando para outro cliente

- **Outros nomes de coluna:** acrescente os apelidos no dicionário `APELIDOS` em `consolidador.py`.
- **Outros campos** (categoria, filial, desconto): inclua em `COLUNAS_PADRAO` e nas abas do relatório.
- **Outras regras de validação:** edite a função `validar_linha`.

## Observações

- Linhas com todos os campos iguais (data, produto, vendedor, quantidade e preço) são tratadas como duplicadas. Use `--manter-duplicados` se vendas idênticas forem possíveis no seu caso.
- Valores sem vírgula, como `1.299`, são lidos como número com ponto decimal. Para evitar ambiguidade, mantenha a vírgula nos decimais.
- Ao abrir o relatório no Excel, as fórmulas são calculadas automaticamente.

## Testes

```bash
pip install pytest
pytest
```

## Tecnologias

Python 3.9+, pandas, openpyxl.
