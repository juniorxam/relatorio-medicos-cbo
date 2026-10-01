# Relatório de Médicos por Setor

Aplicativo Streamlit independente para gerar a relação de médicos da planilha da Intranet e consolidar as especialidades por setor.

## Critério de seleção

São considerados somente registros cujo campo `CARGO`, após normalização de acentos e caixa, seja exatamente um destes:

- `MÉDICO`
- `MEDICO - RQE`
- `MEDICO CLINICO`

A especialidade é obtida de `OCUPACAO`; quando vazia, aparece como `Não informada`.

## Saídas

- Relação detalhada de médicos por setor.
- Relação opcional de servidores sem repetição, usando `NUMFUNC + NUMVINC` como chave e mantendo as colunas `ORDEM`, `NUMFUNC`, `NUMVINC`, `SERVIDOR`, `SETOR`, `CARGO`, `OCUPACAO` e `CARGA HORARIA`.
- Resumo por setor com médicos, vínculos e especialidades.
- Tabela de especialidades por setor com quantidade de médicos.
- Resumo no padrão do relatório de referência: uma linha total por setor e linhas detalhadas por especialidade, com Efetivos, Contrato Temporário, Requisitado e Total Geral.
- Excel com quatro abas mais a aba opcional **Servidores unicos**, A4 paisagem e uma página de largura.
- PDF estatístico no mesmo formato hierárquico do relatório de referência.
- `NUMFUNC`, `NUMVINC`, `ESCALA` e `ORDEM` são exportados como números inteiros, e o CPF é apresentado no formato `000.000.000-00`.

## Execução

```bash
pip install -r requirements.txt
streamlit run app.py
```

O leitor aceita CSV, XLS, XLSX e XLSM. Se um arquivo tabular tiver sido exportado com extensão Excel incorreta, o sistema tenta automaticamente o formato alternativo e, em último caso, lê como texto delimitado.
