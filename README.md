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
- A tela informa quantos registros repetidos foram removidos; identificadores vindos do Excel como `10.0` e `1.0` são normalizados corretamente antes da comparação.
- Registros sem `NUMFUNC` ou `NUMVINC` não são agrupados artificialmente em uma única chave; o sistema mantém esses registros separados e exibe um alerta para conferência.
- Painel gerencial com filtros por setor, vínculo e especialidade.
- Indicadores de médicos únicos, escalas/registros, setores e horas escaladas. A chave oficial para médico/vínculo é a coluna `NUMFUNC-NUMVINC`; identificadores incompletos recebem uma chave técnica distinta por linha para não agrupar registros sem confirmação.
- Análise de capacidade por setor, com horas escaladas, carga contratual, saldo e percentual de cobertura.
- Relatório de horas escaladas por especialidade, com médicos, escalas, carga contratual, saldo e percentual de cobertura, disponível na tela e no Excel. Se a entrada não possuir `CARGA HORARIA ESCALADA`, o relatório usa `CARGA HORARIA` como fallback.
- Distribuição por especialidade e tipo de vínculo, além de identificação de setores com maior concentração.
- Alertas de qualidade para CPF, setor, especialidade, carga horária e vínculos não classificados.
- Resumo por setor com médicos, vínculos e especialidades.
- Tabela de especialidades por setor com quantidade de médicos.
- Resumo no padrão do relatório de referência: uma linha total por setor e linhas detalhadas por especialidade, com Efetivos, Contrato Temporário, Requisitado e Total Geral.
- O resumo termina com uma linha **TOTAL GERAL**, somando todos os setores e especialidades.
- Excel com quatro abas mais a aba opcional **Servidores unicos**, A4 paisagem e uma página de largura.
- PDF estatístico no mesmo formato hierárquico do relatório de referência, com uma segunda página de horas escaladas por especialidade; a exportação usa os filtros selecionados na tela.
- `NUMFUNC`, `NUMVINC`, `ESCALA` e `ORDEM` são exportados como números inteiros, e o CPF é apresentado no formato `000.000.000-00`.

## Execução

```bash
pip install -r requirements.txt
streamlit run app.py
```

O leitor aceita CSV, XLS, XLSX e XLSM. Se um arquivo tabular tiver sido exportado com extensão Excel incorreta, o sistema tenta automaticamente o formato alternativo e, em último caso, lê como texto delimitado.
