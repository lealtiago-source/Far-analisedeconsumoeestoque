Nome do Projeto
Análise de Consumo e Estoque de Medicamentos - Versão Beta (em testes)
Descrição
Este projeto realiza a análise automatizada de dados de dispensação, distribuição e estoque de medicamentos a partir de três planilhas Excel (.xlsx), gerando um relatório consolidado com previsões, alertas de faltas e consumo médio ajustado.
Pontos Críticos da Análise
Nomes de Colunas nas Planilhas
Certifique-se de que os arquivos estejam em .xlsx (ultima versão do Excel)  e que as colunas estejam exatamente com os nomes:
Dispensação.xlsx: Data Dispensação, Medicamento/Produto, Lote, Quantidade Dispensada

Distribuição.xlsx: Data Distribuição, Medicamento/Produto, Lote, Quantidade distribuída (unidades)

Estoque.xlsx: Medicamento/Produto, Lote, Quantidade em Estoque, Validade


No Sigaf, baixe os relatórios:
Relatório de Dispensação por Medicamento – o site só deixa baixar todas dispensações de 1 mês (baixe relatórios em formato .csv dos últimos seis meses e unifique-os por ordem de data de dispensação). Em seguida exclua todas as colunas diferentes das descritas acima para o arquivo Dispensação, (vá em Salvar Como) e salve como Pasta de Trabalho do Excel.

Relatório de Distribuição – busque somente as unidades de saúde que você distribui medicamentos e queira que essas distribuições sejam contabilizadas no consumo do medicamento. Baixe o arquivo .CSV e em seguida exclua todas as colunas diferentes das descritas acima para o arquivo Distribuição, (vá em Salvar Como) e salve como Pasta de Trabalho do Excel.

Posição atual do Estoque – insira somente o nome da unidade de saúde, clique em buscar e baixe o arquivo .CSV. Exclua todas as colunas diferentes das descritas acima para o arquivo Estoque, (vá em Salvar Como) e salve como Pasta de Trabalho do Excel.



Normalização dos nomes
Nomes equivalentes como "comprimido", "cápsula", "drágea" são agrupados automaticamente.
Cálculo de Consumo
Baseado apenas nos dias em que houve dispensação.
Análise de Faltas
Se um medicamento trocou de lote na dispensação e houve um espaço de 15 dias ou mais entre essas dispensações, é contabilizados quantos dias de falta ocorreu no período analisado. (cuidado com medicamentos com baixo consumo mensal médio corrigido pois os dias de falta são superestimados devido as poucas dispensações do medicamento)
Se um medicamento parou de ser dispensado e não tem estoque dele, é mostrado no estoque qual o dia da última dispensação.
Previsão de Fim do Estoque e Validade
A previsão de fim do estoque é estimada com base no consumo médio diário e estoque atual.
A validade apresentada é do lote com data mais próxima.
Se sua validade ficar lilás – sua Previsão fim do estoque ultrapassa a validade – quer dizer, doe ou troque este medicamento (cuidado pois para medicamentos com dois lotes, a validade mostrada é a mais próxima, portanto, use-a primeiro e verifique a validade dos outros lotes e a previsão do fim do estoque)
Se sua validade ficar amarela – sua Previsão fim do estoque é que o medicamento acabe em até 4 meses antes da data de validade (fique tranquilo)
Se sua Previsão fim do estoque ficar vermelho – ATENÇÃO seu estoque acaba em até 4 meses, já é hora de fazer compras.
Como Usar
Acesse o site:

Faça upload dos três arquivos .xlsx.
Após o processamento, baixe o arquivo resultado_analise.xlsx pelo link apresentado.