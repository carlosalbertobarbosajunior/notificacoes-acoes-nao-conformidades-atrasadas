# notificacoes-acoes-nao-conformidades-atrasadas
## Script de notificações para os responsáveis pelas ações de nao conformidades que estão atrasadas


### Problemática:
Notificar os responsáveis pelas não conformidades que estes possuem ações atrasadas que devem ser realizadas.


### View em T-SQL: (view-notificacoes-nao-conformidade.sql)
Tem a finalidade de criar o relacionamento entre as tabelas das ações a serem tomadas e as não conformidades. Também filtra não conformidades já concluídas ou implementadas (corrigidas), além de não considerar ações que já possuem data de implementação.


### Script Python: (notificacoes_acoes_nao_conformidade_atrasadas.py)
O script é dividido em três etapas:
#### Excel:
Abre o arquivo com conectividade com o banco de dados, atualiza seus vínculos, salva e encerra a instância.
#### Pandas:
Extrai as informações das ações e os e-mails dos responsáveis. Classifica as ações atrasadas em graus 1, 2 e 3 de acordo com os dias de atraso.
Por fim, utiliza o campo de responsável para identificar qual e-mail deve ser enviado a notificação.
#### Outlook:
Executa uma instância outlook e dispara os e-mails previamente guardados em variáveis de acordo com o grau de atraso.
