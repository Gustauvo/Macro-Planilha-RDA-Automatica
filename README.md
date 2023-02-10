# Macro-Planilha-RDA-Automatica
VBA EXCEL com SAP GUI

Esta macro usa o VBA da planilha para fazer o download de uma planilha Excel que tem as requisições pendentes no SAP GUI.
A macro usa a transação ME5A para baixar as requisições(RDA) que estão pendentes para fazer pedido de compras em uma planilha Excel e mostra a quantidade de e-mails que chegou no outlook de cada requisição.
O tempo que era gasto para fazer essa atividade era muito grande, sendo que é uma atividade repetitiva e que na empresa em que trabalho faziamos semanalmente, essa macro me ajudou a ganhar muito tempo.

Passo a passo de como usar:
1. Abra o SAP.
2. Clique no botão "Baixar as RDA's pendentes."
3. Preencha o campo Login e Senha, com o seu login e senha do SAP ou clique em "Já estou logado", caso já esteja logado no SAP.
4. Clique em OK.
5. Depois que o Excel for baixado e formatado, tem que salvar a planilha para não perder a formatação.

Obs:
Muito provavelmente o código dara algum erro, pois o SAP pode ser usado com várias configurações prévias diferentes ou que necessite também de informações diferentes para baixar a planilha.
Por causa disso deixei o código bem explicado no VBA, escrevi o que cada parte está fazendo, então caso for necessário alguma mudança que acredito que seja pequena, está de fácil acesso a todos.
IMPORTANTE: Segue a imagem do padrão de layout que uso na transação do ME5A. Muito importante que o layout seja o mesmo para dar certo o código!
