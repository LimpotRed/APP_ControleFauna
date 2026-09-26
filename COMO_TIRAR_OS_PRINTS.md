# Protótipo de Monitoramento Ambiental

Este protótipo funciona inteiramente no navegador. Ele não usa API, conta Google, banco de dados, GPS real nem envio de informações.

## Como abrir

1. Abra o arquivo [index.html](index.html) com o navegador.
2. Use o menu lateral para trocar de tela.
3. Faça as capturas com a janela maximizada.

Também é possível abrir uma tela específica acrescentando um destes trechos ao fim do endereço do arquivo:

- #inicio — painel inicial;
- #fauna — formulário de ocorrência de fauna;
- #flora — formulário de inspeção de flora;
- #mapa — mapa ilustrativo;
- #registros — lista de registros;
- #pendencias — fila de revisão.

## Prints recomendados

1. Painel inicial: indicadores, gráfico, pendências e ações de coleta.
2. Nova ocorrência de fauna: GPS, fotografia, espécie, grupo, condição e entorno.
3. Nova inspeção de árvore: DAP, estado fitossanitário, risco de queda e recomendação.
4. Mapa de registros: pinos de fauna e flora com legenda.
5. Registros de campo: visão de consulta do escritório.

Os números, coordenadas, fotografias ilustrativas e registros são fictícios e existem apenas para demonstrar a atividade acadêmica.

A [planilha-base](MONITORAMENTO_AMBIENTAL_APPSHEET.xlsx) é opcional. Ela pode ser mostrada junto com os prints para evidenciar a estrutura dos dados, mas não é necessária para abrir o protótipo.

O arquivo gerar_planilha_appsheet.ps1 gera novas cópias da planilha-base caso você precise alterar os exemplos.
