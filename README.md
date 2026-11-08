# Painel de presença · LogMax

Tela que mostra, a cada 30 segundos, quantos motoboys chegaram, quem está sem OK, as faltas e as vagas de cada loja.

- `index.html`: a tela (abre no navegador; publicada pelo GitHub Pages).
- `codigo-da-planilha.gs`: código colado no Apps Script da planilha "ESCALAS LOGMAX - FOOD", implantado como App da Web. Entrega só loja, data, horário, código, nome e presença (a tela mostra de hoje até domingo).

## Tela do operador (marcar chegada)
- `operador.html`: tela com senha onde o operador clica em **Chegou** e o OK vai direto para a coluna PRESENÇA da planilha. Também dá para **Desfazer**.
- Para testar com a cópia da planilha: `operador.html?api=ENDERECO_EXEC_DA_COPIA` (aparece a faixa amarela MODO TESTE).
- O código da gravação fica num arquivo separado no Apps Script: `codigo-operador.gs` (nomes com OP_/op, não mexe no `codigo-da-planilha.gs` do painel).
- A senha fica no Apps Script em Configurações do projeto → Propriedades do script → `SENHA_EDICAO`.
