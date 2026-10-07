Olá! Sou a Camilla, da Tyton Capital. Quero colocar no ar o clipping de notícias dos emissores da carteira, enviado de **tyton@tytoncapital.com.br** para **asset@tytoncapital.com.br**, rodando **neste computador** (Windows, com o Outlook desktop aberto e a caixa tyton@ configurada).

## O que já existe

1. O código está pronto no GitHub: repositório **ctolusso/teste**, branch **claude/blissful-pasteur-fsdw3r**, arquivos `clipping.py` e `requirements.txt`. Clone essa branch.
2. O que o `clipping.py` faz hoje:
   1. Busca no Google News (RSS) as notícias publicadas hoje, no horário de Brasília, para cerca de 40 emissores (dicionário `EMISSORES`). Faz uma busca geral e outra restrita a sites financeiros.
   2. Filtra relevância, valida os links e encurta no TinyURL.
   3. Remove notícias repetidas dentro do envio e entre envios (histórico em `dados/historico_clipping.json`). Limita a 6 notícias por emissor.
   4. Acumula as notícias do dia em `dados/Noticias_Diarias.xlsx` (tabela "Noticias_diarias") e manda a planilha em anexo.
   5. Monta o e-mail em HTML no visual da Tyton: faixa verde (#2a3e32) com "TYTON CAPITAL", "Prezados," com um resumo e uma seção numerada por emissor.
   6. Se não houver notícia nova, não envia nada.
3. Modo de teste: `python clipping.py --previa` gera `dados/previa.html` sem enviar e sem gravar histórico. A coleta leva uns 3 minutos. Já testei com notícias reais e funcionou (cerca de 35 notícias de 13 emissores).
4. O envio hoje está na função `enviar_email()` e usa o Microsoft Graph (app do Azure com as variáveis `CLIPPING_TENANT_ID`, `CLIPPING_CLIENT_ID` e `CLIPPING_CLIENT_SECRET`). Isso parou porque depende da TI registrar o app. **Por isso agora o envio vai sair pelo Outlook deste computador.**

## O que preciso que você faça

1. **Backup antes de tudo.** Guarde uma cópia do `clipping.py` atual antes de alterar.
2. **Trocar o envio para o Outlook local** (pywin32 / win32com), enviando pela conta tyton@tytoncapital.com.br. Use `SendUsingAccount`; se a tyton@ for caixa compartilhada, use `SentOnBehalfOfName`. Mantenha o HTML e o anexo do Excel. O código do Graph pode sair.
3. **Teste com o Pedro.** Crie uma opção de teste (por exemplo `--teste`) que:
   1. envia um único e-mail só para **pedro.recco@tytoncapital.com.br**;
   2. **não grava** no histórico, para não "gastar" as notícias do envio real.
   Me mostre a prévia antes de disparar.
4. **Agendar**, só depois da minha aprovação, no Agendador de Tarefas do Windows: **dias úteis, a cada 2 horas, das 7h às 19h** (7h, 9h, 11h, 13h, 15h, 17h e 19h), com destino asset@tytoncapital.com.br. Como o Outlook precisa da sessão do usuário aberta, configure para rodar com o usuário logado e me avise que o computador precisa ficar ligado nesses horários.
5. **Onde salvar os dados.** O código antigo salvava o Excel e o histórico em
   `C:\TYT\Tyton Capital\Tyton - Credito Privado\Credito\Relatórios e Controle\Controle PowerBI\Base de Dados\Pedro R\`
   porque essa planilha alimenta o Power BI. Me pergunte se quero voltar a salvar lá. Atenção: a planilha nova tem uma coluna a mais ("Hora"), então o Power BI pode precisar de ajuste.
6. Depois de tudo funcionando, faça commit e push na mesma branch.

## Regras

1. Nunca coloque senha no código nem peça senha pelo chat.
2. Não envie nada para asset@ antes da minha aprovação.
3. Respostas curtas e objetivas, sem travessões.

## Pendências que ficaram da outra sessão

1. Existe uma rotina na nuvem chamada **"Clipping Tyton"** (Claude Code na web), já desativada. Pode ser excluída na tela de Rotinas do claude.ai.
2. O ambiente de nuvem dessa rotina tem a variável `CLIPPING_EMAIL_SENHA` com a senha do tyton@. Ela não é mais usada e deve ser apagada.
3. Se ainda não foi feito: trocar a senha do tyton@ (ela foi colada num chat) e pedir ao Pedro para revogar a senha de app do Gmail que estava no código antigo.
