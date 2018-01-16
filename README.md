# ProWAGoN
ProWAGoN - um "populador" de lista de bloqueios de anúncios para o Norton Internet Security e Norton Personal Firewall

Alguns anos atrás, existia um excelente "firewall pessoal" para computadores Windows, chamado "AtGuard". Uma das funções que ele possuía era uma "lista de bloqueio" de anúncios. Na época (1999), era algo sensacional, por questões de privacidade, de "limpeza" no visual das páginas, e é bom lembrar que a Internet nas casas funcionava, na maior parte, com modems 56K.

Então, certo dia, a Symantec comprou o AtGuard e o transformou no "Norton Internet Security" (NIS). Na versão 2002 do software, a Symantec começou a guardar esta lista usando criptografia, para proteção contra alterações maliciosas (vírus, malwares etc). Muito bom, muito bem, mas isso virou um problema para os usuários interessados em usar "listas de bloqueio" criados por terceiros -- como o (muito famoso) AGNIS, do professor Eric Howes da Universidade de Illinois (EUA), disponível em http://www.spywarewarrior.com.

_Eu_ era um dos que usavam estas listas terceirizadas, portanto, imaginei uma forma de fazer o Norton usá-las, sem precisar me preocupar com a criptografia da Symantec. Claro, teoricamente, qualquer pessoa poderia abrir a GUI do firewall e inserir as entradas da lista uma por uma, mas isso não seria nem um pouco prático -- somente a lista AGNIS, por exemplo, tem mais de 2000 entradas.

O resultado foi o ProWAGoN, um programinha que disponibilizei gratuitamente para "a comunidade", capaz de carregar o conteúdo de listas e inseri-los no NIS enviando as entradas através da interface gráfica. O professor Howes, criador do AGNIS, ficou extremamente interessado, e o divulgou para o mundo -- mensagens em sites de segurança, distribuição do meu software em seu site, e ainda criou um "readme" (completíssimo) para ele (http://www.spywarewarrior.com/uiuc/prowagon/pw-readme.htm). Jamais poderia agradecê-lo o suficiente.

Não tenho números absolutos, mas pelo tanto de emails que recebia, pedindo auxílio ou (muitas vezes) apenas agradecendo, o ProWAGon foi bastante popular, e recebeu atualizações no decorrer do tempo para operar até a versão 2007 do Norton Internet Security -- última versão do firewall que possuía a "lista de bloqueio". Imagino que a Symantec tenha descontinuado o recurso porque os navegadores web passaram a oferecê-lo, se não nativamente, através de plugins (AdBlock, uBlock, etc).

Ah, os bons tempos. Enfim, está aqui o código do ProWAGoN para quem quiser olhar. Extremamente simples, mas foi muito, muito útil. Dificilmente será usado com o NIS, mas talvez mostre como usar a API SendKeys -- o que é, basicamente, toda a "mágica" do programa.

E, para quem estiver se perguntando sobre o nome: significa simplesmente "Pro"gram "W"ithout "A" "Go"od "N"ame. Digamos que eu estava sem inspiração, e me lembrei do [TWAIN](https://pt.wikipedia.org/wiki/TWAIN), a tecnologia sem um nome interessante.
