# Automação de um Aquivo Excel

Esse projeto se trata de um sistema que entra no boletim pelo Suap (sistema do IFSP) via Selenium.
Pega as notas de todos os bimestres de todas as matérias.
Coloca essas notas em uma planilha Excel, desenvolvida pelo <a href="https://github.com/Caicao001">Caique Caires</a>.
Essa planilha faz a conta para ver quanto falta para ser aprovado na matéria ou se já foi aprovado. O mesmo acontece em relação as áreas de ensino.
Ao fim, o programa envia esse Excel pelo email que o usuário passou.

Versão com o Openpyxl.<br>
A diferença entre a <a href="https://github.com/jesbicaa/AutomacaoExcelNotas_Win32">Versão com Win32</a> é a biblioteca usada para o envio de email.<br>
Essa versão é necessário usar no campo senha o App passwords para conseguir enviar o email.
