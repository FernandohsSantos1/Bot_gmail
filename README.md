# Bot para Automação: gmail

Bot criado para realizar a leitura de uma tabela .xlsx, através da lib pandas, e
enviar automaticamente gmail para o destinatário desejado, com informações desta 
tabela, através da lib pyautogui e pyperclip.

O programa recebe como parametros de entrada:

   - Nome do destinatário que será enviado o email;
   - Email do destinatário;
   - Link do google drive onde se encontra o planilha;
   - Nome da planilha
   
* A automatização do mouse através do pyautogui é calculada a partir de resolução
da tela, então é necessário ajustes para monitores com resoluções diferentes.
* É possivel descobrir a posição do mouse através do comando pyautogui.position()
  
