# Outlook Domain Checker

## 🇮🇹 Italiano

Questo progetto contiene uno script VBA per **Outlook** che verifica se i domini degli indirizzi email esistono prima dell'invio.  
L'obiettivo è evitare errori generici in uscita quando si inviano email a domini inesistenti, ad esempio con provider come Aruba IMAP/SMTP.

### Installazione

1. Apri Outlook.  
2. Premi `ALT + F11` per aprire l'editor VBA.  
3. Vai su `ThisOutlookSession` nel progetto VBA.  
4. Copia il contenuto di `DomainChecker.bas` dentro `ThisOutlookSession`.  
5. Salva e chiudi l'editor.  
6. Riavvia Outlook per attivare la macro.

### Uso

Quando provi a inviare un'email, la macro verifica tutti i destinatari.  
Se il dominio di un destinatario non esiste, l'invio viene **annullato** e compare un messaggio di errore.

### Note di sicurezza

- È necessario abilitare le macro in Outlook.  
- Lo script legge solo i domini degli indirizzi email e non invia dati a terzi.

---

## 🇬🇧 English

This project contains a VBA script for **Outlook** that checks if email domains exist before sending.  
Its goal is to prevent generic sending errors when emails are addressed to non-existent domains, e.g., with providers like Aruba IMAP/SMTP.

### Installation

1. Open Outlook.  
2. Press `ALT + F11` to open the VBA editor.  
3. Go to `ThisOutlookSession` in the VBA project.  
4. Copy the content of `DomainChecker.bas` into `ThisOutlookSession`.  
5. Save and close the editor.  
6. Restart Outlook to activate the macro.

### Usage

When sending an email, the macro checks all recipients.  
If any recipient's domain does not exist, sending is **canceled** and an error message appears.

### Security Notes

- Macros must be enabled in Outlook.  
- The script only reads email domains and does not send any data externally.

