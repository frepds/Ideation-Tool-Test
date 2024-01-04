using OfficeOpenXml;
using Ideation_Tool_Test;

// Pad naar je Excel-bestand
var filePath = @"/Applications/Ideation-Test.xlsx";

// Zorg ervoor dat EPPlus de licentievoorwaarden accepteert.
// ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Open het Excel-bestand
using var package = new ExcelPackage(new FileInfo(filePath));

// Kies het eerste werkblad
var worksheet = package.Workbook.Worksheets[0];

// Loop door de rijen en kolommen
for (var row = 1; row <= worksheet.Dimension.End.Row; row++)
{
    string body = string.Empty, mail = string.Empty, name = string.Empty, notificationTimer = string.Empty;
    var sendMessage = false;

    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
    {
        var value = worksheet.Cells[row, col].Text; // Lees de waarde van de huidige cel
        switch (col)
        {
            case 1:
                body = value; // Body
                break;
            case 2:
                mail = value; // Mail
                break;
            case 3:
                name = value; // Title
                break;
            case 4:
                // Excel leest True of False uit in 1/0 1 is true, 0 is False vandaar de equals
                sendMessage = value.Equals("1");
                break;
            case 5:
                //Notification voor database ROW.
                notificationTimer = value;
                break;
        }
    }

    string unsubscribeQuery = "SELECT CustomerName, City FROM Customers";

    Console.WriteLine($"Final >> Mail: {mail} | Message: {body} | Title: {name}");

    if (sendMessage)
    {
        switch (notificationTimer)
        {
            case "daily":
                MailService.SendEmail(mail, name, Message.CreateHtmlMessage(name, unsubscribeQuery));
                break;
            case "weekly":
                MailService.SendEmail(mail, name, Message.CreateHtmlMessage(name, unsubscribeQuery));
                break;
            case "monthly":
                MailService.SendEmail(mail, name, Message.CreateHtmlMessage(name, unsubscribeQuery));
                break;
        }
    }
}
 
// SendEmail("ontvanger@example.com", "Test Onderwerp", "Dit is een testbericht.");
