using OfficeOpenXml;
using System.Net;
using System.Net.Mail;

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
    string body = string.Empty, mail = string.Empty, name = string.Empty;
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
        }
    }
    
    Console.WriteLine($"Final >> Mail: {mail} | Message: {body} | Title: {name}");
    
    if (sendMessage)
    {
        SendEmail(mail, name, CreateHtmlMessage(name));
    }
}

// SendEmail("ontvanger@example.com", "Test Onderwerp", "Dit is een testbericht.");

static string CreateHtmlMessage(string name)
{
    return $@"
        <!DOCTYPE html>
        <html>
        <body>
            <div style='font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f2f2f2; padding: 20px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.05);'>
                <div style='background-color: #0095DA; color: white; padding: 10px; text-align: center; border-top-left-radius: 8px; border-top-right-radius: 8px;'>
                    <h1>Ideation-Tool Notification</h1>
                </div>
                <div style='padding: 20px; text-align: left; line-height: 1.5; color: #333;'>
                    <p>Dear {name},</p>
                    <p>Thank you for signing on by the Ideation-tool©.</p>
                    <p>Your input is greatly appreciated.</p>
                    <p>With kind regards,</p>
                    <p>Abbott Logistics</p>
                </div>
                <div style='text-align: center; padding-top: 20px; font-size: 12px; color: #888;'>
                    © 2024 Abbott Laboratories - All rights reserved<br>
                      Unsubscribe by following this link: <a href='https://www.google.com/'>Unsubscribe</a>
                </div>
            </div>
        </body>
        </html>";
}

static void SendEmail(string to, string subject, string body)
{
    try
    {
        var fromAddress = new MailAddress("mailserviceabbott@gmail.com", "Abbott HQ"); 
        var toAddress = new MailAddress(to); 
        const string fromPassword = "ppcc wynp fhbi nexc";

        var smtp = new SmtpClient 
        {
            Host = "smtp.gmail.com", // Vervang door de SMTP server van je provider
            Port = 587, // Standaard SMTP poort, kan variëren afhankelijk van je provider
            EnableSsl = true, // Afhankelijk van je provider
            DeliveryMethod = SmtpDeliveryMethod.Network, 
            UseDefaultCredentials = false, 
            Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
        };

        using var message = new MailMessage(fromAddress, toAddress)
        {
            Subject = subject, 
            Body = body,
            IsBodyHtml = true
        };
        
        smtp.Send(message); 
        Console.WriteLine("E-mail succesvol verzonden.");
    }
    catch (Exception ex)
    {
        Console.WriteLine("Fout bij het verzenden van e-mail: " + ex.Message);
    }
}
