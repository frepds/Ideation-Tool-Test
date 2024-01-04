namespace Ideation_Tool_Test;

public class Message
{
    public static string CreateHtmlMessage(string name, string unsubscribe)
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
                      Unsubscribe by following this link: <p onclick='{unsubscribe}'>Unsubscribe</p>
                </div>
            </div>
        </body>
        </html>";
    }
}