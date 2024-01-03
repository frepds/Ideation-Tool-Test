using OfficeOpenXml;

// Pad naar je Excel-bestand
var filePath = @"/Applications/Ideation-Test.xlsx";

// Zorg ervoor dat EPPlus de licentievoorwaarden accepteert.
//ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Open het Excel-bestand
using var package = new ExcelPackage(new FileInfo(filePath));

// Kies het eerste werkblad
var worksheet = package.Workbook.Worksheets[0];

// Loop door de rijen en kolommen
for (var row = 1; row <= worksheet.Dimension.End.Row; row++) 
{ 
    for (var col = 1; col <= worksheet.Dimension.End.Column; col++)
    {
        // Lees de waarde van de huidige cel
        var value = worksheet.Cells[row, col].Text;

        // Print de waarde in de console
        Console.Write(value + "\t");
    }
    Console.WriteLine();
}
