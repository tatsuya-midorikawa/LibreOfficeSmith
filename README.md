![logo](https://raw.githubusercontent.com/tatsuya-midorikawa/LibreOfficeSmith/main/assets/libreoffice_smith.png)
# LibreOfficeSmith
Helper library for LibreOffice operations.

# How to use

## Conver to PDF

```cs
using LibreOfficeSmith.Csharp;

class Program
{
    const string xlsx = @"./src/sample.xlsx";

    static async System.Threading.Tasks.Task Main(string[] args)
    {
        await LibreOffice.ConvertToPdf(xlsx); // output to "./sample.pdf"

        // or
        
        await LibreOffice.ConvertToPdf(xlsx, outputDirectory: "./output"); // output to "./output/sample.pdf"
    }
}
```
