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
        await LibreOffice.ConvertToPdf(xlsx); // output ./sample.pdf

        // or
        
        await LibreOffice.ConvertToPdf(xlsx, outputDirectory: "./output"); // output ./output/sample.pdf
    }
}
```
