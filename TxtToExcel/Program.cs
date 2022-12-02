using System.CommandLine;
using System.CommandLine.Parsing;
using ClosedXML.Excel;

namespace TxtToExcel;

static class Program
{
    static async Task<int> Main(string[] args)
    {
        var sourceFileOption = new Option<FileInfo>(
            name: "--txt",
            description: "The text file to read and import into Excel file")
        {
            Arity = ArgumentArity.ExactlyOne,
            IsRequired = true,
        };
        sourceFileOption.AddValidator(FileExistsValidator(sourceFileOption));

        var targetFileOption = new Option<FileInfo>(
            name: "--xls",
            description: "The target Excel file")
        {
            Arity = ArgumentArity.ExactlyOne,
            IsRequired = true,
        };

        var sheetOption = new Option<string>(
            name: "--sheet",
            description: "Sheet name (it will be created if does not exist)"
        )
        {
            Arity = ArgumentArity.ExactlyOne,
        };

        var cellOption = new Option<string>(
            name: "--cell",
            description: "Starting cell to write imported data",
            getDefaultValue: () => "A1"
        )
        {
            Arity = ArgumentArity.ExactlyOne,
        };

        var rootCommand = new RootCommand("Text file to Excel importer");
        rootCommand.AddOption(sourceFileOption);
        rootCommand.AddOption(targetFileOption);
        rootCommand.AddOption(sheetOption);
        rootCommand.AddOption(cellOption);
        rootCommand.SetHandler(
            (source, target, sheet, cell) => { Import(source, target, sheet, cell); },
            sourceFileOption, targetFileOption, sheetOption, cellOption);

        return await rootCommand.InvokeAsync(args);
    }

    private static ValidateSymbolResult<OptionResult> FileExistsValidator(Option<FileInfo> fileOption)
    {
        return result =>
        {
            var file = result.GetValueForOption(fileOption);
            if (file?.Exists != true)
            {
                result.ErrorMessage = $"{file?.FullName} does not exist";
            }
        };
    }

    private static void Import(FileInfo source, FileInfo target, string sheetName, string cellAddress)
    {
        using var workbook = target.Exists
            ? new XLWorkbook(target.FullName)
            : new XLWorkbook();
        var worksheet = GetWorksheet(sheetName, workbook);

        var cell = worksheet.Cell(cellAddress);
        foreach (var line in File.ReadLines(source.FullName))
        {
            cell.Value = line;
            cell = cell.CellBelow();
        }

        if (target.Exists)
        {
            workbook.Save();
        }
        else
        {
            workbook.SaveAs(target.FullName);
        }
    }

    private static IXLWorksheet GetWorksheet(string sheetName, XLWorkbook workbook)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            return workbook.Worksheets.Count == 0
                ? workbook.Worksheets.Add()
                : workbook.Worksheets.First();

        return workbook.Worksheets.Contains(sheetName)
            ? workbook.Worksheets.Worksheet(sheetName)
            : workbook.Worksheets.Add(sheetName);
    }
}