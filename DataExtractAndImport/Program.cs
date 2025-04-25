using CommandLine;

namespace DataExtractAndImport;

[Verb("export", aliases: ["e"], HelpText = "Exports data from an mCase instance to an Excel file.")]
internal class ExportOptions
{
    [Option(
        'c',
        "connectionString",
        HelpText = "A connection string to connect to the mCase database, can rely on interactive authentication or MFA.",
        Default = "data source=localhost; initial catalog=mCASE_ADMIN; integrated security=True; min pool size=10; max pool size=500; multipleactiveresultsets=True; application name=mCaseAdminWebEF6; Connect Timeout=1200; Command Timeout=1200;  TrustServerCertificate=True"
    )]
    public string? ConnectionString { get; set; }

    [Option(
        'f',
        "filename",
        HelpText = "Name of the export file.",
        Default = "Security Matrix.xlsx"
    )]
    public string? Filename { get; set; }
}

[Verb("import", aliases: ["i"], HelpText = "Import data from an Excel file to an mCase instance.")]
internal class ImportOptions
{
    [Option(
        'f',
        "filename",
        Required = false,
        HelpText = "Name of the import file.",
        Default = "SecurityMatrix.xlsx"
    )]
    public string? Filename { get; set; }
}

internal class Program
{
    public static void Main(string[] args)
    {
        Parser
            .Default.ParseArguments<ExportOptions, ImportOptions>(args)
            .WithParsed<ExportOptions>((opt) => Export.Run(opt.ConnectionString!, opt.Filename!))
            .WithParsed<ImportOptions>(opt => throw new NotImplementedException())
            .WithNotParsed(Console.WriteLine);
    }
}
