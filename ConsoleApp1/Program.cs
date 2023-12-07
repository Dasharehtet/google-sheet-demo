// See https://aka.ms/new-console-template for more information

using Microsoft.Extensions.Configuration;

namespace ConsoleApp1;

public abstract class Program
{
    private static void Main ()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.Development.json", optional: true);
        
        IConfiguration config = builder.Build();

        var serviceAccountEmail = config.GetSection("ServiceAccountEmail").Value;
        var privateKey = config.GetSection("PrivateKey").Value;
        var spreadsheetId = config.GetSection("SpreadsheetId").Value;
        var sheet = config.GetSection("First").Value;
        var folderId = config.GetSection("FolderId").Value;
        var applicationName = config.GetSection("ApplicationName").Value;
        var googleService = new GoogleService(
            serviceAccountEmail!,
            privateKey!,
            spreadsheetId!,
            sheet!, 
            applicationName!);

        var data = googleService.GetData();
        googleService.SetData();
        googleService.UpdateData();
        googleService.DeleteData();
        
        googleService.FindSheet();
        googleService.CreateSpreadsheet(folderId!);
    }
}