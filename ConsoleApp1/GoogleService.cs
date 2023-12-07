using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using File = Google.Apis.Drive.v3.Data.File;

namespace ConsoleApp1
{
    public class GoogleService
    {
        private const string SpreadsheetMime = "application/vnd.google-apps.spreadsheet";
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets, DriveService.Scope.Drive };
        private readonly string _spreadsheetId;
        private readonly string _sheet;
        private readonly SheetsService _sheetsService;
        private readonly DriveService _driveService;

        public GoogleService(string serviceAccountEmail, string privateKey, string spreadsheet, string sheet, string applicationName)
        {
            _spreadsheetId = spreadsheet;
            _sheet = sheet;

            var credential = new ServiceAccountCredential(
                new ServiceAccountCredential.Initializer(serviceAccountEmail)
                {
                    Scopes = Scopes,
                }.FromPrivateKey(privateKey));

            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
            
            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }
        
        public IEnumerable<IList<object>> GetData()
        {
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, $"{_sheet}!A1:F1");
            var response = request.Execute();
            var rows = response.Values;
            
            return rows;
        }

        public void SetData()
        {
            var valueRange = new ValueRange();
            
            var objectList = new List<object>() {"Hello!", "This", "was", "inserted", "by", "C#"};
            valueRange.Values = new List<IList<object>>() {objectList};

            var appendRequest = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, $"{_sheet}!A:F");
            appendRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            appendRequest.Execute();
        }
        
        public void UpdateData()
        {
            var valueRange = new ValueRange();
            
            var objectList = new List<object>() {"updated"};
            valueRange.Values = new List<IList<object>>() {objectList};

            var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, $"{_sheet}!A2");
            updateRequest.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            updateRequest.Execute();
        }

        public void DeleteData()
        {
            var requestBody = new ClearValuesRequest();

            var deleteRequest = _sheetsService.Spreadsheets.Values.Clear(requestBody, _spreadsheetId, $"{_sheet}!B2:F");
            deleteRequest.Execute();
        }

        public void AddSheet()
        {
            var addSheetRequest = new AddSheetRequest();
            var sheetProperties = new SheetProperties
            {
                Title = "Test"
            };


            addSheetRequest.Properties = sheetProperties;
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };
            batchUpdateSpreadsheetRequest.Requests.Add(new Request
            {
                AddSheet = addSheetRequest
            });

            var batchUpdateRequest =
                _sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, _spreadsheetId);

            batchUpdateRequest.Execute();
        }
        
        public void FindSheet()
        {
            var sheetMetadataRequest = _sheetsService.Spreadsheets.Get(_spreadsheetId);

            var response = sheetMetadataRequest.Execute();
            foreach (var sheet in response.Sheets)
            {
                Console.WriteLine($"{sheet.Properties.Title}{Environment.NewLine}");
            }
        }

        public void CreateSpreadsheet(string folderId)
        {
            var fileMetadata = new File
            {
                Name = "testSpreadsheet",
                Kind = "drive.file",
                MimeType = SpreadsheetMime,
                Parents = new List<string>
                {
                    folderId
                }
            };

            var createFileRequest = _driveService.Files.Create(fileMetadata);
            createFileRequest.Execute();
        }
    }
}