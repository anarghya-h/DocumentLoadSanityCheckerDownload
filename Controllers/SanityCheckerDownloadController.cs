using DocumentLoadSanityCheckerDownload.Models;
using DocumentLoadSanityCheckerDownload.Models.Configs;
using DocumentLoadSanityCheckerDownload.Services;
using Microsoft.AspNetCore.Mvc;
using RestSharp;
using RestSharp.Serializers.NewtonsoftJson;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography.Xml;
using DocumentFormat.OpenXml.EMMA;

namespace DocumentLoadSanityCheckerDownload.Controllers
{
    public class SanityCheckerDownloadController : Controller
    {
        #region Private members
        SDxConfig sDxConfig;
        StorageAccountConfig accountConfig;
        AuthenticationService authService;
        BlobStorageService blobService;
        const string RevisionScheme = "RevLUBGlobal";
        Dictionary<string, int> PicklistColumns;
        #endregion
        public SanityCheckerDownloadController(SDxConfig config, StorageAccountConfig config1, AuthenticationService service) 
        {
            sDxConfig = config;
            accountConfig = config1;
            authService = service;
            blobService = new BlobStorageService(accountConfig.ConnectionString, accountConfig.Container);
            PicklistColumns = new Dictionary<string, int>()
            {
                { "revisionCodeData", 1 },
                { "originatorCompanyData", 2 },
                { "docStatusCodeData", 3 },
                { "languageData", 4 },
                { "disciplineData", 5 },
                { "exportControlClassData", 8 },
                { "mediaData", 9 },
                { "securityCodes", 12 },
                { "Users", 13 }
            };
        }
        [HttpGet]
        public async Task<IActionResult> IndexAsync()
        {
            PlantCodeData plantDetail = new()
            {
                UID = "PL_HLP",
                Name = "HLP"
            };
            var data = await GetDataFromSDxAsync(plantDetail.UID);
            UpdateDocSanityCheckerMacro(data, plantDetail);
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> GetDocSanityCheckerAsync(PlantCodeData plantDetail)
        {
            var data = await GetDataFromSDxAsync(plantDetail.UID);
            UpdateDocSanityCheckerMacro(data, plantDetail);

            MemoryStream fileContent = blobService.DownloadFileFromBlob("ReportTemplate/LUB LATEST DOCUMENT LOAD SANITY CHECKER - " + plantDetail.Name + ".xlsm");
            return File(fileContent.ToArray(), System.Net.Mime.MediaTypeNames.Application.Octet, "DOCUMENT LOAD SANITY CHECK.xlsm");
        }


        #region Private methods
        private void UpdateDocSanityCheckerMacro(SDxData data, PlantCodeData plantData)
        {
            char[] reference = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            try
            {
                MemoryStream stream = blobService.DownloadFileFromBlob("ReportTemplate/LUB DOCUMENT LOAD SANITY CHECKER.xlsm");
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(stream, true))
                {
                    // Access the main Workbook part, which contains all references.
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;

                    // get sheet by name
                    var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Picklist");

                    // get worksheetpart by sheet id
                    WorksheetPart worksheetpart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    uint rowIndex = 2;
                    //writing the revision code picklist
                    foreach(var code in data.revisionCodeData.MajorRevision)
                    {
                        Cell cell = GetCell(worksheetpart.Worksheet, reference[PicklistColumns["revisionCodeData"] - 1].ToString(), rowIndex);
                        cell.CellValue = new CellValue(code);
                        cell.DataType = CellValues.String;
                        rowIndex++;
                    }

                    //writing the originator company picklist
                    rowIndex = 2;
                    foreach(var companyName in data.originatorCompanyData)
                    {
                        Cell cell = GetCell(worksheetpart.Worksheet, reference[PicklistColumns["originatorCompanyData"] - 1].ToString(), rowIndex);
                        cell.CellValue = new CellValue(companyName.Name);
                        cell.DataType = CellValues.String;
                        rowIndex++;
                    }

                    // Save the worksheet.
                    worksheetpart.Worksheet.Save();

                    // for recacluation of formula
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                    //Uploading the file into the Storage account
         

                }
                stream.Position = 0;
                blobService.UploadFileToBlob("ReportTemplate/LUB LATEST DOCUMENT LOAD SANITY CHECKER - " + plantData.Name + ".xlsm", stream);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Cell cell = null;
            //Get the row
            Row row = worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row != null)
            {
                //Get the cell at the particular column
                cell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0);
                if (cell != null) return cell;
            }
            return cell;

        }

        private async Task<List<PlantCodeData>> GetPlantsAsync()
        {
            RestClient client = new RestClient().UseNewtonsoftJson();
            List<PlantCodeData> plantCodes = new List<PlantCodeData>();
            string OdataQueryUrl = sDxConfig.ServerBaseUri + $"Plants?$select=UID,Name&$count=true";
            var request = new RestRequest(OdataQueryUrl);
            request.AddHeader("Authorization", "Bearer " + authService.tokenResponse.AccessToken);
            var response = await client.GetAsync<OdataQueryResponse<PlantCodeData>>(request);
            if(response != null) 
                plantCodes.AddRange(response.Value);
            return plantCodes;

        }

        private async Task<List<LoginUserData>> GetLoginUsersAsync(string PlantUID)
        {            
            List<LoginUserData> Users = new List<LoginUserData>();
            
            var request = new RestRequest();
            request.AddHeader("Authorization", "Bearer " + authService.tokenResponse.AccessToken);
            request.AddHeader("SPFConfigUID", PlantUID);

            string OdataQueryUrl = sDxConfig.ServerBaseUri + $"SHLLoginUsers?$count=true&$select=Name,Address_email&$filter=Organization eq 'Shell'";
            RestClient client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var countResponse = await client.GetAsync<OdataQueryResponse<LoginUserData>>(request);

            OdataQueryUrl = OdataQueryUrl + "&$top=" + countResponse.Count.ToString();
            client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var Response = await client.GetAsync<OdataQueryResponse<LoginUserData>>(request);
            Users.AddRange(Response.Value);

            while (Response.NextLink != null)
            {
                client = new RestClient(Response.NextLink).UseNewtonsoftJson();
                Response = await client.GetAsync<OdataQueryResponse<LoginUserData>>(request);
                Users.AddRange(Response.Value);
            }
            return Users;

        }

        private async Task<List<AreaCodeData>> GetAreaCodesAsync(string PlantUID)
        {
            List<AreaCodeData> areaCodes = new List<AreaCodeData>();

            var request = new RestRequest();
            request.AddHeader("Authorization", "Bearer " + authService.tokenResponse.AccessToken);
            request.AddHeader("SPFConfigUID", PlantUID);

            string OdataQueryUrl = sDxConfig.ServerBaseUri + $"Locations?$count=true&$select=Name,Description";
            RestClient client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var countResponse = await client.GetAsync<OdataQueryResponse<AreaCodeData>>(request);

            OdataQueryUrl = OdataQueryUrl + "&$top=" + countResponse.Count.ToString();
            client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var Response = await client.GetAsync<OdataQueryResponse<AreaCodeData>>(request);
            areaCodes.AddRange(Response.Value);

            while (Response.NextLink != null)
            {
                client = new RestClient(Response.NextLink).UseNewtonsoftJson();
                Response = await client.GetAsync<OdataQueryResponse<AreaCodeData>>(request);
                areaCodes.AddRange(Response.Value);
            }
            return areaCodes;

        }

        private async Task<List<FlocData>> GetFlocDataAsync(string PlantUID)
        {
            List<FlocData> floc3Data = new List<FlocData>();

            var request = new RestRequest();
            request.AddHeader("Authorization", "Bearer " + authService.tokenResponse.AccessToken);
            request.AddHeader("SPFConfigUID", PlantUID);

            string OdataQueryUrl = sDxConfig.ServerBaseUri + $"Units?$select=Name,Floc_Level_2&$count=true";
            RestClient client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var countResponse = await client.GetAsync<OdataQueryResponse<FlocData>>(request);

            OdataQueryUrl = OdataQueryUrl + "&$top=" + countResponse.Count.ToString();
            client = new RestClient(OdataQueryUrl).UseNewtonsoftJson();
            var Response = await client.GetAsync<OdataQueryResponse<FlocData>>(request);
            floc3Data.AddRange(Response.Value);

            while (Response.NextLink != null)
            {
                client = new RestClient(Response.NextLink).UseNewtonsoftJson();
                Response = await client.GetAsync<OdataQueryResponse<FlocData>>(request);
                floc3Data.AddRange(Response.Value);
            }
            return floc3Data;

        }

        private async Task<SDxData> GetDataFromSDxAsync(string PlantUID)
        {
            SDxData data = null;
            try
            {
                var request = new RestRequest();
                request.AddHeader("Authorization", "Bearer " + authService.tokenResponse.AccessToken);
                request.AddHeader("SPFConfigUID", PlantUID);

                //fetching the details for Revision code
                string odataUrl = sDxConfig.ServerBaseUri + $"RevisionSchemes?$filter=Name eq 'RevLUBGlobal'&$select=Name,Major_Seq,Minor_Seq&$count=true";
                RestClient client = new RestClient(odataUrl).UseNewtonsoftJson();
                var RevisionCodeDataTask = client.GetAsync<OdataQueryResponse<RevisionCodeData>>(request);
                //response.Value[0].MajorRevision = response.Value[0].Major_Seq.Split(',').ToList();

                //Fetching the details for Originating company
                odataUrl = sDxConfig.ServerBaseUri + $"Organizations?$select=Name&$count=true";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var OrigCompanyDataTask = client.GetAsync<OdataQueryResponse<NameData>>(request);

                // Fetching the details for document status code
                odataUrl = sDxConfig.ServerBaseUri + $"SelectLists('e1CFIHOS_documentstatuscode')?$expand=Items($filter=IsDisabled eq false;$select=Name)";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var DocStatusCodeDataTask = client.GetAsync<DocStatusCodeData>(request);

                // Fetching the details for launguage
                odataUrl = sDxConfig.ServerBaseUri + $"SelectLists('e1SDADocLanguageCode')?$expand=Items($select=Name;$filter=IsDisabled eq false)";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var LanguageDataTask = client.GetAsync<LanguageData>(request);

                // Fetching the details for Discipline
                odataUrl = sDxConfig.ServerBaseUri + $"Objects?$filter=Class eq 'SDADiscipline'&$select=Name&$expand=SDADisciplineDocumentClass_12($select=CFIHOSPrimKey)&$count=true";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var DisciplineDocTypeDataTask = client.GetAsync<OdataQueryResponse<DocDisciplineData>>(request);

                // Fetching the details for export control classification
                odataUrl = sDxConfig.ServerBaseUri + $"SelectLists('e1SDA_exportcontrolclassification')?$expand=Items($select=Name;$filter=IsDisabled eq false)";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var ExportControlClassDataTask = client.GetAsync<ExportControlClassData>(request);

                // Fetching the details for storage media
                odataUrl = sDxConfig.ServerBaseUri + $"SelectLists('e1SDA_mediatype')?$expand=Items($select=Name;$count=true)";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var StorageMediaDataTask = client.GetAsync<StorageMediaData>(request);

                // Fetching the details for security codes
                odataUrl = sDxConfig.ServerBaseUri + $"SecurityCodes?$select=Name&$count=true";
                client = new RestClient(odataUrl).UseNewtonsoftJson();
                var SecurityCodeDataTask = client.GetAsync<OdataQueryResponse<NameData>>(request);

                //Fetching the details of users
                var UserDataTask = GetLoginUsersAsync(PlantUID);

                //Fetching the details of Area codes
                var AreaCodesDataTask = GetAreaCodesAsync(PlantUID);

                //Fetching the details of Floc
                var FlocDataTask = GetFlocDataAsync(PlantUID);

                await Task.WhenAll(RevisionCodeDataTask, OrigCompanyDataTask, DocStatusCodeDataTask, LanguageDataTask, DisciplineDocTypeDataTask, ExportControlClassDataTask, StorageMediaDataTask, SecurityCodeDataTask, UserDataTask, AreaCodesDataTask, FlocDataTask);

                data = new SDxData()
                {
                    revisionCodeData = RevisionCodeDataTask.Result.Value[0],
                    originatorCompanyData = OrigCompanyDataTask.Result.Value,
                    docStatusCodeData = DocStatusCodeDataTask.Result.Items,
                    languageData = LanguageDataTask.Result.Items,
                    disciplineData = DisciplineDocTypeDataTask.Result.Value,
                    exportControlClassData = ExportControlClassDataTask.Result.Items,
                    mediaData = StorageMediaDataTask.Result.Items,
                    securityCodes = SecurityCodeDataTask.Result.Value,
                    Users = UserDataTask.Result,
                    areaCodes = AreaCodesDataTask.Result,
                    flocData = FlocDataTask.Result
                };
                data.revisionCodeData.MajorRevision = data.revisionCodeData.Major_Seq.Split(',').ToList();
                data.revisionCodeData.MinorRevision = data.revisionCodeData.Minor_Seq.Split(',').ToList();
                return data;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        #endregion
    }
}
