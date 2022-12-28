using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentLoadSanityCheckerDownload.Models;
using DocumentLoadSanityCheckerDownload.Models.Configs;
using DocumentLoadSanityCheckerDownload.Services;
using DocumentLoadSanityCheckerDownload.ViewModels;
using Microsoft.AspNetCore.Mvc;
using RestSharp;
using RestSharp.Serializers.NewtonsoftJson;
using System;

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
        const string sheetName = "Picklist";
        Dictionary<string, int> PicklistColumns;
        List<string> columnHeaders;
        List<string> columnHeadersGerman;
        List<string> YesNoList;
        #endregion
        public SanityCheckerDownloadController(SDxConfig config, StorageAccountConfig config1, AuthenticationService service) 
        {
            sDxConfig = config;
            accountConfig = config1;
            authService = service;
            blobService = new BlobStorageService(accountConfig.ConnectionString, accountConfig.Container);
            columnHeaders = new List<string>()
            {
                "revision code",
                "originator company",
                "document status code",
                "language",
                "discipline",
                "document type short code",
                "discipline document type short code",
                "export control classification",
                "storage media",
                "site critical document",
                "compliance record",
                "security code",
                "document owner",
                "area code",
                "area name",
                "Floc Level 2",
                "Floc Level 3",
                "Floc Level 23"

            };
            columnHeadersGerman = new List<string>()
            {
                "Revisionsnummer",
                "Hersteller",
                "Dokumentenstatus",
                "Sprache",
                "Disziplin",
                "Dokumententyp",
                "Disziplin Dokumententyp",
                "Exportkontrolle",
                "Speichermedium",
                "Kritisches Dokument",
                "Compliance Record",
                "Sicherheitseinstufung",
                "Dokumentenverantwortliche",
                "Ablagenummer",
                "area name - GLC",
                "Floc Level 2",
                "Floc Level 3",
                "Floc Level 23"
            };

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
                { "Users", 13 },
                { "areaCodes", 14 },
                { "flocData", 16 }
            };
            YesNoList = new List<string>()
            {
                "yes",
                "no"
            };
        }
        [HttpGet]
        public async Task<IActionResult> IndexAsync()
        {
            PlantCodeInputOutput output = new()
            {
                code = await GetPlantsAsync()
            };
            return View(output);
        }

        [HttpPost]
        public IActionResult GetDocSanityChecker(PlantCodeInputOutput plantDetail)
        {
            var task = GetDataFromSDxAsync(plantDetail.plant.UID);
            task.Wait();
            var data = task.Result;
            MemoryStream ms = UpdateDocSanityCheckerMacro(data, plantDetail.plant);

            //MemoryStream fileContent = blobService.DownloadFileFromBlob("ReportTemplate/LUB LATEST DOCUMENT LOAD SANITY CHECKER - " + plantDetail.Name + ".xlsm");
            if(ms != null)
                return File(ms.ToArray(), System.Net.Mime.MediaTypeNames.Application.Octet, string.Concat("LUB LATEST DOCUMENT LOAD SANITY CHECKER - ", plantDetail.plant.UID.AsSpan(3), ".xlsm"));
            return null;
        }


        #region Private methods
        private MemoryStream UpdateDocSanityCheckerMacro(SDxData data, PlantCodeData plantData)
        {
            MemoryStream stream = null;
            char[] reference = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            SheetData sheetData = new();
            Row row = new();
            uint rowIndex = 2;
            try
            {
                if(plantData.UID == "PL_GLC")
                {
                    foreach (var column in columnHeadersGerman)
                    {
                        Cell cell = new()
                        {
                            CellValue = new CellValue(column),
                            DataType = CellValues.String
                        };
                        if (column == "Floc Level 2" || column == "Floc Level 3" || column.Contains("area"))
                            cell.CellValue = new CellValue(string.Concat(column, " - ", plantData.UID.AsSpan(3)));
                        if (column == "Floc Level 23")
                            cell.CellValue = new CellValue(string.Concat(column, "-", plantData.UID.AsSpan(3), "-Merge"));
                        row.AppendChild(cell);
                    }
                }
                else
                {
                    foreach (var column in columnHeaders)
                    {
                        Cell cell = new()
                        {
                            CellValue = new CellValue(column),
                            DataType = CellValues.String
                        };
                        if (column == "Floc Level 2" || column == "Floc Level 3")
                            cell.CellValue = new CellValue(string.Concat(column, " - ", plantData.UID.AsSpan(3)));
                        if (column == "Floc Level 23")
                            cell.CellValue = new CellValue(string.Concat(column, "-", plantData.UID.AsSpan(3), "-Merge"));
                        row.AppendChild(cell);
                    }
                }
                
                sheetData.AppendChild(row);

                //writing the revision code picklist
                foreach (var code in data.revisionCodeData.MajorRevision)
                {
                    Row r = new();
                    Cell cell = new()
                    {
                        CellValue = new CellValue(code),
                        CellReference = reference[PicklistColumns["revisionCodeData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the originator company picklist
                rowIndex = 2;
                foreach (var companyName in data.originatorCompanyData)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(companyName.Name),
                        CellReference = reference[PicklistColumns["originatorCompanyData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;

                }

                //writing the document status code picklist
                rowIndex = 2;
                foreach (var docStatus in data.docStatusCodeData)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(docStatus.Name),
                        CellReference = reference[PicklistColumns["docStatusCodeData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the language picklist
                rowIndex = 2;
                foreach (var language in data.languageData)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(language.Name),
                        CellReference = reference[PicklistColumns["languageData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the discipline and doc type picklist
                rowIndex = 2;
                foreach (var discipline in data.disciplineData)
                {
                    foreach(var doctype in discipline.DisciplineDocumentClass)
                    {
                        Row r;
                        if (rowIndex > sheetData.ChildElements.Count)
                            r = new();
                        else
                            r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                        Cell cell = new()
                        {
                            CellValue = new CellValue(discipline.Name),
                            CellReference = reference[PicklistColumns["disciplineData"] - 1].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);
                    
                        cell = new()
                        {
                            CellValue = new CellValue(doctype.CFIHOSPrimKey),
                            CellReference = reference[PicklistColumns["disciplineData"]].ToString() + rowIndex,
                            DataType = CellValues.Number
                        };
                        r.AppendChild(cell);
                    
                        cell = new()
                        {
                            CellValue = new CellValue(discipline.Name + doctype.CFIHOSPrimKey),
                            CellReference = reference[PicklistColumns["disciplineData"] + 1].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);
                        if (rowIndex > sheetData.ChildElements.Count)
                            sheetData.AppendChild(r);
                        rowIndex++;
                    }
                }

                //writing the export control class picklist
                rowIndex = 2;
                foreach (var exportclass in data.exportControlClassData)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(exportclass.Name),
                        CellReference = reference[PicklistColumns["exportControlClassData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the storage media picklist
                rowIndex = 2;
                foreach (var media in data.mediaData)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1]; 
                    Cell cell = new()
                    {
                        CellValue = new CellValue(media.Name),
                        CellReference = reference[PicklistColumns["mediaData"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the site critical and complaince record picklist
                rowIndex = 2;
                foreach(var entry in YesNoList)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(entry),
                        CellReference = reference[PicklistColumns["mediaData"]].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                
                    cell = new()
                    {
                        CellValue = new CellValue(entry),
                        CellReference = reference[PicklistColumns["mediaData"] + 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    }; 
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the security code picklist
                rowIndex = 2;
                foreach (var code in data.securityCodes)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(code.Name),
                        CellReference = reference[PicklistColumns["securityCodes"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the document user picklist
                rowIndex = 2;
                foreach (var user in data.Users)
                {
                    Row r;
                    if (rowIndex > sheetData.ChildElements.Count)
                        r = new();
                    else
                        r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                    Cell cell = new()
                    {
                        CellValue = new CellValue(user.Email.ToLower()),
                        CellReference = reference[PicklistColumns["Users"] - 1].ToString() + rowIndex,
                        DataType = CellValues.String
                    };
                    r.AppendChild(cell);
                    if (rowIndex > sheetData.ChildElements.Count)
                        sheetData.AppendChild(r);
                    rowIndex++;
                }

                //writing the area code picklist
                rowIndex = 2;
                if(data.areaCodes.Count != 0)
                {
                    foreach(var code in data.areaCodes)
                    {
                        Row r;
                        if (rowIndex > sheetData.ChildElements.Count)
                            r = new();
                        else
                            r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                        Cell cell = new()
                        {
                            CellValue = new CellValue(code.Name),
                            CellReference = reference[PicklistColumns["areaCodes"] - 1].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);

                        cell = new()
                        {
                            CellValue = new CellValue(code.Description),
                            CellReference = reference[PicklistColumns["areaCodes"]].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);
                        if (rowIndex > sheetData.ChildElements.Count)
                            sheetData.AppendChild(r);
                        rowIndex++;
                    }
                }

                //writing the floc picklist
                rowIndex = 2;
                foreach (var floc in data.flocData)
                    {
                        Row r;
                        if (rowIndex > sheetData.ChildElements.Count)
                            r = new();
                        else
                            r = (Row)sheetData.ChildElements[Convert.ToInt32(rowIndex) - 1];
                        Cell cell = new()
                        {
                            CellValue = new CellValue(floc.FlocLevel2),
                            CellReference = reference[PicklistColumns["flocData"] - 1].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);

                        cell = new()
                        {
                            CellValue = new CellValue(floc.Name),
                            CellReference = reference[PicklistColumns["flocData"]].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);

                        cell = new()
                        {
                            CellValue = new CellValue(floc.FlocLevel2 + floc.Name),
                            CellReference = reference[PicklistColumns["flocData"] + 1].ToString() + rowIndex,
                            DataType = CellValues.String
                        };
                        r.AppendChild(cell);
                        if (rowIndex > sheetData.ChildElements.Count)
                            sheetData.AppendChild(r);
                        rowIndex++;                
                }
                if(plantData.UID == "PL_GLC")
                    stream = blobService.DownloadFileFromBlob("ReportTemplate/LUB DOCUMENT LOAD SANITY CHECKER GERMAN.xlsm");
                else
                    stream = blobService.DownloadFileFromBlob("ReportTemplate/LUB DOCUMENT LOAD SANITY CHECKER.xlsm");
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(stream, true))
                {
                    // Access the main Workbook part, which contains all references.
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;

                    // get sheet by name
                    var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                    // get worksheetpart by sheet id
                    WorksheetPart worksheetpart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;

                    // get worksheetpart by sheet id
                    worksheetpart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    //Remove existing sheet data and replace with new data
                    worksheetpart.Worksheet.RemoveAllChildren<SheetData>();
                    worksheetpart.Worksheet.AddChild(sheetData);

                    // Save the worksheet.
                    worksheetpart.Worksheet.Save();

                    // for recacluation of formula
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    spreadSheet.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

                }
                stream.Position = 0;
                blobService.UploadFileToBlob("ReportTemplate/LUB LATEST DOCUMENT LOAD SANITY CHECKER - " + plantData.Name + ".xlsm", stream);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return stream;
            
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
                string odataUrl = sDxConfig.ServerBaseUri + $"RevisionSchemes?$filter=Name eq '{RevisionScheme}'&$select=Name,Major_Seq,Minor_Seq&$count=true";
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
