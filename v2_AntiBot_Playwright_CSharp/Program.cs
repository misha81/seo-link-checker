using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Playwright;

namespace SeoLinkChecker
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "SEO Playwright Checker";

        static readonly string SpreadsheetId = "INSERT_YOUR_SPREADSHEET_ID_HERE";
        static readonly string SheetName = "Sheet1";

        static async Task Main(string[] args)
        {
            Console.WriteLine("Authenticating with Google Sheets...");
            GoogleCredential credential;
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            Console.WriteLine("Launching Playwright...");
            using var playwright = await Playwright.CreateAsync();

            string wsEndpoint = "ws://127.0.0.1:9222/devtools/browser/INSERT_YOUR_DEBUG_PORT_HERE";

            Console.WriteLine("Connecting to anti-detect browser...");
            await using var browser = await playwright.Chromium.ConnectOverCDPAsync(wsEndpoint);

            var context = browser.Contexts[0];

            var range = $"{SheetName}!A2:F";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = await request.ExecuteAsync();
            var values = response.Values;

            if (values != null && values.Count > 0)
            {
                Dictionary<string, string> checkedCache = new Dictionary<string, string>();

                for (int i = 0; i < values.Count; i++)
                {
                    var row = values[i];
                    var anchor = row.Count > 1 ? row[1].ToString().Trim().ToLower() : "";
                    var targetLink = row.Count > 2 ? row[2].ToString().Trim().ToLower() : "";
                    var readyLink = row.Count > 4 ? row[4].ToString().Trim() : "";

                    int rowIndex = i + 1;

                    if (string.IsNullOrEmpty(readyLink) || !readyLink.StartsWith("http")) continue;

                    string cacheKey = $"{readyLink}|{targetLink}";
                    if (checkedCache.ContainsKey(cacheKey))
                    {
                        string cachedResult = checkedCache[cacheKey];
                        Console.WriteLine($"\n[{i + 2}] [CACHE] Skipping duplicate: {readyLink}. Found status: {cachedResult}");

                        bool cachedIsSuccess = (cachedResult == "success");
                        string cachedError = cachedIsSuccess ? "" : cachedResult;

                        await UpdateRowInSheet(service, SpreadsheetId, SheetName, rowIndex, cachedError, cachedIsSuccess);
                        continue;
                    }

                    Console.WriteLine($"\n[{i + 2}] Checking: {readyLink}");
                    string errorStatus = "";
                    bool isSuccess = false;

                    var page = await context.NewPageAsync();

                    page.SetDefaultNavigationTimeout(20000);
                    page.SetDefaultTimeout(20000);

                    try
                    {
                        var gotoResponse = await page.GotoAsync(readyLink, new PageGotoOptions { WaitUntil = WaitUntilState.Load });

                        if (gotoResponse == null)
                        {
                            errorStatus = "Error / DNS";
                        }
                        else if (gotoResponse.Status == 404)
                        {
                            errorStatus = "Error 404";
                        }
                        else
                        {
                            await page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");
                            await Task.Delay(3000);

                            await page.EvaluateAsync(@"() => {
                                const elements = document.querySelectorAll('button, span, div');
                                for (let el of elements) {
                                    let text = (el.innerText || '').toLowerCase().trim();
                                    if (text === 'more' || text.includes('show more') || text.includes('load') || text.includes('більше') || text.includes('view more') || text.includes('ver más')) {
                                        if (!el.closest('a')) {
                                            try { el.click(); } catch(e) {}
                                        }
                                    }
                                }
                            }");
                            await Task.Delay(2000);

                            bool hasNoIndex = await page.EvaluateAsync<bool>(@"() => {
                                const metaTags = document.querySelectorAll('meta[name=""robots""]');
                                for (let meta of metaTags) {
                                    if (meta.content.toLowerCase().includes('noindex')) return true;
                                }
                                return false;
                            }");

                            if (hasNoIndex)
                            {
                                errorStatus = "no index";
                            }
                            else
                            {
                                var linkCheckResult = await page.EvaluateAsync<string>(@"
    ([targetLink, anchor]) => {
        try {
            if (!targetLink) return 'no link';
            const links = Array.from(document.querySelectorAll('a'));
            let linkFound = false;
            let foundTexts = [];

            for (let a of links) {
                if (!a.href) continue;

                if (a.href.toLowerCase().includes(targetLink.toLowerCase())) {
                    linkFound = true;
                    
                    if (!anchor || anchor.trim() === '') return 'success';
                    
                    let expectedAnchor = anchor.toLowerCase().trim();

                    let text = (a.innerText || '').trim();
                    if (text.length > 0) foundTexts.push(text);
                    if (text.toLowerCase().includes(expectedAnchor)) return 'success';

                    let img = a.querySelector('img');
                    if (img) {
                        let altText = (img.alt || '').trim();
                        if (altText.length > 0) foundTexts.push(altText);
                        
                        if (altText.toLowerCase().includes(expectedAnchor)) return 'success';
                        
                        return 'success'; 
                    }
                }
            }

            if (linkFound) {
                let debugText = foundTexts.join('');
                if (!debugText) return 'incorrect link_pink_icon';
                return 'incorrect link';
            }

            return 'no link';
        } catch (err) {
            return 'JS Error: ' + err.message;
        }
    }
", new[] { targetLink, anchor });

                                if (linkCheckResult == "success")
                                    isSuccess = true;
                                else
                                    errorStatus = linkCheckResult;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Loading error: {ex.Message}");
                        errorStatus = "Error / DNS";
                    }
                    finally
                    {
                        try
                        {
                            await page.CloseAsync();
                        }
                        catch
                        {
                            Console.WriteLine("Tab already closed or inaccessible.");
                        }
                    }

                    await UpdateRowInSheet(service, SpreadsheetId, SheetName, rowIndex, errorStatus, isSuccess);
                    Console.WriteLine($"Result: {(isSuccess ? "OK" : errorStatus)}");

                    checkedCache[cacheKey] = isSuccess ? "success" : errorStatus;
                }
            }
            else
            {
                Console.WriteLine("Data not found.");
            }
            Console.WriteLine("Check completed.");
        }

        static async Task UpdateRowInSheet(SheetsService service, string spreadsheetId, string sheetName, int rowIndex, string errorText, bool isSuccess)
        {
            Color bgColor;

            if (isSuccess)
            {
                bgColor = new Color { Red = 0.43f, Green = 0.82f, Blue = 0.56f };
            }
            else if (errorText == "incorrect link_pink_icon")
            {
                bgColor = new Color { Red = 0.96f, Green = 0.76f, Blue = 0.76f };
                errorText = "incorrect link";
            }
            else
            {
                bgColor = new Color { Red = 1f, Green = 0f, Blue = 0f };
            }

            var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { errorText } } };
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, $"{sheetName}!F{rowIndex + 1}");
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            await updateRequest.ExecuteAsync();

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = new List<Request>() };
            var cellFormat = new CellFormat { BackgroundColor = bgColor };

            var sheetInfoRequest = service.Spreadsheets.Get(spreadsheetId);
            var sheetInfo = await sheetInfoRequest.ExecuteAsync();
            int? sheetId = null;
            foreach (var s in sheetInfo.Sheets)
            {
                if (s.Properties.Title == sheetName) { sheetId = s.Properties.SheetId; break; }
            }

            batchUpdateRequest.Requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange { SheetId = sheetId, StartRowIndex = rowIndex, EndRowIndex = rowIndex + 1, StartColumnIndex = 4, EndColumnIndex = 5 },
                    Cell = new CellData { UserEnteredFormat = cellFormat },
                    Fields = "userEnteredFormat.backgroundColor"
                }
            });

            await service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).ExecuteAsync();
        }
    }
}