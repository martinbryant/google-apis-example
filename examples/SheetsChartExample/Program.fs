// For more information see https://aka.ms/fsharp-console-apps

open System.IO
open System.Threading
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Sheets.v4
open Google.Apis.Slides.v1
open Google.Apis.Slides.v1.Data
open Google.Apis.Util.Store
open Helpers

let secret = GoogleClientSecrets.FromFile("client-secret.json").Secrets

let path = Path.Combine("credentials")

let scopes = [|
    "https://www.googleapis.com/auth/presentations"
    "https://www.googleapis.com/auth/spreadsheets"
|]

let credentials =
    task {
        return! GoogleWebAuthorizationBroker.AuthorizeAsync(
            secret,
            scopes,
            "martinbryant.dev@gmail.com",
            CancellationToken.None,
            FileDataStore(path, true))
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously

let clientService = BaseClientService.Initializer(
    HttpClientInitializer = credentials,
    ApplicationName = "test-slides-api")

let sheetService = new SheetsService(clientService)
let spreadsheetId = "1X-tCqmt4FDCG_JU0gD2PfOOwQYHCjv6o_XpedMOOSmk"
let spreadsheet =
    task {
        return! sheetService.Spreadsheets.Get(spreadsheetId).ExecuteAsync()
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously

let chartId = spreadsheet.Sheets[0].Charts[0].ChartId

let slideService = new SlidesService(clientService)

let updates = BatchUpdatePresentationRequest(
    Requests = ResizeArray<Request>())

let replaceText = SubstringMatchCriteria(
    Text = "{{example-chart}}",
    MatchCase = true)

let chartRequest = ReplaceAllShapesWithSheetsChartRequest(
    ContainsText = replaceText,
    SpreadsheetId = spreadsheetId,
    ChartId = chartId,
    LinkingMode = "NOT_LINKED_IMAGE")

let request = Request(
    ReplaceAllShapesWithSheetsChart = chartRequest)

updates.Requests.Add(request)

let updateResponse =
    task {
        return! slideService.Presentations.BatchUpdate(updates, "1rcTWMDktfWUJRZVX7RdOqNHnLss7alvvCJNYJSqheHA").ExecuteAsync()
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously

let chartReplacements =
    updateResponse.Replies
    |> Seq.filter (fun response -> response.ReplaceAllShapesWithSheetsChart <> null)
    |> Seq.sumBy (fun response -> response.ReplaceAllShapesWithSheetsChart.OccurrencesChanged |> Int.fromNullable)

printfn $"Processed {chartReplacements} chart replacement"
