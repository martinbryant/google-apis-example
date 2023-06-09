// For more information see https://aka.ms/fsharp-console-apps

open System.IO
open System.Threading
open Google.Apis.Auth.OAuth2
open Google.Apis.Services
open Google.Apis.Slides.v1
open Google.Apis.Slides.v1.Data
open Google.Apis.Util.Store
open Helpers

let secret = GoogleClientSecrets.FromFile("client-secret.json").Secrets

let path = Path.Combine("credentials")

let scopes = [| "https://www.googleapis.com/auth/presentations" |]

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

let slidesService = new SlidesService(clientService)

let subStringMatch = SubstringMatchCriteria(
    Text = "{{example-name}}",
    MatchCase = true)

let replaceText = ReplaceAllTextRequest(
    ContainsText = subStringMatch,
    ReplaceText = "Compositional IT")

let replaceTextRequest = Request(
    ReplaceAllText = replaceText)


let replaceImageText = SubstringMatchCriteria(
    Text = "{{example-logo}}",
    MatchCase = true)

let replaceShape = ReplaceAllShapesWithImageRequest(
    ImageUrl = "https://d1q6f0aelx0por.cloudfront.net/product-logos/library-fsharp-logo.png",
    ImageReplaceMethod = "CENTER_INSIDE",
    ContainsText = replaceImageText)
let replaceImageRequest = Request(
    ReplaceAllShapesWithImage = replaceShape)

let updates = BatchUpdatePresentationRequest(
    Requests = ResizeArray<Request>())

updates.Requests.Add(replaceTextRequest)
updates.Requests.Add(replaceImageRequest)

let updateResponse =
    task {
        return! slidesService.Presentations.BatchUpdate(updates, "1rcTWMDktfWUJRZVX7RdOqNHnLss7alvvCJNYJSqheHA").ExecuteAsync()
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously

let textReplacements =
    updateResponse.Replies
    |> Seq.filter (fun response -> response.ReplaceAllText <> null)
    |> Seq.sumBy (fun response -> response.ReplaceAllText.OccurrencesChanged |> Int.fromNullable)

let imageReplacements =
    updateResponse.Replies
    |> Seq.filter (fun response -> response.ReplaceAllShapesWithImage <> null)
    |> Seq.sumBy (fun response -> response.ReplaceAllText.OccurrencesChanged |> Int.fromNullable)

printfn $"Processed {textReplacements} text replacements and {imageReplacements} image replacements"

//Processed 2 text replacements and 1 images replacements

