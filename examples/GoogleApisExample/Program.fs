// For more information see https://aka.ms/fsharp-console-apps

open System.IO
open System.Threading
open Google.Apis.Auth.OAuth2
open Google.Apis.Drive.v3
open Google.Apis.Drive.v3.Data
open Google.Apis.Services
open Google.Apis.Util.Store

let secret = GoogleClientSecrets.FromFile("client-secret.json").Secrets

let path = Path.Combine("credentials")

let scopes = [| "https://www.googleapis.com/auth/drive" |]

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

let driveService = new DriveService(clientService)

let existingFileId = "1i3nAr8MYj50H_S-yjh5zAQQ1u_QqeL3AQH51S5Lt2d8"

let newFile = File(Name = "copy-file")

let file =
    task {
        return! driveService.Files.Copy(newFile, existingFileId).ExecuteAsync()
    }
    |> Async.AwaitTask
    |> Async.RunSynchronously
    
printf $"Created new file with Id {file.Id}"