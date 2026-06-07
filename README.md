# OwlCore.Storage.OneDrive [![Version](https://img.shields.io/nuget/v/OwlCore.Storage.OneDrive.svg)](https://www.nuget.org/packages/OwlCore.Storage.OneDrive)

An implementation of OwlCore.Storage that uses Microsoft Graph to access OneDrive.

## Install

Published releases are available on [NuGet](https://www.nuget.org/packages/OwlCore.Storage.OneDrive). To install, run the following command in the [Package Manager Console](https://docs.nuget.org/docs/start-here/using-the-package-manager-console).

    PM> Install-Package OwlCore.Storage.OneDrive
    
Or using [dotnet](https://docs.microsoft.com/en-us/dotnet/core/tools/dotnet)

    > dotnet add package OwlCore.Storage.OneDrive

## Usage

Before you begin, obtain an instance of a `GraphServiceClient` from either: 
- The official [`Microsoft.Graph`](https://learn.microsoft.com/en-us/graph/sdks/create-client?tabs=CS) libraries. 
- The [`CommunityToolkit.Graph`](https://github.com/CommunityToolkit/Graph-Controls) helpers. Recommended for apps in the Microsoft Store.


```cs
var graphClient = CreateGraphClient(...);

// Then, get the desired drive to work in.
var drive = await graphClient.Me.Drive.GetAsync();

// To get folder from a known folder Id:
var knownFolderId = "someId";
var driveItem = await graphClient.Drives[drives.Id].Items[knownFolderId].GetAsync();

// Or, to get user's root folder in OneDrive:
var driveItem = await graphClient.Drives[drive.Id].Root.GetAsync();

// Pass starting item to a new OneDrive folder.
// If you already have the drive, prefer the drive-aware constructor for better performance.
var oneDrive = new OneDriveFolder(graphClient, drive!, driveItem);

// Interact with OneDrive via the storage abstraction.
// Extension methods GetFilesAsync and GetFolderAsync also available.
await foreach (var item in oneDrive.GetItemsAsync(StorableType.All))
{
    if (item is IFile file)
    {
        using var stream = await file.OpenStreamAsync();
        // ...
    }

    if (item is IFolder folder)
    {
        var root = await folder.GetRootAsync();
        // ... 
    }
}
```

## Financing

We accept donations [here](https://github.com/sponsors/Arlodotexe) and [here](https://www.patreon.com/arlodotexe), and we do not have any active bug bounties.

## Versioning

Version numbering follows the Semantic versioning approach. However, if the major version is `0`, the code is considered alpha and breaking changes may occur as a minor update.

## License

All OwlCore code is licensed under the MIT License. OwlCore is licensed under the MIT License. See the [LICENSE](./src/LICENSE.txt) file for more details.

