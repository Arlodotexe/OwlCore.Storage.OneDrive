using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System;

namespace OwlCore.Storage.OneDrive;

/// <summary>
/// A folder implementation that interacts with a folder in OneDrive.
/// </summary>
public class OneDriveFolder : IChildFolder, IGetItem, IGetItemRecursive, IGetRoot, ICreatedAtOffset, ILastAccessedAtOffset, ILastModifiedAtOffset, IModifiableFolder, ICreateRenamedCopyOf, IMoveRenamedFrom
{
    private readonly GraphServiceClient _graphClient;
    private readonly string? _driveId;
    private OneDriveCreatedAtProperty? _createdAt;
    private OneDriveCreatedAtOffsetProperty? _createdAtOffset;
    private OneDriveLastAccessedAtProperty? _lastAccessedAt;
    private OneDriveLastAccessedAtOffsetProperty? _lastAccessedAtOffset;
    private OneDriveLastModifiedAtProperty? _lastModifiedAt;
    private OneDriveLastModifiedAtOffsetProperty? _lastModifiedAtOffset;

    /// <summary>
    /// Creates a new instance of <see cref="OneDriveFolder"/>.
    /// </summary>
    public OneDriveFolder(GraphServiceClient graphClient, DriveItem driveItem)
    {
        _graphClient = graphClient;
        _driveId = driveItem.ParentReference?.DriveId;
        DriveItem = driveItem;
    }

    /// <summary>
    /// Creates a new instance of <see cref="OneDriveFolder"/> using a known drive for better performance.
    /// </summary>
    /// <param name="graphClient">The authenticated Graph client.</param>
    /// <param name="drive">The drive that contains the item.</param>
    /// <param name="driveItem">The item that backs this folder.</param>
    public OneDriveFolder(GraphServiceClient graphClient, Drive drive, DriveItem driveItem)
    {
        _graphClient = graphClient;
        _driveId = drive.Id;
        DriveItem = driveItem;
    }

    /// <inheritdoc />
    public string Id => DriveItem.Id!;

    /// <inheritdoc />
    public string Name => DriveItem.Name!;

    /// <summary>
    /// The graph item that was provided as the backing implementation for this file.
    /// </summary>
    public DriveItem DriveItem { get; }

    /// <inheritdoc />
    public ICreatedAtProperty CreatedAt => _createdAt ??= new OneDriveCreatedAtProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public ICreatedAtOffsetProperty CreatedAtOffset => _createdAtOffset ??= new OneDriveCreatedAtOffsetProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public ILastAccessedAtProperty LastAccessedAt => _lastAccessedAt ??= new OneDriveLastAccessedAtProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public ILastAccessedAtOffsetProperty LastAccessedAtOffset => _lastAccessedAtOffset ??= new OneDriveLastAccessedAtOffsetProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public ILastModifiedAtProperty LastModifiedAt => _lastModifiedAt ??= new OneDriveLastModifiedAtProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public ILastModifiedAtOffsetProperty LastModifiedAtOffset => _lastModifiedAtOffset ??= new OneDriveLastModifiedAtOffsetProperty(this, _graphClient, DriveItem.Id!);

    /// <inheritdoc />
    public virtual async IAsyncEnumerable<IStorableChild> GetItemsAsync(StorableType type = StorableType.All, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (type == StorableType.None)
            throw new ArgumentOutOfRangeException(nameof(type));

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
        var result = await _graphClient.Drives[driveId].Items[Id].Children.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);

        foreach (var item in result!.Value!)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (item.Folder is not null && type.HasFlag(StorableType.Folder))
                yield return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, item);

            if (item.File is not null && type.HasFlag(StorableType.File))
                yield return new OneDriveFile(_graphClient, new Drive { Id = driveId }, item);
        }
    }

    /// <inheritdoc />
    public Task<IStorableChild> GetItemRecursiveAsync(string id, CancellationToken cancellationToken = default) => GetItemAsync(id, cancellationToken);

    /// <inheritdoc />
    public async Task<IStorableChild> GetItemAsync(string id, CancellationToken cancellationToken = default)
    {
        try
        {
            var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
            var driveItem = await _graphClient.Drives[driveId].Items[id].GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);

            if (driveItem?.Folder is not null)
                return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, driveItem);

            if (driveItem?.File is not null)
                return new OneDriveFile(_graphClient, new Drive { Id = driveId }, driveItem);
        }
        catch
        {
            // ignored
        }

        throw new FileNotFoundException();
    }

    /// <inheritdoc />
    public virtual async Task<IFolder?> GetParentAsync(CancellationToken cancellationToken = default)
    {
        if (DriveItem.ParentReference is null)
            return null;

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
        var parentDriveItem = await _graphClient.Drives[driveId].Items[DriveItem.ParentReference.Id].GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);

        return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, parentDriveItem!);
    }

    /// <inheritdoc />
    public async Task<IFolder?> GetRootAsync(CancellationToken cancellationToken = default)
    {
        if (DriveItem.Root is null)
            return null;

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
        var rootDriveItem = await _graphClient.Drives[driveId].Root.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);

        return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, rootDriveItem!);
    }

    // ── IModifiableFolder ────────────────────────────────────────────────────

    /// <inheritdoc />
    public Task<IFolderWatcher> GetFolderWatcherAsync(CancellationToken cancellationToken = default)
        => throw new NotSupportedException("Folder watching is not supported for OneDrive.");

    /// <inheritdoc />
    public async Task<IChildFile> CreateFileAsync(string name, bool overwrite = false, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);

        // Check if a file with this name already exists in this folder.
        if (!overwrite)
        {
            try
            {
                var existing = await _graphClient.Drives[driveId].Items[Id].ItemWithPath(name).GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
                if (existing?.File is not null)
                    return new OneDriveFile(_graphClient, new Drive { Id = driveId }, existing);
            }
            catch
            {
                // Item does not exist — fall through to create.
            }
        }

        // Upload an empty file; this creates or replaces the item.
        var created = await _graphClient.Drives[driveId].Items[Id].ItemWithPath(name).Content.PutAsync(
            new MemoryStream(),
            cancellationToken: cancellationToken).ConfigureAwait(false);

        return new OneDriveFile(_graphClient, new Drive { Id = driveId }, created!);
    }

    /// <inheritdoc />
    public async Task<IChildFolder> CreateFolderAsync(string name, bool overwrite = false, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(name));

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);

        if (!overwrite)
        {
            try
            {
                var existing = await _graphClient.Drives[driveId].Items[Id].ItemWithPath(name).GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
                if (existing?.Folder is not null)
                    return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, existing);
            }
            catch
            {
                // Item does not exist — fall through to create.
            }
        }

        var newFolder = new DriveItem
        {
            Name = name,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object>
            {
                { "@microsoft.graph.conflictBehavior", overwrite ? "replace" : "rename" }
            }
        };

        var created = await _graphClient.Drives[driveId].Items[Id].Children.PostAsync(newFolder, cancellationToken: cancellationToken).ConfigureAwait(false);
        return new OneDriveFolder(_graphClient, new Drive { Id = driveId }, created!);
    }

    /// <inheritdoc />
    public async Task DeleteAsync(IStorableChild item, CancellationToken cancellationToken = default)
    {
        if (item is null)
            throw new ArgumentNullException(nameof(item));

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
        await _graphClient.Drives[driveId].Items[item.Id].DeleteAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    // ── ICreateRenamedCopyOf ─────────────────────────────────────────────────

    /// <inheritdoc cref="ICreateCopyOf.CreateCopyOfAsync(IFile, bool, CancellationToken, CreateCopyOfDelegate)"/>
    public async Task<IChildFile> CreateCopyOfAsync(IFile sourceFile, bool overwrite, CancellationToken cancellationToken, CreateCopyOfDelegate fallback)
    {
        if (sourceFile is null)
            throw new ArgumentNullException(nameof(sourceFile));

        if (sourceFile is not OneDriveFile oneDriveSource)
            return await fallback(this, sourceFile, overwrite, cancellationToken).ConfigureAwait(false);

        var result = await TryGraphCopyAsync(oneDriveSource, sourceFile.Name, overwrite, cancellationToken).ConfigureAwait(false);
        return result ?? await fallback(this, sourceFile, overwrite, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc cref="ICreateRenamedCopyOf.CreateCopyOfAsync(IFile, bool, string, CancellationToken, CreateRenamedCopyOfDelegate)"/>
    public async Task<IChildFile> CreateCopyOfAsync(IFile sourceFile, bool overwrite, string desiredName, CancellationToken cancellationToken, CreateRenamedCopyOfDelegate fallback)
    {
        if (sourceFile is null)
            throw new ArgumentNullException(nameof(sourceFile));

        if (sourceFile is not OneDriveFile oneDriveSource)
            return await fallback(this, sourceFile, overwrite, desiredName, cancellationToken).ConfigureAwait(false);

        var fileName = string.IsNullOrWhiteSpace(desiredName) ? sourceFile.Name : desiredName;
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(desiredName));

        var result = await TryGraphCopyAsync(oneDriveSource, fileName, overwrite, cancellationToken).ConfigureAwait(false);
        return result ?? await fallback(this, sourceFile, overwrite, desiredName, cancellationToken).ConfigureAwait(false);
    }

    // ── IMoveRenamedFrom ─────────────────────────────────────────────────────

    /// <inheritdoc cref="IMoveFrom.MoveFromAsync(IChildFile, IModifiableFolder, bool, CancellationToken, MoveFromDelegate)"/>
    public async Task<IChildFile> MoveFromAsync(IChildFile sourceFile, IModifiableFolder sourceFolder, bool overwrite, CancellationToken cancellationToken, MoveFromDelegate fallback)
    {
        if (sourceFile is null)
            throw new ArgumentNullException(nameof(sourceFile));

        if (sourceFolder is null)
            throw new ArgumentNullException(nameof(sourceFolder));

        if (sourceFile is not OneDriveFile oneDriveSource)
            return await fallback(this, sourceFile, sourceFolder, overwrite, cancellationToken).ConfigureAwait(false);

        var result = await TryGraphMoveAsync(oneDriveSource, sourceFile.Name, overwrite, cancellationToken).ConfigureAwait(false);
        return result ?? await fallback(this, sourceFile, sourceFolder, overwrite, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc cref="IMoveRenamedFrom.MoveFromAsync(IChildFile, IModifiableFolder, bool, string, CancellationToken, MoveRenamedFromDelegate)"/>
    public async Task<IChildFile> MoveFromAsync(IChildFile sourceFile, IModifiableFolder sourceFolder, bool overwrite, string desiredName, CancellationToken cancellationToken, MoveRenamedFromDelegate fallback)
    {
        if (sourceFile is null)
            throw new ArgumentNullException(nameof(sourceFile));

        if (sourceFolder is null)
            throw new ArgumentNullException(nameof(sourceFolder));

        if (sourceFile is not OneDriveFile oneDriveSource)
            return await fallback(this, sourceFile, sourceFolder, overwrite, desiredName, cancellationToken).ConfigureAwait(false);

        var fileName = string.IsNullOrWhiteSpace(desiredName) ? sourceFile.Name : desiredName;
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Value cannot be null or whitespace.", nameof(desiredName));

        var result = await TryGraphMoveAsync(oneDriveSource, fileName, overwrite, cancellationToken).ConfigureAwait(false);
        return result ?? await fallback(this, sourceFile, sourceFolder, overwrite, desiredName, cancellationToken).ConfigureAwait(false);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private async Task<OneDriveFile?> TryGraphCopyAsync(OneDriveFile source, string destName, bool overwrite, CancellationToken cancellationToken)
    {
        try
        {
            var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);

            if (!overwrite)
            {
                try
                {
                    var existing = await _graphClient.Drives[driveId].Items[Id].ItemWithPath(destName).GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
                    if (existing is not null)
                        return new OneDriveFile(_graphClient, new Drive { Id = driveId }, existing);
                }
                catch
                {
                    // Does not exist — proceed with copy.
                }
            }

            var copyBody = new Microsoft.Graph.Drives.Item.Items.Item.Copy.CopyPostRequestBody
            {
                Name = destName,
                ParentReference = new ItemReference
                {
                    DriveId = driveId,
                    Id = Id
                }
            };

            // Graph copy is async; poll for the copied item by path.
            await _graphClient.Drives[driveId].Items[source.Id].Copy.PostAsync(copyBody, cancellationToken: cancellationToken).ConfigureAwait(false);

            // Return a wrapper for the destination item immediately; caller can decide when to observe it.
            return new OneDriveFile(_graphClient, new Drive { Id = driveId }, new DriveItem
            {
                Id = source.Id,
                Name = destName,
                ParentReference = new ItemReference { DriveId = driveId, Id = Id }
            });
        }
        catch
        {
            return null;
        }
    }

    private async Task<OneDriveFile?> TryGraphMoveAsync(OneDriveFile source, string destName, bool overwrite, CancellationToken cancellationToken)
    {
        try
        {
            var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);

            if (!overwrite)
            {
                try
                {
                    var existing = await _graphClient.Drives[driveId].Items[Id].ItemWithPath(destName).GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
                    if (existing is not null)
                        return new OneDriveFile(_graphClient, new Drive { Id = driveId }, existing);
                }
                catch
                {
                    // Does not exist — proceed with move.
                }
            }

            var patchBody = new DriveItem
            {
                Name = destName,
                ParentReference = new ItemReference
                {
                    DriveId = driveId,
                    Id = Id
                }
            };

            var moved = await _graphClient.Drives[driveId].Items[source.Id].PatchAsync(patchBody, cancellationToken: cancellationToken).ConfigureAwait(false);
            return moved is null ? null : new OneDriveFile(_graphClient, new Drive { Id = driveId }, moved);
        }
        catch
        {
            return null;
        }
    }

    private async Task<string> GetDriveIdAsync(CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(_driveId))
            return _driveId!;

        var drive = await _graphClient.Me.Drive.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        return drive!.Id!;
    }
}
