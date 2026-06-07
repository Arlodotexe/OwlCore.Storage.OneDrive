using Microsoft.Graph;
using Microsoft.Graph.Models;
using OwlCore.Storage.OneDrive.Streams;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OwlCore.Storage.OneDrive;

/// <summary>
/// A file implementation that interacts with a file in OneDrive.
/// </summary>
public class OneDriveFile : IFile, IChildFile, ICreatedAtOffset, ILastAccessedAtOffset, ILastModifiedAtOffset
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
    /// Creates a new instance of <see cref="OneDriveFile"/>.
    /// </summary>
    public OneDriveFile(GraphServiceClient graphClient, DriveItem driveItem)
    {
        _graphClient = graphClient;
        _driveId = driveItem.ParentReference?.DriveId;
        DriveItem = driveItem;
    }

    /// <summary>
    /// Creates a new instance of <see cref="OneDriveFile"/> using a known drive for better performance.
    /// </summary>
    /// <param name="graphClient">The authenticated Graph client.</param>
    /// <param name="drive">The drive that contains the item.</param>
    /// <param name="driveItem">The item that backs this file.</param>
    public OneDriveFile(GraphServiceClient graphClient, Drive drive, DriveItem driveItem)
    {
        _graphClient = graphClient;
        _driveId = drive.Id;
        DriveItem = driveItem;
    }

    /// <summary>
    /// The graph item that was provided as the backing implementation for this file.
    /// </summary>
    public DriveItem DriveItem { get; }

    /// <inheritdoc />
    public string Id => DriveItem.Id!;

    /// <inheritdoc />
    public string Name => DriveItem.Name!;

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
    public virtual async Task<IFolder?> GetParentAsync(CancellationToken cancellationToken = default)
    {
        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);
        var parent = await _graphClient.Drives[driveId].Items[DriveItem.ParentReference!.Id].GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);

        return new OneDriveFolder(_graphClient, parent!);
    }

    /// <inheritdoc />
    public async Task<Stream> OpenStreamAsync(FileAccess accessMode = FileAccess.Read, CancellationToken cancellationToken = default)
    {
        if (accessMode != FileAccess.Read && accessMode != FileAccess.Write && accessMode != FileAccess.ReadWrite)
            throw new ArgumentOutOfRangeException(nameof(accessMode));

        var driveId = await GetDriveIdAsync(cancellationToken).ConfigureAwait(false);

        // All modes go through OneDriveWriteStream so callers always get a seekable,
        // random-access buffer regardless of whether they intend to write.
        return await OneDriveWriteStream.CreateAsync(_graphClient, driveId, Id, accessMode, cancellationToken).ConfigureAwait(false);
    }

    private async Task<string> GetDriveIdAsync(CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(_driveId))
            return _driveId!;

        var drive = await _graphClient.Me.Drive.GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        return drive!.Id!;
    }
}