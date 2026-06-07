using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

namespace OwlCore.Storage.OneDrive.Streams;

/// <summary>
/// A seekable, sparse in-memory stream backed by lazily loaded fixed-size blocks.
/// Only touched blocks are cached in memory.
/// </summary>
/// <remarks>
/// <list type="bullet">
///   <item><see cref="FileAccess.Read"/> preloads nothing and lazily fetches requested blocks from OneDrive.</item>
///   <item><see cref="FileAccess.Write"/> starts empty and allocates only written blocks.</item>
///   <item><see cref="FileAccess.ReadWrite"/> lazily fetches existing blocks and allocates written blocks.</item>
/// </list>
/// </remarks>
public sealed class OneDriveWriteStream : Stream
{
    /// <summary>
    /// The maximum size that is uploaded with a single PUT request.
    /// </summary>
    private const long LargeFileThreshold = 4 * 1024 * 1024;

    /// <summary>
    /// The slice size used when uploading via resumable upload sessions.
    /// </summary>
    private const int UploadSliceSize = 10 * 320 * 1024;

    /// <summary>
    /// The in-memory block size used for sparse caching.
    /// </summary>
    private const int BlockSize = 64 * 1024;

    private readonly GraphServiceClient _graphClient;
    private readonly string _driveId;
    private readonly string _itemId;
    private readonly FileAccess _accessMode;
    private readonly Dictionary<long, byte[]> _blocks = new();
    private readonly long _sourceLength;

    private long _length;
    private long _position;
    private bool _committed;
    private bool _disposed;

    private OneDriveWriteStream(GraphServiceClient graphClient, string driveId, string itemId, FileAccess accessMode, long sourceLength)
    {
        _graphClient = graphClient;
        _driveId = driveId;
        _itemId = itemId;
        _accessMode = accessMode;
        _sourceLength = sourceLength;
        _length = accessMode == FileAccess.Write ? 0 : sourceLength;
        _position = 0;
    }

    /// <summary>
    /// Creates a new sparse stream for the specified OneDrive item.
    /// </summary>
    /// <param name="graphClient">The authenticated Graph client.</param>
    /// <param name="driveId">The drive identifier.</param>
    /// <param name="itemId">The drive item identifier.</param>
    /// <param name="accessMode">The requested access mode.</param>
    /// <param name="cancellationToken">A token that can cancel initialization.</param>
    /// <returns>A fully initialized sparse stream.</returns>
    public static async Task<OneDriveWriteStream> CreateAsync(
        GraphServiceClient graphClient,
        string driveId,
        string itemId,
        FileAccess accessMode,
        CancellationToken cancellationToken = default)
    {
        if (accessMode != FileAccess.Read && accessMode != FileAccess.Write && accessMode != FileAccess.ReadWrite)
            throw new ArgumentOutOfRangeException(nameof(accessMode));

        long sourceLength = 0;

        if (accessMode != FileAccess.Write)
        {
            var driveItem = await graphClient.Drives[driveId].Items[itemId].GetAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            sourceLength = driveItem?.Size ?? 0;
        }

        return new OneDriveWriteStream(graphClient, driveId, itemId, accessMode, sourceLength);
    }

    /// <inheritdoc />
    public override bool CanRead => !_disposed && _accessMode != FileAccess.Write;
    /// <inheritdoc />
    public override bool CanSeek => !_disposed;
    /// <inheritdoc />
    public override bool CanWrite => !_disposed && _accessMode != FileAccess.Read;

    /// <inheritdoc />
    public override long Length
    {
        get
        {
            ThrowIfDisposed();
            return _length;
        }
    }

    /// <inheritdoc />
    public override long Position
    {
        get
        {
            ThrowIfDisposed();
            return _position;
        }
        set
        {
            ThrowIfDisposed();
            if (value < 0)
                throw new ArgumentOutOfRangeException(nameof(value));
            _position = value;
        }
    }

    /// <inheritdoc />
    public override void Flush()
    {
        ThrowIfDisposed();
    }

    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count)
    {
        ThrowIfDisposed();
        if (!CanRead)
            throw new NotSupportedException("Stream does not support reading.");

        ValidateBufferArgs(buffer, offset, count);

        if (count == 0 || _position >= _length)
            return 0;

        var remaining = (int)Math.Min(count, _length - _position);
        var totalRead = 0;

        while (remaining > 0)
        {
            var blockIndex = _position / BlockSize;
            var blockOffset = (int)(_position % BlockSize);
            var bytesThisBlock = Math.Min(remaining, BlockSize - blockOffset);
            var block = GetBlock(blockIndex);

            Buffer.BlockCopy(block, blockOffset, buffer, offset + totalRead, bytesThisBlock);

            _position += bytesThisBlock;
            totalRead += bytesThisBlock;
            remaining -= bytesThisBlock;
        }

        return totalRead;
    }

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin)
    {
        ThrowIfDisposed();

        var nextPosition = origin switch
        {
            SeekOrigin.Begin => offset,
            SeekOrigin.Current => _position + offset,
            SeekOrigin.End => _length + offset,
            _ => throw new ArgumentOutOfRangeException(nameof(origin))
        };

        if (nextPosition < 0)
            throw new IOException("Attempted to seek before the beginning of the stream.");

        _position = nextPosition;
        return _position;
    }

    /// <inheritdoc />
    public override void SetLength(long value)
    {
        ThrowIfDisposed();
        if (!CanWrite)
            throw new NotSupportedException("Stream does not support writing.");
        if (value < 0)
            throw new ArgumentOutOfRangeException(nameof(value));

        if (value < _length)
        {
            var newLastBlock = value == 0 ? -1 : (value - 1) / BlockSize;

            var toRemove = new List<long>();
            foreach (var key in _blocks.Keys)
            {
                if (key > newLastBlock)
                    toRemove.Add(key);
            }

            foreach (var key in toRemove)
                _blocks.Remove(key);

            if (newLastBlock >= 0 && _blocks.TryGetValue(newLastBlock, out var lastBlock))
            {
                var keepLength = (int)(value % BlockSize);
                if (keepLength == 0)
                    keepLength = BlockSize;

                Array.Clear(lastBlock, keepLength, BlockSize - keepLength);
            }
        }

        _length = value;
        if (_position > _length)
            _position = _length;
        _committed = false;
    }

    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        ThrowIfDisposed();
        if (!CanWrite)
            throw new NotSupportedException("Stream does not support writing.");

        ValidateBufferArgs(buffer, offset, count);

        var remaining = count;
        var sourceOffset = offset;

        while (remaining > 0)
        {
            var blockIndex = _position / BlockSize;
            var blockOffset = (int)(_position % BlockSize);
            var bytesThisBlock = Math.Min(remaining, BlockSize - blockOffset);
            var block = GetWritableBlock(blockIndex);

            Buffer.BlockCopy(buffer, sourceOffset, block, blockOffset, bytesThisBlock);

            _position += bytesThisBlock;
            sourceOffset += bytesThisBlock;
            remaining -= bytesThisBlock;
        }

        if (_position > _length)
            _length = _position;

        _committed = false;
    }

    /// <summary>
    /// Uploads the current in-memory sparse content to OneDrive.
    /// </summary>
    /// <param name="cancellationToken">A token that can cancel the upload.</param>
    public async Task CommitAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();

        if (_accessMode == FileAccess.Read)
            throw new InvalidOperationException("Cannot commit a read-only stream.");

        if (_committed)
            return;

        var savedPosition = _position;
        _position = 0;

        try
        {
            if (_length <= LargeFileThreshold)
            {
                await _graphClient.Drives[_driveId].Items[_itemId].Content
                    .PutAsync(this, cancellationToken: cancellationToken)
                    .ConfigureAwait(false);
            }
            else
            {
                var uploadSessionBody = new DriveUpload.CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", "replace" }
                        }
                    }
                };

                var uploadSession = await _graphClient.Drives[_driveId].Items[_itemId]
                    .CreateUploadSession
                    .PostAsync(uploadSessionBody, cancellationToken: cancellationToken)
                    .ConfigureAwait(false);

                var uploadTask = new LargeFileUploadTask<DriveItem>(uploadSession!, this, UploadSliceSize, _graphClient.RequestAdapter);
                await uploadTask.UploadAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            }

            _committed = true;
        }
        finally
        {
            if (!_disposed)
                _position = Math.Min(savedPosition, _length);
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing && !_committed && _accessMode != FileAccess.Read)
        {
            try
            {
                CommitAsync(CancellationToken.None).GetAwaiter().GetResult();
            }
            catch
            {
                // Ignored on dispose.
            }
        }

        _disposed = true;
        base.Dispose(disposing);
    }

    private byte[] GetBlock(long blockIndex)
    {
        if (_blocks.TryGetValue(blockIndex, out var cached))
            return cached;

        var block = new byte[BlockSize];

        if (_accessMode != FileAccess.Write)
        {
            var loaded = LoadBlockFromRemote(blockIndex).GetAwaiter().GetResult();
            Buffer.BlockCopy(loaded, 0, block, 0, BlockSize);
        }

        _blocks[blockIndex] = block;
        return block;
    }

    private byte[] GetWritableBlock(long blockIndex)
    {
        if (_blocks.TryGetValue(blockIndex, out var cached))
            return cached;

        var block = new byte[BlockSize];

        if (_accessMode == FileAccess.ReadWrite && blockIndex * BlockSize < _sourceLength)
        {
            var loaded = LoadBlockFromRemote(blockIndex).GetAwaiter().GetResult();
            Buffer.BlockCopy(loaded, 0, block, 0, BlockSize);
        }

        _blocks[blockIndex] = block;
        return block;
    }

    private async Task<byte[]> LoadBlockFromRemote(long blockIndex)
    {
        var start = blockIndex * BlockSize;
        var block = new byte[BlockSize];

        if (start >= _sourceLength)
            return block;

        var bytesToRead = (int)Math.Min(BlockSize, _sourceLength - start);
        var rangeHeader = $"bytes={start}-{start + bytesToRead - 1}";

        using var remote = await _graphClient.Drives[_driveId].Items[_itemId].Content
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.Headers.Add("Range", rangeHeader);
            }, cancellationToken: CancellationToken.None)
            .ConfigureAwait(false);

        if (remote is null)
            return block;

        var offset = 0;
        while (offset < bytesToRead)
        {
            var read = await ReadRemoteBlockAsync(remote, block, offset, bytesToRead - offset).ConfigureAwait(false);
            if (read == 0)
                break;
            offset += read;
        }

        return block;
    }

#if NETSTANDARD2_0
    private static Task<int> ReadRemoteBlockAsync(Stream remote, byte[] buffer, int offset, int count)
        => remote.ReadAsync(buffer, offset, count, CancellationToken.None);
#else
    private static ValueTask<int> ReadRemoteBlockAsync(Stream remote, byte[] buffer, int offset, int count)
        => remote.ReadAsync(buffer.AsMemory(offset, count), CancellationToken.None);
#endif

    private static void ValidateBufferArgs(byte[] buffer, int offset, int count)
    {
        if (buffer is null)
            throw new ArgumentNullException(nameof(buffer));
        if (offset < 0)
            throw new ArgumentOutOfRangeException(nameof(offset));
        if (count < 0)
            throw new ArgumentOutOfRangeException(nameof(count));
        if (buffer.Length - offset < count)
            throw new ArgumentException("Invalid offset and length.");
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(OneDriveWriteStream));
    }
}
