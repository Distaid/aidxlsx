using Aid.Xlsx.Cells;
using Aid.Xlsx.Columns;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Rows;

namespace Aid.Xlsx.Worksheets;

internal sealed class XlsxWorksheetInternals : IDisposable
{
    private bool _disposed;

    internal required XlsxCellsCollection CellsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XlsxColumnsCollection ColumnsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XlsxRowsCollection RowsCollection
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    internal required XlsxRanges MergedRanges
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        init;
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        CellsCollection.ValueSlice.DereferenceSlice();
        CellsCollection.Clear();
        ColumnsCollection.Clear();
        RowsCollection.Clear();
        MergedRanges.RemoveAll();
        _disposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(XlsxWorksheetInternals));
    }
}
