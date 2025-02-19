using XlsxHelper;

namespace EasyExcel;

public sealed class ExcelReader : IDisposable
{
    private readonly Workbook _workbook;
    private readonly FileStream _fileStream;

    public ExcelReader(Stream stream)
    {
        var filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

        _fileStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        stream.CopyTo(_fileStream);
        
        _workbook = XlsxReader.OpenWorkbook(_fileStream);
    }

    public IEnumerable<Dictionary<string, string?>> GetRows(int aba = 0,  int headerRow = 1, int startRow = 2)
    {
        var worksheet = _workbook.Worksheets.ElementAtOrDefault(aba);
        if (worksheet == null)
            throw new Exception("Excel worksheet not found!");
        
        var listHeader = new List<(ColumnName columnName, string? value)>();
        var rowData = worksheet.WorksheetReader.ToArray();
        var rows = new Dictionary<string, string?>[rowData.Length];
        int rowIndex = 0;
        foreach (var row in worksheet.WorksheetReader)
        {
            if (row.RowNumber == headerRow)
            {
                var r = new Dictionary<string, string?>();
                foreach (var cell in row.Cells)
                {
                    listHeader.Add((cell.ColumnName, cell.CellValue));
                    r.Add(cell.CellValue ?? cell.ColumnName.ToString(), cell.CellValue);
                }
                rows[rowIndex] = r;
            }

            if (row.RowNumber >= startRow)
            {
                var r = new Dictionary<string, string?>();

                var count = listHeader.Count;
                for (int i = 0; i < count; i++)
                {
                    var cellValue = row.Any(x => x.ColumnName == listHeader[i].columnName) ? row[listHeader[i].columnName].CellValue : null;
                    r.Add(listHeader[i].value ?? i.ToString(), cellValue);
                }

                rows[rowIndex] = r;
            }
            rowIndex++;
        }

        return rows;
    }

    public void Dispose()
    {
        _fileStream.Dispose();
        _workbook.Dispose();
    }
}