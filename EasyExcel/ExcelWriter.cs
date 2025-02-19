using System.Reflection;
using ClosedXML.Excel;

namespace EasyExcel;

public class ExcelWriter : IDisposable
{
    private readonly XLWorkbook _workbook = new();
    public int AbaIndex { get; set; } = 0;

    public void SetAbaIndex(int abaIndex)
    {
        if (abaIndex >= _workbook.Worksheets.Count)
            throw new Exception("Aba não encontrada.");

        AbaIndex = abaIndex;
    }

    public void AddAba(string abaName)
    {
        _workbook.Worksheets.Add(abaName);
    }

    public void CreateRow(int index, params string?[] values)
    {
        if (AbaIndex >= _workbook.Worksheets.Count)
            throw new Exception("Aba não encontrada.");

        var aba = _workbook.Worksheets.ElementAt(AbaIndex);

        var row = aba.Row(index);
        int column = 1;
        foreach (var value in values)
        {
            var cell = row.Cell(column);
            cell.Value = value;
            cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            cell.Style.Border.SetOutsideBorderColor(XLColor.Black);
            column++;
        }
    }

    public void CreateRowCustom(int index, params (string value, XLColor borderColor)[] values)
    {
        if (AbaIndex >= _workbook.Worksheets.Count)
            throw new Exception("Aba não encontrada.");

        var aba = _workbook.Worksheets.ElementAt(AbaIndex);

        var row = aba.Row(index);
        int column = 1;
        foreach (var value in values)
        {
            var cell = row.Cell(column);
            cell.Value = value.value;
            cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            cell.Style.Border.SetOutsideBorderColor(value.borderColor);
            column++;
        }
    }

    public Stream Export<T>(IEnumerable<T> entities, bool noHeader = false)
    {
        AddAba(typeof(T).Name);
        var rowLastIndex = 1;

        // Header
        var properties = typeof(T).GetProperties().Where(p => p.GetCustomAttribute<ImportIgnore>() == null);
        if (!noHeader)
        {
            CreateRow(rowLastIndex, properties.Select(x => x.Name).ToArray());
            rowLastIndex++;
        }


        // Body
        foreach (var entity in entities)
        {
            CreateRow(rowLastIndex, properties.Select(x =>
            {
                var obj = x.GetValue(entity);
                if (obj is bool b)
                    return b ? "S" : "N";

                return Convert.ToString(obj) ?? "";
            }).ToArray());

            rowLastIndex++;
        }

        return Export();
    }

    public Stream Export()
    {
        var stream = new MemoryStream();
        _workbook.SaveAs(stream);
        stream.Seek(0, SeekOrigin.Begin);
        return stream;
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}