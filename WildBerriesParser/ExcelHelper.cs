using Microsoft.Office.Interop.Excel;

namespace WildBerriesParser;

public class ExcelHelper : IDisposable
{
    private readonly Application _excel;
    private Workbook _workBook;
    private string _filePath;

    public ExcelHelper()
    {
        _excel = new Application();
    }
    
    public bool Open(string filePath, string keyValue)
    {
        try
        {
            if (File.Exists(filePath))
            {
                _workBook = _excel.Workbooks.Open(filePath);
                var sheet = _workBook.Sheets;
                var sheetPivot = (Worksheet)sheet.Add(Type.Missing, sheet[sheet.Count - 1], Type.Missing, Type.Missing);
                sheetPivot.Name = keyValue;
            }
            else
            {
                _workBook = _excel.Workbooks.Add();
                var sh = _workBook.Sheets;
                Worksheet sheetPivot = (Worksheet)sh.Add(sh[1], Type.Missing, Type.Missing, Type.Missing);
                sheetPivot.Name = keyValue;
                _filePath = filePath;
            }

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        return false;
    }
    
    public void Save()
    {
        if (!string.IsNullOrEmpty(_filePath))
        {
            _workBook.SaveAs(_filePath);
            _filePath = null;
        }
        else
        {
            _workBook.Save();
        }
    }
    
    public bool Set(string column, int row, object data)
    {
        try
        {
            ((Worksheet)_excel.ActiveSheet).Cells[row, column] = data;

            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        return false;
    }
    
    public void ColumnsAutoFit()
    {
        ((Worksheet)_excel.ActiveSheet).Columns.EntireColumn.AutoFit();
    }
    
    public void Dispose()
    {
        try
        {
            _workBook.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}