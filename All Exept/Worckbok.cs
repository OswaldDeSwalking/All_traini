using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Programm;

public class Worckbok
{
    private string path = @"c:\s9\med.xlsx";

    public void WorkbookCon( out ISheet sheetst)
    {
        
        var workbookStream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        IWorkbook workbook = new XSSFWorkbook(workbookStream);
        ISheet sheet = workbook.GetSheetAt(0);
        sheetst = sheet;
    }
}