using Microsoft.Data.Sqlite;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Programm;

class Program
{
    private static string path = @"c:\s9\.txt";
    public static void Main(string[] args)
    {
        Worckbok worckbok = new Worckbok();
        SQlconnection sQlconnection = new SQlconnection();
        DataFormatter dataFormatter = new DataFormatter();
        StreamWriter sw = new StreamWriter(path);
        string saveEmploye_id = null;
        worckbok.WorkbookCon(out ISheet sheet);
        for (int i = 0; i < sheet.LastRowNum; i++)
        {
            string Cell_FIO = dataFormatter.FormatCellValue(sheet.GetRow(i).GetCell(0)).TrimEnd().TrimStart();
           saveEmploye_id = sQlconnection.SQl_Select(Cell_FIO, saveEmploye_id, dataFormatter,sheet,i,sw);
        }
        sw.Close();
    }
}
