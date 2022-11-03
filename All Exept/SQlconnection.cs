using Microsoft.Data.Sqlite;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Programm;

public class SQlconnection
{
   string connectionString = @"Data Source=C:\Program Files (x86)\Kodeks\Techexpert-Intranet\storage\isupb\db\arm.db";
   public string SQl_Select(string Cell_FIO, string employyee, DataFormatter dataFormatter, ISheet sheet, int i,
      StreamWriter streamWriter)
   {
      var connection = new SqliteConnection(connectionString);
      connection.Open();
      SqliteCommand cmd = connection.CreateCommand();
      cmd.CommandText = $"SELECT distinct EP.EmployeeId,(E.Surname||' '||E.Name||' '||E.Patronymic) FROM EmployeePositions EP join Employees E on EP.EmployeeId = E.Id where EP.DateOfEnding is null and EmploymentType =1;";
      SqliteDataReader sqliteDataReader = cmd.ExecuteReader();
      while (sqliteDataReader.Read())
      {
         string EmployeeId = sqliteDataReader.GetString(0);
         string SQL_FIO = sqliteDataReader.GetString(1).TrimEnd().TrimStart();
         if (Cell_FIO.Equals(SQL_FIO))
         {
            employyee = EmployeeId;
            break;
         }
         if (Cell_FIO.Equals("Медицинский осмотр"))
         {
            var datectyal=dataFormatter.FormatCellValue(sheet.GetRow(i).GetCell(2));
            Console.WriteLine("5"+"\t"+employyee+"\t"+"2"+"\t"+datectyal+"\t"+"1"+"\t"+"1");
            streamWriter.WriteLine("5"+"\t"+employyee+"\t"+"2"+"\t"+datectyal+"\t"+"1"+"\t"+"1");
            break;
         }
      }
      return employyee;
   }
}