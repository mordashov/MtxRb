using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mtx
{
    class Program
    {

        //private string _mainConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Environment.CurrentDirectory + "\\mtx.accdb";
        private static string _mainConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Dropbox\Task\Ryghova\work\mtx.accdb";

        static void Main(string[] args)
        {
            string connectionString = _mainConnectionString;
            string sql = @"
                    SELECT Eksvsp.Код
                    , Eksvsp.ФИО
                    , Eksvsp.ТАБЕЛЬНЫЙ_НОМЕР
                    , Eksvsp.ДОЛЖНОСТЬ
                    , Eksvsp.КОД_ПОДРАЗДЕЛЕНИЯ
                    , Eksvsp.НАИМЕНОВАНИЕ_ПОДРАЗДЕЛЕНИЯ
                    , Eksvsp.КОД_ТБ
                    , Eksvsp.НАИМЕНОВАНИЕ_ТБ
                    , Eksvsp.ЛОГИН
                    , Eksvsp.СТАТУС
                    , Eksvsp.РОЛИ
                    , Eksvsp.tn 
                    FROM Eksvsp 
                    INNER JOIN sap ON Eksvsp.tn = sap.[табельный номер]
                    WHERE Eksvsp.РОЛИ IS NOT NULL AND Eksvsp.РОЛИ <> '' ;";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, connectionString);
            DataSet ds = new DataSet();
            try
            {
                da.Fill(ds, "t");
            }
            catch (Exception)
            {
                Console.WriteLine("Не могу получить доступ к базе данных!");
                Environment.Exit(0);
            }

            OleDbConnection connection = new OleDbConnection
            {
                ConnectionString = _mainConnectionString
            };

            connection.Open();

            OleDbCommand command = new OleDbCommand(
                $@"DELETE FROM EksvspNew", connection);

            da.InsertCommand = command;
            da.InsertCommand.ExecuteNonQuery();
            int i = 0;
            int count = ds.Tables["t"].Rows.Count;
            foreach (DataRow fieldRow in ds.Tables["t"].Rows)
            {
                string id = fieldRow["Код"].ToString();
                string[] roles = fieldRow["РОЛИ"].ToString().Split('|');



                foreach (var role in roles)
                {
                    string roleTrim = role.Trim();

                    command = new OleDbCommand(
                        $@"INSERT INTO EksvspNew ( 
                            ФИО
                            ,ТАБЕЛЬНЫЙ_НОМЕР
                            ,ДОЛЖНОСТЬ
                            , КОД_ПОДРАЗДЕЛЕНИЯ
                            , НАИМЕНОВАНИЕ_ПОДРАЗДЕЛЕНИЯ
                            , КОД_ТБ, НАИМЕНОВАНИЕ_ТБ
                            , ЛОГИН
                            , СТАТУС
                            , tn
                            , РОЛИ ) 
                        SELECT Eksvsp.ФИО
                        , Eksvsp.ТАБЕЛЬНЫЙ_НОМЕР
                        , Eksvsp.ДОЛЖНОСТЬ
                        , Eksvsp.КОД_ПОДРАЗДЕЛЕНИЯ
                        , Eksvsp.НАИМЕНОВАНИЕ_ПОДРАЗДЕЛЕНИЯ
                        , Eksvsp.КОД_ТБ
                        , Eksvsp.НАИМЕНОВАНИЕ_ТБ
                        , Eksvsp.ЛОГИН
                        , Eksvsp.СТАТУС
                        , Eksvsp.tn
                        , ""{roleTrim}"" 
                        FROM Eksvsp 
                        WHERE Код = {id}", connection);

                    da.InsertCommand = command;
                    da.InsertCommand.ExecuteNonQuery();
                }
                Console.WriteLine($"Обработано {i} строк из {count}");
                i++;
            }
            connection.Close();

        }
    }
}
