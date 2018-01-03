using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ConvertXMLtoXLS.Logic
{
    public class FileReader
    {

        public List<CustomTuple> readDataFromXml()
        {
           
            XElement xelement = XElement.Load(@"Path");
            IEnumerable<XElement> employees = xelement.Elements();
            // Read the entire XML
            var sTable = new List<CustomTuple>();
    

            foreach (var employee in employees)
            {
                sTable.Add(new CustomTuple(
                    employee.Element("EmpId").Value,
                    employee.Element("Name").Value,
                    employee.Element("Sex").Value,
                    employee.Element("Phone").Value,
                    employee.Element("Address").Value
                    ));
                
            }
            return sTable;
        }

        public void SaveDataToExcel(List<CustomTuple> oDataFromXml)
        {
            var connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Path\ExcelWorkBook.xls;Extended Properties=Excel 8.0";
            using (var excelConnection = new OleDbConnection(connectionString))
            {
                if (excelConnection.State != ConnectionState.Open) { excelConnection.Open(); }

                var sqlText = "CREATE TABLE YourTableNameHere ([EmpId] VARCHAR(100), [Name] VARCHAR(100), [Sex] VARCHAR(100), [Phone] VARCHAR(100), [Address] VARCHAR(100))";

               //Create worksheet
                var command = new OleDbCommand(sqlText, excelConnection);
                command.ExecuteNonQuery();
                //insert data to worksheet           
                var commandText = $"Insert Into YourTableNameHere (EmpId, Name, Sex, Phone,Address) Values (?,?,?,?,?)";
                var commandx = new OleDbCommand(commandText, excelConnection);
                {
                    for (int i = 0; i < oDataFromXml.Count; i++)
                    {
                        commandx.Parameters.Add("EmpId", OleDbType.VarChar, 100);
                        commandx.Parameters.Add("Name", OleDbType.VarWChar, 100);
                        commandx.Parameters.Add("Sex", OleDbType.VarWChar, 100);
                        commandx.Parameters.Add("Phone", OleDbType.VarWChar, 100);
                        commandx.Parameters.Add("Address", OleDbType.VarWChar, 100);
                        commandx.Parameters[0].Value = oDataFromXml[i].Item1;
                        commandx.Parameters[1].Value = oDataFromXml[i].Item2;
                        commandx.Parameters[2].Value = oDataFromXml[i].Item3;
                        commandx.Parameters[3].Value = oDataFromXml[i].Item4;
                        commandx.Parameters[4].Value = oDataFromXml[i].Item5;
                        commandx.Prepare();
                        commandx.ExecuteNonQuery();
                    }
                }

                }

            }


        /// <summary>
        /// Insert new rows to existing excel - xls.
        /// HDR=YES which means thath we gonna skip first row, thath contains column names. 
        /// </summary>
        /// <param name="sSheetName"></param>
        /// <param name="sFirstParam"></param>
        /// <param name="sSecondParam"></param>
        /// <param name="sThirdParam"></param>
        public static void InsertDataToExistingFile(string sSheetName, string sFirstParam, string sSecondParam, string sThirdParam)
        {
            string connectionString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""",
                @"C:\Path\filename.xls");

            using (var excelConnection = new OleDbConnection(connectionString))
            {
                if (excelConnection.State != ConnectionState.Open) { excelConnection.Open(); }

                var sqlText = "CREATE TABLE [" + sSheetName + "$] ([1Column] VARCHAR(100), [2Column] VARCHAR(100), [3Column] VARCHAR(100))";

                //Create new column if not exist.
                var command = new OleDbCommand(sqlText, excelConnection);
                command.ExecuteNonQuery();

                var commandText = $"INSERT Into [" + sSheetName + "$] (1Column,2Column,3Column) Values ('" + sFirstParam + "','" + sSecondParam + "','" + sThirdParam + "')";
                OleDbCommand cmd = new OleDbCommand(commandText, excelConnection);
                cmd.Prepare();
                cmd.ExecuteNonQuery();
                excelConnection.Close();
            }
        }


    }
    }

