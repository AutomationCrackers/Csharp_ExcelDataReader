using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ExcelDataReader;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;

namespace Csharp_ExcelDataReader.ExcelDataReader
{
    public class ExcelDataReader
    {
        //Step 1 Find the Excel==========>Get the Excelfilepath with the Excelname
        public static string ExcelFilePath() 
        {
            string currentDirectoryPath = Environment.CurrentDirectory;
            string actualPath = currentDirectoryPath.Substring(0, currentDirectoryPath.LastIndexOf("bin"));
            string projectPath = new Uri(actualPath).LocalPath;
            string ExcelfilePath = projectPath + "\\Excelsheets\\TestData.xlsx";
            return ExcelfilePath;
        }

        //Step 2 ========>Storing all the Excel Values into Memory Collection DataTable
        /*Example here we have 
                                Excel filename - "TestData" 
                                Sheetname - "LoginCredentials"     */
        //Pass the Excelfile full path and sheetname as input to this method

        public static DataTable ExcelToDataTable(string filename, string sheetname) 
        {
            DataTable resultTable = new DataTable();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    //Get all the Tables
                    DataTableCollection table = result.Tables;

                    //Store it in the DataTable and Return
                    resultTable = table[sheetname];
                    return resultTable;
                }
            }
        }
        //Step3========>DataTable to String object Variable
        public static string ReadData(int rowNumber, string columnName, DataTable datacolumn)
        {
            try 
            {
                string data = datacolumn.Rows[rowNumber][columnName].ToString();
                return data.ToString();
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error in ReadData " + ex.Message);
                return null;
            }
        }
    }
}
