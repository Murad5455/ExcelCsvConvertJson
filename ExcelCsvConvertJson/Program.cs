using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    static void Main()
    {

        string excelFilePath = "C:\\Excel\\Test.xlsx";

        string jsonFilePath = "C:\\Excel\\TestJson.txt";

        List<Dictionary<string, object>> excelData = ReadExcelData(excelFilePath);

        WriteJsonFile(jsonFilePath, excelData);

        Console.WriteLine("Melumatlar text faylina yazildi");
    }
    static List<Dictionary<string, object>> ReadExcelData(string filePath)
    {
      
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        List<Dictionary<string, object>> excelData = new List<Dictionary<string, object>>();
      
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            if (package.Workbook.Worksheets.Count > 0)
            {
                var worksheet = package.Workbook.Worksheets[0]; 

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= rowCount; row++) 
                {
                    Dictionary<string, object> rowData = new Dictionary<string, object>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        string key = worksheet.Cells[1, col].Value.ToString();
                        object value = worksheet.Cells[row, col].Value; 

                        rowData.Add(key, value);
                    }

                    excelData.Add(rowData);
                }
            }
            else
            {
                Console.WriteLine("Excel fayli tapilmadi.");
            }
        }

        return excelData;
    }
    
    static void WriteJsonFile(string filePath, object data)
    {
        using (StreamWriter file = new StreamWriter(filePath))
        {
            string jsonData = JsonConvert.SerializeObject(data, Formatting.Indented);
            file.Write(jsonData);
        }
    }


}
