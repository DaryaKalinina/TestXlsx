using System;
using System.IO;
using System.Text;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Core.ExcelPackage;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace TestXlsx
{
    
    class Program
    {
        static void Main(string[] args)
        {

            string xlsxPath = @"../../../file/Price_Kompjuternaja_perifеrija_2018_07_10.xlsx";
            string txtPath = @"../../../file/Price_Kompjuternaja_perifеrija_2018_07_10.txt";
            //для txt
            StreamWriter newFile = new StreamWriter(txtPath, false, Encoding.Default);

            try
            {

                using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {

                    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                    {
                        WorkbookPart workbookPart = doc.WorkbookPart;
                        SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                        SharedStringTable sst = sstpart.SharedStringTable;

                        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                        Worksheet sheet = worksheetPart.Worksheet;
                        //количество столбцев 
                        //var cells = sheet.Descendants<Cell>();
                        //количество строк
                        var rows = sheet.Descendants<Row>();

                        foreach (Row row in rows)
                        {
                            int i = 1;
                            ArrayList infoAboutProduct = new ArrayList();
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                                {
                                    int ssid = int.Parse(c.CellValue.Text);
                                    string str = sst.ChildElements[ssid].InnerText;

                                    infoAboutProduct.Add(str);
                                }
                                else if (c.CellValue != null)
                                {
                                    infoAboutProduct.Add(c.CellValue.Text);
                                }

                            }
                            try
                            {

                                newFile.WriteLine("Код: {0} Артикул: {1} Наименование: {2} Производитель: {3} Ед. изм. : {4} Цена : {5};",
                                infoAboutProduct[0], infoAboutProduct[1], infoAboutProduct[2], infoAboutProduct[3], infoAboutProduct[4],
                                infoAboutProduct[5]);
                            }
                            catch
                            {
                                //игнорировать exception
                                continue;
                            }
                        }

                        newFile.Close();
                        doc.Close();

                    }

                }
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine("Файл не найден");
            }
            catch (Exception e)
            {
                 Console.WriteLine(e.Message);
            }

        }
    }
}


