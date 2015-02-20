using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Documents;

using System.Windows.Forms;
using System.Globalization;
using System.IO;
using Excel;
using ICSharpCode;
using System.Xml.Xsl;
using System.Xml;

using ExcelLibrary.SpreadSheet; 


using Microsoft.Office.Interop.Excel;


namespace Bdconnection
{

      class Ex
      {
          // чтение эксель и перегон его в dataset
         public DataSet FileExcelToDataSet(string filePath)
          {
              FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
              
              //1. Reading from a binary Excel file ('97-2003 format; *.xls)
              IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
              //3. DataSet - The result of each spreadsheet will be created in the result.Tables
              DataSet result = excelReader.AsDataSet();
              //4. DataSet - Create column names from first row
              excelReader.IsFirstRowAsColumnNames = true;
              DataSet result1 = excelReader.AsDataSet();

           

              //6. Free resources (IExcelDataReader is IDisposable)
              excelReader.Close();
              return result;





          }
          // выгрузка базы в эксель документ 
         public void WriteExcel(string filename, System.Data.DataTable vrachi, System.Data.DataTable sertif)
         {
             // формирование таблицы...
             System.Data.DataTable exTable = new System.Data.DataTable();
             System.Data.DataColumn col1 = new System.Data.DataColumn("IDDOKT");
            
             System.Data.DataColumn col5 = new System.Data.DataColumn("FIO");
             System.Data.DataColumn col6 = new System.Data.DataColumn("N_SERT");
             System.Data.DataColumn col7 = new System.Data.DataColumn("REG_NUM");
             System.Data.DataColumn col8 = new System.Data.DataColumn("PRVS_L");
             System.Data.DataColumn col9 = new System.Data.DataColumn("DATE_END");
             System.Data.DataColumn col10 = new System.Data.DataColumn("DATE_VN");

             System.Data.DataTable tableForExport = new System.Data.DataTable();
             tableForExport.Columns.Add(col1);
          
             tableForExport.Columns.Add(col5);
             tableForExport.Columns.Add(col6);
             tableForExport.Columns.Add(col7);
             tableForExport.Columns.Add(col8);
             tableForExport.Columns.Add(col9);
             tableForExport.Columns.Add(col10);
             int kol = 0;
             // занос данных  в таблицу 
             try
             {

                 for (int i = 0; i < vrachi.Rows.Count; i++)
                 {
                     kol++;
                     // найденные сертификаты врачей
                     DataRow[] findrows = sertif.Select("IDDOKT = " + vrachi.Rows[i][0]);

                    

                     foreach (DataRow x in findrows)
                     {
                         DateTime current = DateTime.Today;
                         if (((DateTime)x[2])>current) 
                           {  

                         string fio = vrachi.Rows[i][4] + " " + vrachi.Rows[i][5] + " " + vrachi.Rows[i][6];

                         string Date = Convert.ToString(x[2]);
                         tableForExport.Rows.Add
                             (
                         Convert.ToString(vrachi.Rows[i][0]),
                       
                           fio,
                          x[0],
                          x[1],
                         Convert.ToString(x[4]),
                        Date,
                        vrachi.Rows[i][9]
                            ); 
                         }
                     }
                 }
             }
             catch (Exception erors) { MessageBox.Show(erors.Message+" Ошибка при форм таблицы"+kol.ToString()); }
                

                 ExcelLibrary.SpreadSheet.Workbook workbook = new ExcelLibrary.SpreadSheet.Workbook();
                 ExcelLibrary.SpreadSheet.Worksheet worksheet = new ExcelLibrary.SpreadSheet.Worksheet("Первый Лист");
                
                

              

                 
                 
                 // звгрузка файла эксель данными о врачах и сертификатах из таблицы tableForExport
             // шапка 
                 worksheet.Cells[0,0] = new ExcelLibrary.SpreadSheet.Cell("Код врача");
                 worksheet.Cells[0,1] = new ExcelLibrary.SpreadSheet.Cell("ФИО");
                 worksheet.Cells[0,2] = new ExcelLibrary.SpreadSheet.Cell("Номер сертификата");
                 worksheet.Cells[0,3] = new ExcelLibrary.SpreadSheet.Cell("Регистрационный номер");
                 worksheet.Cells[0,4] = new ExcelLibrary.SpreadSheet.Cell("Специальность");
                 worksheet.Cells[0,5] = new ExcelLibrary.SpreadSheet.Cell("Дата окончания");
                 worksheet.Cells[0,6] = new ExcelLibrary.SpreadSheet.Cell("Дата внесения в реестр");
            // заполнение данными
             for (int i = 0; i < tableForExport.Rows.Count; i++) 
                 {
                     for (int j = 0; j < tableForExport.Columns.Count; j++) 
                     {
                         try
                         {
                            
                                 worksheet.Cells[i + 1, j] = new ExcelLibrary.SpreadSheet.Cell(

                                Convert.ToString(tableForExport.Rows[i][j]));
                            
                         }
                         catch (Exception exption) {MessageBox.Show(exption.Message+"это тут");}
                     }
                 }


              
                 workbook.Worksheets.Add(worksheet);
                 workbook.Save(filename);   

            
            
             }


     
     
      }
     
 
}


   
