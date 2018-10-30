using ClosedXML.Excel;
using System;

namespace PostClosedXML2
{
  class Program
  {
    static void Main(string[] args)
    {
      using (var workbook = new XLWorkbook())
      {
        //Generamos la hoja
        var worksheet = workbook.Worksheets.Add("FixedBuffer");
        //Generamos la cabecera
        worksheet.Cell("A1").Value = "Nombre";
        worksheet.Cell("B1").Value = "Color";
        //Le damos el formato a la cabecera        
        worksheet.Range("A1:B1").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        worksheet.Range("A1:B1").Style.Border.TopBorder = XLBorderStyleValues.Thick;
        worksheet.Range("A1:B1").Style.Border.LeftBorder = XLBorderStyleValues.Thick;
        worksheet.Range("A1:B1").Style.Border.RightBorder = XLBorderStyleValues.Thick;
        worksheet.Range("A1:B1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        worksheet.Range("A1:B1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        worksheet.Range("A1:B1").Style.Font.FontSize = 8;



        workbook.SaveAs("HelloWorld.xlsx");
      }
    }
  }
}
