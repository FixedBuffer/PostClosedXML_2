using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace PostClosedXML2
{
  class Program
  {
    static IEnumerable<XLColor> GetColors()
    {
      yield return XLColor.Red;
      yield return XLColor.Amber;
      yield return XLColor.AppleGreen;
      yield return XLColor.AtomicTangerine;
      yield return XLColor.BallBlue;
      yield return XLColor.Bittersweet;
      yield return XLColor.CalPolyPomonaGreen;
      yield return XLColor.CosmicLatte;
      yield return XLColor.DimGray;
    }


    static void Main(string[] args)
    {
      using (var workbook = new XLWorkbook())
      {
        //Generamos la hoja
        var worksheet = workbook.Worksheets.Add("FixedBuffer");
        //Generamos la cabecera
        worksheet.Cell("A1").Value = "Nombre";
        worksheet.Cell("B1").Value = "Color";

        //-----------Le damos el formato a la cabecera----------------
        var rango = worksheet.Range("A1:B1");
        rango.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick); //Generamos las lineas exteriores
        rango.Style.Border.SetInsideBorder(XLBorderStyleValues.Medium); //Generamos las lineas interiores
        rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; //Alineamos horizontalmente
        rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;  //Alineamos verticalmente
        rango.Style.Font.FontSize = 14; //Indicamos el tamaño de la fuente
        rango.Style.Fill.BackgroundColor = XLColor.AliceBlue; //Indicamos el color de background
        worksheet.Columns(1, 2).AdjustToContents();

        int nRow = 2;
        //Genero la tabla de colores
        foreach(var color in GetColors())
        {
          worksheet.Cell(nRow, 1).Value = color.ToString();
          worksheet.Cell(nRow, 2).Style.Fill.BackgroundColor = color;
          nRow++;
        }

        //Aplico los formatos
        rango = worksheet.Range(2, 1, nRow-1, 2);
        rango.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
        rango.Style.Border.SetInsideBorder(XLBorderStyleValues.Medium);
        rango.Style.Font.SetFontName("Liberation Mono"); //Utilizo una fuente monoespacio
        rango.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        rango.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        workbook.SaveAs("CellFormating.xlsx");
      }
    }
  }
}
