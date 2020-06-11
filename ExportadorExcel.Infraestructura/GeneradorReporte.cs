using ExportadorExcel.Infraestructura.Interfaces;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExportadorExcel.Infraestructura
{
    public class GeneradorReporte : IGeneradorReporte
    {
        public byte[] Generar<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos") 
        {
            Stream stream = new MemoryStream(recurso);
            return GenerarArrayDeBytesExcel(stream, datos, nombreHojaDatos);
        }

        public byte[] Generar<T>(byte[] recurso, string hojasDeCalculo, params IEnumerable<T>[] datosHojasDeCalculos)
        {
            Stream stream = new MemoryStream(recurso);
            return Generar(stream, hojasDeCalculo, datosHojasDeCalculos);
        }
        
        private byte[] Generar<T>(Stream excelPlantilla, string hojasDeCalculo, params IEnumerable<T>[] datosHojasDeCalculos)
        {
            ExcelPackage excel = new ExcelPackage(excelPlantilla);

            string[] excelsheets = hojasDeCalculo.Split(';');

            for (int i = 0; i < datosHojasDeCalculos.Length; i++)
            {
                ExcelWorksheet sheet = excel.Workbook.Worksheets[excelsheets[i]];

                CargarDatos(datosHojasDeCalculos[i], sheet);
                ConvertirPropiedadesEnColumnas<T>(sheet);
            }

            return excel.GetAsByteArray();
        }

        private static byte[] GenerarArrayDeBytesExcel<T>(Stream excelPlantilla, IEnumerable<T> datos, string nombreHojaDatos)
        {
            var excel = new ExcelPackage(excelPlantilla);
            var sheet = excel.Workbook.Worksheets[nombreHojaDatos];

            CargarDatos(datos, sheet);
            ConvertirPropiedadesEnColumnas<T>(sheet);

            return excel.GetAsByteArray();
        }

        private static void CargarDatos<T>(IEnumerable<T> datos, ExcelWorksheet sheet)
        {
            sheet.InsertRow(2, datos.Count() - 1);
            sheet.Cells["A2"].LoadFromCollection(datos);
        }

        private static void ConvertirPropiedadesEnColumnas<T>(ExcelWorksheet sheet)
        {
            var indiceColumna = 1;
            foreach (var propiedad in typeof(T).GetProperties())
            {
                if (propiedad.PropertyType == typeof(DateTime))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                sheet.SetValue(1, indiceColumna++, propiedad.Name);
            }
        }
    }
}
