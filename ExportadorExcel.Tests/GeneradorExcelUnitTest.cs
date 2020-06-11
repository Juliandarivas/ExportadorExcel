using System;
using System.Collections.Generic;
using System.IO;
using ExportadorExcel.Infraestructura;
using ExportadorExcel.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace ExportadorExcel.Tests
{
    [TestClass]
    public class GeneradorExcelUnitTest
    {
        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeUnaListaPlanaEnLaHojaDatos()
        {
            var datos = new List<DatosFake>()
            {
                new DatosFake("Felipe", "Díaz", 28, new DateTime(2020, 4, 3)),
                new DatosFake("Carlos", "Londoño", 30, DateTime.Now)
            };

            var sheet = ObtenerHojaDeCalculo(datos);

            AcertarDatosDeHojaDeCalculo(sheet, datos);
        }

        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeMultiplesListasPlanasEnMultiplesHojasDatos()
        {
            var datos1 = CrearDatos();

            var datos2 = CrearDatos();

            var generadorReporte = new GeneradorReporte();

            var bytesResultantes = generadorReporte
                .Generar(Resources.PruebaGenerador, "Datos;DatosPrueba", datos1, datos2);


            ExcelWorksheet hojaDeCalculoDatos = ObtenerHojaDeCalculo(bytesResultantes, "Datos");
            ExcelWorksheet hojaDeCalculoDatosPrueba = ObtenerHojaDeCalculo(bytesResultantes, "DatosPrueba");

            AcertarDatosDeHojaDeCalculo(hojaDeCalculoDatos, datos1);
            AcertarDatosDeHojaDeCalculo(hojaDeCalculoDatosPrueba, datos2);
        }

        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeMultiplesListasPlanasEnMultiplesHojasDatos_SolicitandoArrayDeBytes()
        {
            var datos = CrearDatos();

            var datos2 = CrearDatos();

            var generadorReporte = new GeneradorReporte();

            Byte[] bytesResultantes = generadorReporte
                .Generar(Resources.PruebaGenerador, "Datos;DatosPrueba", datos, datos2);

            ExcelWorksheet hojaDeCalculoDatos = ObtenerHojaDeCalculo(bytesResultantes, "Datos");
            ExcelWorksheet hojaDeCalculoDatosPrueba = ObtenerHojaDeCalculo(bytesResultantes, "DatosPrueba");

            AcertarDatosDeHojaDeCalculo(hojaDeCalculoDatos, datos);
            AcertarDatosDeHojaDeCalculo(hojaDeCalculoDatosPrueba, datos2);
        }

        private static void AcertarDatosDeHojaDeCalculo(ExcelWorksheet sheet, List<DatosFake> datos)
        {
            Assert.AreEqual("Nombre", sheet.GetValue<string>(1, 1));
            Assert.AreEqual("Apellido", sheet.GetValue<string>(1, 2));
            Assert.AreEqual("Edad", sheet.GetValue<string>(1, 3));
            Assert.AreEqual("FechaNacimiento", sheet.GetValue<string>(1, 4));

            Assert.AreEqual("Felipe", sheet.GetValue<string>(2, 1));
            Assert.AreEqual(datos[0].FechaNacimiento, sheet.GetValue<DateTime>(2, 4));

            Assert.AreEqual("Carlos", sheet.GetValue<string>(3, 1));
        }

        private static List<DatosFake> CrearDatos()
        {
            return new List<DatosFake>()
            {
                new DatosFake("Felipe", "Díaz", 28, new DateTime(2020, 4, 3)),
                new DatosFake("Carlos", "Londoño", 30, DateTime.Now)
            };
        }

        private static ExcelWorksheet ObtenerHojaDeCalculo<T>(List<T> datos)
        {
            Byte[] bytesResultantes = ObtenerBytesResultantes(datos);
            return ObtenerHojaDeCalculo(bytesResultantes, "Datos");
        }

        private static ExcelWorksheet ObtenerHojaDeCalculo(byte[] bytesResultantes, string hojaDeCalculo)
        {
            using var straemResultante = new MemoryStream(bytesResultantes);
            var excel = new ExcelPackage(straemResultante);
            return excel.Workbook.Worksheets[hojaDeCalculo];
        }

        private static byte[] ObtenerBytesResultantes<T>(List<T> datos)
        {
            var generadorReporte = new GeneradorReporte();
            return generadorReporte.Generar(Resources.PruebaGenerador, datos);
        }
    }

    public class DatosFake
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int Edad { get; set; }
        public DateTime FechaNacimiento { get; set; }

        public DatosFake(string nombre, string apellido, int edad, DateTime fechaNacimiento)
        {
            Nombre = nombre;
            Apellido = apellido;
            Edad = edad;
            FechaNacimiento = fechaNacimiento;
        }
    }

    public class DatosFake2
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int Edad { get; set; }
        public DateTime FechaNacimiento { get; set; }

        public DatosFake2(string nombre, string apellido, int edad, DateTime fechaNacimiento)
        {
            Nombre = nombre;
            Apellido = apellido;
            Edad = edad;
            FechaNacimiento = fechaNacimiento;
        }
    }
}
