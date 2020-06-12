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
        private GeneradorReporte _generadorReporte;

        [TestInitialize]
        public void Inicializar()
        {
            _generadorReporte = new GeneradorReporte();
        }

        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeUnaListaPlanaEnLaHojaDatos()
        {
            var datos = new List<DatosPersonales>()
            {
                new DatosPersonales("Felipe", "Díaz", 28, new DateTime(2020, 4, 3)),
                new DatosPersonales("Carlos", "Londoño", 30, DateTime.Now)
            };

            byte[] bytesActuales = ObtenerBytesResultantes(datos);

            var hojaActual = ObtenerHojaDeCalculo(bytesActuales, "Datos");

            AcertarDatosDeHojaDeCalculo(hojaActual, datos);
        }

        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeMultiplesListasPlanasEnMultiplesPaginas()
        {
            var datos = CrearDatos();

            var datosAdicionales = CrearDatos();

            var bytesActuales = _generadorReporte.Generar(Resources.PruebaGenerador, "Datos;DatosPrueba", datos, datosAdicionales);

            ExcelWorksheet hojaActual = ObtenerHojaDeCalculo(bytesActuales, "Datos");
            ExcelWorksheet hojaActualAdicional = ObtenerHojaDeCalculo(bytesActuales, "DatosPrueba");

            AcertarDatosDeHojaDeCalculo(hojaActual, datos);
            AcertarDatosDeHojaDeCalculo(hojaActualAdicional, datosAdicionales);
        }

        [Ignore]
        [TestMethod]
        public void Debe_GenerarExcel_LlenarLosDatosDeMultiplesListasPlanasDeDiferentesTiposDeDatosEnMultiplesPaginas()
        {
            var datosPersonales = CrearDatos();
            var datosFamiliares = CrearDatosFamiliares();
            Assert.Fail();
        }

        private void AcertarDatosDeHojaDeCalculo(ExcelWorksheet sheet, List<DatosPersonales> datos)
        {
            Assert.AreEqual("Nombre", sheet.GetValue<string>(1, 1));
            Assert.AreEqual("Apellido", sheet.GetValue<string>(1, 2));
            Assert.AreEqual("Edad", sheet.GetValue<string>(1, 3));
            Assert.AreEqual("FechaNacimiento", sheet.GetValue<string>(1, 4));

            Assert.AreEqual(datos[0].Nombre, sheet.GetValue<string>(2, 1));
            Assert.AreEqual(datos[1].Nombre, sheet.GetValue<string>(3, 1));
            Assert.AreEqual(datos[0].FechaNacimiento, sheet.GetValue<DateTime>(2, 4));
        }

        private List<DatosPersonales> CrearDatos()
        {
            return new List<DatosPersonales>()
            {
                new DatosPersonales("Felipe", "Díaz", 28, new DateTime(2020, 4, 3)),
                new DatosPersonales("Carlos", "Londoño", 30, DateTime.Now)
            };
        }

        private List<DatosFamiliares> CrearDatosFamiliares()
        {
            return new List<DatosFamiliares>()
            {
                new DatosFamiliares("Arturo", "Jimenez", 28, new DateTime(2020, 4, 3)),
                new DatosFamiliares("Ofelia", "Lozano", 30, DateTime.Now),
            };
        }

        private static ExcelWorksheet ObtenerHojaDeCalculo(byte[] bytesResultantes, string hojaDeCalculo)
        {
            using var straemResultante = new MemoryStream(bytesResultantes);
            var excel = new ExcelPackage(straemResultante);
            return excel.Workbook.Worksheets[hojaDeCalculo];
        }

        private byte[] ObtenerBytesResultantes<T>(List<T> datos)
        {
            return _generadorReporte.Generar(Resources.PruebaGenerador, datos);
        }
    }

    public class DatosPersonales
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int Edad { get; set; }
        public DateTime FechaNacimiento { get; set; }

        public DatosPersonales(string nombre, string apellido, int edad, DateTime fechaNacimiento)
        {
            Nombre = nombre;
            Apellido = apellido;
            Edad = edad;
            FechaNacimiento = fechaNacimiento;
        }
    }

    public class DatosFamiliares
    {
        public string Nombre { get; set; }
        public string Apellido { get; set; }
        public int Edad { get; set; }
        public DateTime FechaNacimiento { get; set; }

        public DatosFamiliares(string nombre, string apellido, int edad, DateTime fechaNacimiento)
        {
            Nombre = nombre;
            Apellido = apellido;
            Edad = edad;
            FechaNacimiento = fechaNacimiento;
        }
    }
}
