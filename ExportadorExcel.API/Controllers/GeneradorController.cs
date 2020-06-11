using System;
using System.Collections.Generic;
using ExportadorExcel.API.Properties;
using ExportadorExcel.Infraestructura.Interfaces;
using ExportadorExcel.Modelos;
using Microsoft.AspNetCore.Mvc;

namespace ExportadorExcel.API.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class GeneradorController : ControllerBase
    {
        private readonly IGeneradorReporte _generadorReporte;

        public GeneradorController(IGeneradorReporte generadorReporte)
        {
            _generadorReporte = generadorReporte;
        }

        [HttpGet("Elemento")]
        public IActionResult ObtenerElemento()
        {
            List<Producto> productos = CrearProductos();

            byte[] bytesExcel = _generadorReporte.Generar(Resources.PlantillaElemento, productos);

            return File(bytesExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Productos.xlsx");
        }


        [HttpGet("Elementos")]
        public IActionResult ObtenerElementos()
        {
            List<Producto> productos = CrearProductos();

            List<Producto> productosAdicionales = CrearProductos();

            byte[] bytesExcel = _generadorReporte.Generar(Resources.PlantillaElementos, "DatosProductos;DatosAlmacenes",  productos, productosAdicionales);

            return File(bytesExcel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Productos.xlsx");
        }

        private List<Producto> CrearProductos()
        {
            return new List<Producto>
            {
                new Producto(1, "FNX0001", "Silla Oficina Paris", new DateTime(2020, 5, 1), 500_000),
                new Producto(2, "FNX0002", "Silla Oficina Dublin", new DateTime(2020, 5, 1), 3_000_000),
                new Producto(3, "FNX0003", "Silla Oficina Roma", new DateTime(2020, 5, 1), 400_000),
                new Producto(4, "FNX0004", "Silla Oficina Berlin", new DateTime(2020, 5, 1), 800_000),
                new Producto(5, "FNX0005", "Sillón El Cairo", new DateTime(2020, 5, 1), 1_200_000),
                new Producto(6, "FNX0006", "Escritorio sala de juntas New York", new DateTime(1980, 1, 1), 1_500_000),
                new Producto(7, "FNX0007", "Silla rimax", new DateTime(2020, 5, 1), 2_000_000),
                new Producto(8, "FNX0008", "Alfombra sala", new DateTime(2020, 5, 1), 3_100_000),
                new Producto(9, "FNX0009", "Mesa para comedor de vidrio", new DateTime(2020, 5, 1), 2_000_000),
                new Producto(10, "FNX0010", "Mesa redonda de vidrio", new DateTime(2020, 5, 1), 5_000_000),
            };
        }

        private List<Almacen> CrearAlmacenes()
        {
            return new List<Almacen>
            {
                new Almacen(1, "AEROPUERTO", "7444147", "Aeropuerto El Dorado Local 2-314"),
                new Almacen(2, "ANDINO", "7444149", "Cra 11 #	82-01 Local	- 159 C.C. Andino"),
                new Almacen(3, "ATLANTIS", "7444151", "Clle 81	# 13-05	Local 5-6 C.C. Atlantis"),
                new Almacen(4, "CAFAM   FLORESTA", "7444153", "Trv 48 F #	96-50 Local-101-103	C.C. Floresta"),
                new Almacen(5, "PALATINO", "7444155", "Cra 7 #	139-07 Local - 221 C.C.	Palatino"),
                new Almacen(6, "PUENTE  AEREO", "7444157", "Puente Aereo Local 5 y 6"),
                new Almacen(7, "RETIRO  SPA", "3764061", "Calle 82	#12-07 L-121/122/123 C.C. El Retiro"),
                new Almacen(8, "SALITRE PLAZA", "7444159", "CRA 68 B Nº 24 - 39 Local 143 C.C. Salitre Plaza"),
                new Almacen(9, "SANTA ANA", "7444161", "Calle   110 No. 9B - 04 L-105/106 C.C. Santa Ana"),
                new Almacen(10, "GRAN ESTACION", "7444163", "Av. Clle 26  No.62-49 Local 112 CC  Gran Estación"),
                new Almacen(11, "SANTAFE SPA  (1)", "7444165", "Cll. 183 #	45-03 L-1-35 C.C. Santafe"),
                new Almacen(13, "UNICENTRO", "7444144", "Av. 15 #	123-30 Local 1-252 C.C.	Unicentro"),
                new Almacen(14, "CENTRO CHIA", "8844050", "Av. Pradilla 900 este Chia Local 12 - 17 CC Centro Chia"),
                new Almacen(15, "TITAN", "7433259", "Cra 72  No  80  -94 Local 2-58 puerta 2 calle 80 C.C Titán   Plaza"),
                new Almacen(16, "PLAZA DE LAS AMERICAS", "7444169", "Cra 71 D  No. 6–94 Sur  Local 10-41 C.C. Plaza de las Américas")
            };
        }
    }
}
