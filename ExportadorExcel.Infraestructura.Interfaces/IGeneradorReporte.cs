using System.Collections.Generic;

namespace ExportadorExcel.Infraestructura.Interfaces
{
    public interface IGeneradorReporte
    {
        byte[] Generar<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos");
        byte[] Generar<T>(byte[] recurso, string hojasDeCalculo, params IEnumerable<T>[] datosHojasDeCalculos);
    }
}