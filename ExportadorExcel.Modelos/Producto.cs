using System;

namespace ExportadorExcel.Modelos
{
    public class Producto
    {
        public int Id { get; }
        public string Referencia { get; }
        public string Descripcion { get; }
        public DateTime FechaCreacion { get; }
        public decimal Valor { get; }

        public Producto(int id, string referencia, string descripcion, DateTime fechaCreacion, decimal valor)
        {
            Id = id;
            Referencia = referencia;
            Descripcion = descripcion;
            FechaCreacion = fechaCreacion;
            Valor = valor;
        }
    }
}
