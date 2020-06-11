namespace ExportadorExcel.Modelos
{
    public class Almacen
    {
        public int Id { get; }
        public string Nombre { get; }
        public string Telefono { get; }
        public string Direccion { get; }

        public Almacen(int id, string nombre, string telefono, string direccion)
        {
            Id = id;
            Nombre = nombre;
            Telefono = telefono;
            Direccion = direccion;
        }
    }
}
