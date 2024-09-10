using System;
using System.Collections.Generic;
using System.Data.SqlClient;

public class Producto
{
    public int ProductID { get; set; }  // Identificador único del producto
    public string Nombre { get; set; }  // Nombre del producto
    public string Descripción { get; set; }  // Descripción breve del producto
    public string Categoría { get; set; }  // Categoría del producto
    public decimal Precio { get; set; }  // Precio del producto
    public int CantidadEnStock { get; set; }  // Cantidad disponible en inventario
    public DateTime FechaDeIngreso { get; set; }  // Fecha en que el producto ingresó al inventario
    public DateTime FechaDeVenta { get; set; }  // Fecha de venta del producto
    public string Proveedor { get; set; }  // Nombre del proveedor del producto
    public string UbicaciónEnAlmacén { get; set; }  // Ubicación física del producto en el almacén
    public bool Activo { get; set; }  // Estado del producto (true para activo, false para inactivo)
}

