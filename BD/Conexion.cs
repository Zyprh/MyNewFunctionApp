using System;
using System.Data.SqlClient;

public class Conexion : IDisposable
{
    private SqlConnection _conexion;

    public Conexion(string connectionString)
    {
        _conexion = new SqlConnection(connectionString);
    }

    public void AbrirConexion()
    {
        if (_conexion.State == System.Data.ConnectionState.Closed)
        {
            _conexion.Open();
        }
    }

    public SqlConnection ObtenerConexion()
    {
        return _conexion;
    }

    public void Dispose()
    {
        if (_conexion != null && _conexion.State == System.Data.ConnectionState.Open)
        {
            _conexion.Close();
        }
        _conexion?.Dispose();
    }
}

