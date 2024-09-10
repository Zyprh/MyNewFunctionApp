using Azure.Storage.Blobs;
using Azure.Security.KeyVault.Secrets;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ClosedXML.Excel;
using Newtonsoft.Json;
using Azure.Identity;

namespace FunctionApp1
{
    public static class ProductFunctions
    {

        [FunctionName("ReporteVenta")]
        public static async Task<IActionResult> ReporteVenta(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            int año = data?.año ?? DateTime.Now.Year;
            int mes = data?.mes ?? DateTime.Now.Month;

            string fechainicio = new DateTime(año, mes, 1).ToString("yyyy-MM-dd");
            string fechafinal = new DateTime(año, mes, DateTime.DaysInMonth(año, mes)).ToString("yyyy-MM-dd");

            string query = $"SELECT ProductID, Nombre, Descripción, Categoría, Precio, CantidadEnStock, FechaDeIngreso, FechaDeVenta, Proveedor, UbicaciónEnAlmacén FROM dbo.Productos WHERE FechaDeVenta BETWEEN '{fechainicio}' AND '{fechafinal}'";

            string kvURL = Environment.GetEnvironmentVariable("KeyVaultURL");
            if (string.IsNullOrEmpty(kvURL))
            {
                log.LogError("El URL de 'KeyVault' no está en las variables de entorno.");
                return new BadRequestObjectResult("Falta el URL.");
            }

            var client = new SecretClient(new Uri(kvURL), new DefaultAzureCredential());
            KeyVaultSecret sqlSecret = await client.GetSecretAsync("DbConnectionString");
            string connectionString = sqlSecret.Value;
            KeyVaultSecret blobSecret = await client.GetSecretAsync("ConectarBS");
            string blobConnectionString = blobSecret.Value;

            var blobServiceClient = new BlobServiceClient(blobConnectionString);
            var blobContainerClient = blobServiceClient.GetBlobContainerClient("reporteventas");

            try
            {
                List<Producto> productos = new List<Producto>();

                using (var conexion = new Conexion(connectionString))
                {
                    conexion.AbrirConexion();

                    using (var cmd = new SqlCommand(query, conexion.ObtenerConexion()))
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        while (reader.Read())
                        {
                            productos.Add(new Producto
                            {
                                ProductID = (int)reader["ProductID"],
                                Nombre = (string)reader["Nombre"],
                                Descripción = (string)reader["Descripción"],
                                Categoría = (string)reader["Categoría"],
                                Precio = (decimal)reader["Precio"],
                                CantidadEnStock = (int)reader["CantidadEnStock"],
                                FechaDeIngreso = (DateTime)reader["FechaDeIngreso"],
                                FechaDeVenta = (DateTime)reader["FechaDeVenta"],
                                Proveedor = (string)reader["Proveedor"],
                                UbicaciónEnAlmacén = (string)reader["UbicaciónEnAlmacén"]
                            });
                        }
                    }
                }

                // Generar PDF y Excel
                string pdfPath = GenerarPdf(fechainicio, fechafinal, productos, blobConnectionString);
                string excelPath = GenerarExcel(productos);

                // Subir archivos generados al Blob Storage
                await SubirArchivoABlobStorage(pdfPath, $"Reporte_{mes}_{año}.pdf", blobContainerClient);
                await SubirArchivoABlobStorage(excelPath, $"Reporte_{mes}_{año}.xlsx", blobContainerClient);

                // Asegurarse de eliminar archivos temporales después de subirlos
                File.Delete(pdfPath);
                File.Delete(excelPath);
            }
            catch (Exception ex)
            {
                log.LogError($"Error procesando la solicitud: {ex.Message}");
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }

            return new OkObjectResult("Reporte generado y subido a Blob Storage.");
        }

        public static string GenerarPdf(string fechainicio, string fechafinal, List<Producto> productos, string blobConnectionString)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string pdfFileName = Path.Combine($"ReporteVentas_{fechainicio}_{fechafinal}_{timestamp}.pdf");

            using (var pdfStream = new FileStream(pdfFileName, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var pdfDoc = new Document(PageSize.A4.Rotate()))
            {
                PdfWriter.GetInstance(pdfDoc, pdfStream);
                pdfDoc.Open();

                // Descargar la imagen
                byte[] imageBytes = DownloadImageFromBlobAsync("reporteventas", "logo.png", blobConnectionString).Result;
                using (var imageStream = new MemoryStream(imageBytes))
                {
                    var image = iTextSharp.text.Image.GetInstance(imageStream);
                    image.ScaleToFit(100f, 100f); // Ajusta el tamaño si es necesario
                    image.Alignment = Element.ALIGN_CENTER;
                    pdfDoc.Add(image);
                }

                // Información de la empresa
                Paragraph companyInfo = new Paragraph("Nombre de la empresa: Kellysolution\nCorreo: kelly@gmail.com\nCelular: 926261263", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12));
                companyInfo.Alignment = Element.ALIGN_CENTER;
                companyInfo.SpacingBefore = 20f;
                companyInfo.SpacingAfter = 20f;
                pdfDoc.Add(companyInfo);

                Paragraph title = new Paragraph($"Reporte de ventas del mes {fechainicio} a {fechafinal}", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18));
                title.Alignment = Element.ALIGN_CENTER;
                title.SpacingAfter = 20f;
                pdfDoc.Add(title);

                PdfPTable table = new PdfPTable(10);
                table.WidthPercentage = 100;
                table.SetWidths(new float[] { 8f, 20f, 25f, 15f, 12f, 12f, 15f, 15f, 15f, 20f });

                string[] headers = { "ID", "Nombre", "Descripción", "Categoría", "Precio", "Cantidad", "Fecha Ingreso", "Fecha Venta", "Proveedor", "Ubicación" };
                foreach (string header in headers)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(header, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)));
                    cell.BackgroundColor = BaseColor.GRAY;
                    cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    table.AddCell(cell);
                }

                foreach (Producto producto in productos)
                {
                    table.AddCell(new PdfPCell(new Phrase(producto.ProductID.ToString(), FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.Nombre, FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.Descripción, FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.Categoría, FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.Precio.ToString("C2"), FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.CantidadEnStock.ToString(), FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.FechaDeIngreso.ToString("dd/MM/yyyy"), FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.FechaDeVenta.ToString("dd/MM/yyyy"), FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.Proveedor, FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                    table.AddCell(new PdfPCell(new Phrase(producto.UbicaciónEnAlmacén, FontFactory.GetFont(FontFactory.HELVETICA, 10))));
                }

                pdfDoc.Add(table);
                pdfDoc.Close();
            }

            return pdfFileName;
        }

        public static string GenerarExcel(List<Producto> productos)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string excelFileName = Path.Combine($"ReporteVentas_{timestamp}.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte de Ventas");

                // Encabezado
                var headerCell = worksheet.Cell("A1");
                headerCell.Value = "Reporte de Ventas";
                headerCell.Style.Font.SetBold();
                headerCell.Style.Font.SetFontSize(16);
                headerCell.Style.Fill.BackgroundColor = XLColor.Green;
                headerCell.Style.Font.SetFontColor(XLColor.White);
                headerCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range("A1:J1").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Títulos de columnas
                worksheet.Cell("A2").Value = "ID";
                worksheet.Cell("B2").Value = "Nombre";
                worksheet.Cell("C2").Value = "Descripción";
                worksheet.Cell("D2").Value = "Categoría";
                worksheet.Cell("E2").Value = "Precio";
                worksheet.Cell("F2").Value = "Cantidad";
                worksheet.Cell("G2").Value = "Fecha Ingreso";
                worksheet.Cell("H2").Value = "Fecha Venta";
                worksheet.Cell("I2").Value = "Proveedor";
                worksheet.Cell("J2").Value = "Ubicación";

                // Estilo de los encabezados
                worksheet.Range("A2:J2").Style.Font.SetBold();
                worksheet.Range("A2:J2").Style.Fill.BackgroundColor = XLColor.Gray;
                worksheet.Range("A2:J2").Style.Font.SetFontColor(XLColor.White);
                worksheet.Range("A2:J2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int row = 3;
                foreach (Producto producto in productos)
                {
                    worksheet.Cell($"A{row}").Value = producto.ProductID;
                    worksheet.Cell($"B{row}").Value = producto.Nombre;
                    worksheet.Cell($"C{row}").Value = producto.Descripción;
                    worksheet.Cell($"D{row}").Value = producto.Categoría;
                    worksheet.Cell($"E{row}").Value = producto.Precio;
                    worksheet.Cell($"F{row}").Value = producto.CantidadEnStock;
                    worksheet.Cell($"G{row}").Value = producto.FechaDeIngreso.ToString("dd/MM/yyyy");
                    worksheet.Cell($"H{row}").Value = producto.FechaDeVenta.ToString("dd/MM/yyyy");
                    worksheet.Cell($"I{row}").Value = producto.Proveedor;
                    worksheet.Cell($"J{row}").Value = producto.UbicaciónEnAlmacén;
                    row++;
                }

                // Ajustar ancho de las columnas
                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(excelFileName);
            }

            return excelFileName;
        }

        public static async Task SubirArchivoABlobStorage(string filePath, string blobName, BlobContainerClient containerClient)
        {
            try
            {
                var blobClient = containerClient.GetBlobClient(blobName);
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    await blobClient.UploadAsync(fileStream, overwrite: true);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error al subir el archivo '{blobName}' al Blob Storage: {ex.Message}", ex);
            }
        }
        public static async Task<byte[]> DownloadImageFromBlobAsync(string containerName, string imageName, string connectionString)
        {
            var blobServiceClient = new BlobServiceClient(connectionString);
            var blobContainerClient = blobServiceClient.GetBlobContainerClient(containerName);
            var blobClient = blobContainerClient.GetBlobClient(imageName);

            using (var memoryStream = new MemoryStream())
            {
                await blobClient.DownloadToAsync(memoryStream);
                return memoryStream.ToArray();
            }
        }

    }

}