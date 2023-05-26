using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using Spire.Xls;
using System.Net;
// guardar excel


namespace RegistroDeUsuarioRenab
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; 

            // URL del endpoint
            //string url = string.Empty;

            // Crear cliente HTTP
           // var client = new HttpClient();

            // Ruta del archivo de Excel
            var filePath = @"C:\RegistroDeUsuarioRenab\planilla.xlsx";
            var filePathSave = @"C:\RegistroDeUsuarioRenab\resultados.xlsx";

            // Lista para almacenar los datos del Excel
            var cedulas = new List<registro>();

            using (var packagefile = new ExcelPackage(new FileInfo(filePath))){

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("El archivo Excel no existe en la ruta especificada.");
                    Console.WriteLine("_____________________");
                    Console.WriteLine("Presiona cualquier tecla para salir.");
                    Console.ReadKey();
                    return;
                }

                // leer los datos del Excel utilizando EPPlus
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Obtener la primera hoja del libro
                    var startRow = worksheet.Dimension.Start.Row; // Obtener el número de la primera fila
                    var endRow = worksheet.Dimension.End.Row; // Obtener el número de la última fila

                    // Recorrer cada fila y obtener el valor de la celda "A, B" (primera columna)
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var cedula = worksheet.Cells[row, 1].Value?.ToString().Trim();
                        var fechaVencimiento = worksheet.Cells[row, 2].Value?.ToString().Trim();

                        if (!string.IsNullOrEmpty(cedula))
                        {
                            // Validar si el valor de la celda cumple con el patrón de la cédula
                            if (Regex.IsMatch(cedula, @"^(PE|E|N|[23456789](?:AV|PI)?|1[0123]?(?:AV|PI)?)-(\d{1,5})-(\d{1,6})$"))
                            {
                                cedulas.Add(new registro { cedula = cedula, fechaVencimiento = fechaVencimiento}); // Agregar la cédula a la lista
                            }
                        }
                    }
                }
                //---------------------------------------------------------------------------------------------------------------
                //crear xlsx de almacenado (2° archivo xlsx)
                ExcelPackage package2 = new ExcelPackage(new FileInfo(filePathSave));
                // Seleccionar o crear hoja de trabajo
                ExcelWorksheet worksheet2 = package2.Workbook.Worksheets["Encontrados"] ?? package2.Workbook.Worksheets.Add("Encontrados");
                // Agregar una fila en blanco al inicio de la hoja de trabajo
                worksheet2.Cells["A1"].Value = "Inicializar xlsx";

                // get the exact day time
                DateTime now = DateTime.Now;
                
                // Definir valores de la fila
                object[] rowData = new object[] { "Cédula", "Fecha Vencimiento", "Creado?", "Fecha de Consulta:" };
                // Agregar la fila al final de la hoja de trabajo
                int rowNumber = worksheet2.Dimension.End.Row + 1;
                //inicializar los campos
                worksheet2.Cells[rowNumber, 1].LoadFromArrays(new object[][] { rowData });
                // Guardar el archivo xlsx
                package2.Save();


                // Iteracion por cada cedula verificada
                int iteracion = 0;

 
// 116
                    foreach (var cedulaIngresada in cedulas){
                        DateTime nowFor = DateTime.Now;
                        // Obtener los valores de provincia, tomo y asiento de la cédula
                        var regex = new Regex(@"^(\w+)-(\d+)-(\d+)$");
                        var match = regex.Match(cedulaIngresada.cedula);

                        iteracion++;
                        
                        Console.WriteLine($"Cédula {iteracion}: {cedulaIngresada.cedula}: {cedulaIngresada.fechaVencimiento} ");
                        rowNumber = worksheet2.Dimension.End.Row + 1;
                        worksheet2.Cells[rowNumber, 1].Value = cedulaIngresada.cedula;
                        worksheet2.Cells[rowNumber, 2].Value = cedulaIngresada.fechaVencimiento;
                        worksheet2.Cells[rowNumber, 3].Value = "SI";
                        worksheet2.Cells[rowNumber, 4].Value = nowFor.ToString("dd-MM-yyyy HH:mm:ss");
                        package2.Save();
                }

            }

                Console.WriteLine("Presiona cualquier tecla para salir.");
                Console.ReadKey();
            }

            
        }
    }

