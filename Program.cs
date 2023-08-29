using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace TurnoApp
{
    class Program
    {
        static readonly string CLIENT_ID = "1089617708991-6guk84qua1u2gv8v1poohj60najt97hk.apps.googleusercontent.com";
        static readonly string CLIENT_SECRET = "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";
        static readonly string REDIRECT_URI = "http://localhost";
        static readonly string[] SCOPES = { SheetsService.Scope.Spreadsheets, SheetsService.Scope.DriveFile };
        static SheetsService sheetsService;

        static UserCredential GetCredentials()
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            return credential;
        }

        static SheetsService CreateSheetsService()
        {
            UserCredential credential = GetCredentials();
            return new SheetsService(new Google.Apis.Services.BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "TurnoApp"
            });
        }

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<Registro> registros = CargarRegistrosDesdeExcel("Registros.xlsx", out int numeroTurno);
            sheetsService = CreateSheetsService();
            string spreadsheetId = "1wqVsSCL0ccH3aGNVd1XWhVOFIUwDgfqLS-toJVWBZ4o";
            string range = "Registros!A:G";

            numeroTurno = CargarNumeroTurno("numeroTurno.txt"); // Cargar el valor del contador desde el archivo

            bool mostrarMenu = true; // Variable para controlar si se muestra el menú principal



            while (true)
            {

                if (mostrarMenu)
                {
                    string opcion = MostrarMenu();

                    switch (opcion)
                    {
                        case "1":
                            IngresarNuevoRegistro(registros, spreadsheetId, range, ref numeroTurno);
                            mostrarMenu = false; // Ocultar el menú después de ingresar un nuevo registro

                            break;


                        case "2":
                            VerRegistros(registros);
                            Console.WriteLine("Presiona Enter para continuar...");
                            Console.ReadLine();
                            mostrarMenu = false; // Ocultar el menú después de ver registros


                            break;

                        case "3":
                            return;


                        case "4":
                            GenerarReporte(registros, "Registros.xlsx");
                            break;





                        default:
                            Console.WriteLine("Opción inválida. Intente nuevamente.");
                            break;
                    }
                }
                else
                {
                    mostrarMenu = true; // Restablecer el valor para mostrar el menú en el siguiente ciclo
                }
                }
            }

        
        static string MostrarMenu()
        {
            Console.Clear();
            Console.Clear(); // Limpiar la pantalla antes de mostrar el formulario de ingreso


            Console.WriteLine(new string('=', 25));
            Console.WriteLine("\n=== Registro de Unidades  ===\n");
            
            Console.WriteLine("   1. Ingresar nuevo registro");
            Console.WriteLine("   2. Ver registros");
            Console.WriteLine("   3. Salir");
            Console.WriteLine(new string('=', 25));
            Console.WriteLine("\n");
            Console.Write("Selecciona una opción: ");
            return Console.ReadLine();

        }

        static List<Registro> CargarRegistrosDesdeExcel(string excelFilePath, out int numeroTurno)
        {
            List<Registro> registros = new List<Registro>();
            numeroTurno = 1; // Valor predeterminado si no se encuentra en el archivo Excel

            if (File.Exists(excelFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Registros");

                    if (worksheet != null)
                    {
                        numeroTurno = Convert.ToInt32(worksheet.Cells["H1"].Value ?? 1);

                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            string noLicencia = worksheet.Cells[row, 1].Value?.ToString();
                            string nombreApellido = worksheet.Cells[row, 2].Value?.ToString();
                            string placas = worksheet.Cells[row, 3].Value?.ToString();
                            string empresa = worksheet.Cells[row, 4].Value?.ToString();

                            DateTime fechaHoraIngreso;
                            if (DateTime.TryParse(worksheet.Cells[row, 5].Value?.ToString(), out fechaHoraIngreso))
                            {
                                string turno = worksheet.Cells[row, 6].Value?.ToString();

                                registros.Add(new Registro(noLicencia, nombreApellido, placas, empresa, fechaHoraIngreso, turno));
                            }
                        }
                    }
                }
            }

            return registros;
        }




        static void IngresarNuevoRegistro(List<Registro> registros, string spreadsheetId, string range, ref int numeroTurno)
        {
            Console.Clear(); // Limpiar la pantalla antes de mostrar el formulario de ingreso
            Console.WriteLine("\n=== Formulario de registro ===\n");
            Console.WriteLine(new string('=', 45));
            Console.WriteLine("Por favor ingresa los siguientes datos:");
            Console.WriteLine(new string('=', 45));

            Console.WriteLine("\n");

           Console.Write("No. Licencia: ");
            string noLicencia = Console.ReadLine();

            Console.Write("Nombre y Apellido: ");
            string nombreApellido = Console.ReadLine();

            Console.Write("Placas de la Unidad: ");
            string placas = Console.ReadLine();

            Console.Write("Empresa: ");
            string empresa = Console.ReadLine();

            DateTime fechaHoraIngreso = DateTime.Now;

            Registro nuevoRegistro = new Registro(noLicencia, nombreApellido, placas, empresa, fechaHoraIngreso, numeroTurno.ToString());

            if (!ExisteRegistro(registros, nuevoRegistro))
            {
                registros.Add(nuevoRegistro);
                // Incrementar el contador y guardarlo en el archivo
                numeroTurno++;
                GuardarNumeroTurno("numeroTurno.txt", numeroTurno);

                AgregarRegistroAGoogleSheets(nuevoRegistro, spreadsheetId, range);
                AgregarRegistroAExcel(nuevoRegistro, "Registros.xlsx");

                Console.Clear(); // Limpiar la pantalla antes de mostrar el formulario de ingreso

                Console.WriteLine("\n=== Datos Ingresados ===");
                Console.WriteLine("-------------------------:");
                Console.WriteLine($"{"No. Licencia:",-25} {nuevoRegistro.NoLicencia}");
                Console.WriteLine($"{"Nombre y Apellido:",-25} {nuevoRegistro.NombreApellido}");
                Console.WriteLine($"{"Placas de la Unidad:",-25} {nuevoRegistro.Placas}");
                Console.WriteLine($"{"Empresa:",-25} {nuevoRegistro.Empresa}");
                Console.WriteLine($"{"Fecha y Hora de Ingreso:",-25} {nuevoRegistro.FechaHoraIngreso}");
                Console.WriteLine($"{"No. Turno:",-25} {nuevoRegistro.Turno}");

               



                Console.WriteLine("\nRegistro agregado exitosamente.\n");
               
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("\nEl registro ya existe. No se agregó duplicado.");
            }
        }

        static int CargarNumeroTurno(string filePath)
        {
            if (File.Exists(filePath))
            {
                string content = File.ReadAllText(filePath);
                if (int.TryParse(content, out int numeroTurno))
                {
                    return numeroTurno;
                }
            }
            return 1; // Valor predeterminado si no se encuentra el archivo o no se puede leer el valor
        }


        static void GuardarNumeroTurno(string filePath, int numeroTurno)
        {
            File.WriteAllText(filePath, numeroTurno.ToString());
        }




        static void VerRegistros(List<Registro> registros)
        {
            Console.Clear(); // Limpiar la pantalla antes de mostrar el formulario de ingreso

            Console.WriteLine("\n=== Registros ===\n");

            Console.WriteLine("\n");
            Console.WriteLine(new string('=', 120));
            Console.WriteLine($"| {"No. Licencia",-12} | {"Nombre y Apellido",-20} | {"Placas de la Unidad",-15} | {"Empresa",-12} | {"Fecha y Hora de Ingreso",-25} | {"No. Turno",-12} |");
            Console.WriteLine(new string('-', 120));

            foreach (var registro in registros)
            {
                Console.WriteLine($"| {registro.NoLicencia,-12} | {registro.NombreApellido,-20} | {registro.Placas,-15} | {registro.Empresa,-12} | {registro.FechaHoraIngreso.ToString("dd/MM/yyyy hh:mm tt"),-25} | {FormatearNumeroIngreso(registro.Turno),-12} |");
            }

            Console.WriteLine(new string('=', 120)); // Mostrar línea horizontal al final

            // Console.WriteLine("Presiona Enter para continuar...");
            // Console.ReadLine();
        }
        static void GenerarReporte(List<Registro> registros, string excelFilePath)
        {
            Console.WriteLine("Generando reporte PDF...");

            string titulo = "Reporte de ingresos a Suministros y Alimentos";
            string fechaHoraGenerado = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string nombreArchivo = "ReporteIngresos.pdf";

            Document doc = new Document(PageSize.LETTER.Rotate());
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(nombreArchivo, FileMode.Create));
            doc.Open();

            doc.AddTitle(titulo);
            doc.AddCreationDate();

            PdfPTable tabla = new PdfPTable(6);
            tabla.WidthPercentage = 100;

            PdfPCell celdaTitulo = new PdfPCell(new Phrase(titulo));
            celdaTitulo.Colspan = 6;
            celdaTitulo.HorizontalAlignment = Element.ALIGN_CENTER;
            celdaTitulo.PaddingBottom = 10;
            tabla.AddCell(celdaTitulo);

            PdfPCell celdaFechaHora = new PdfPCell(new Phrase("Fecha y Hora de Generado: " + fechaHoraGenerado));
            celdaFechaHora.Colspan = 6;
            celdaFechaHora.HorizontalAlignment = Element.ALIGN_CENTER;
            celdaFechaHora.PaddingBottom = 10;
            tabla.AddCell(celdaFechaHora);

            tabla.AddCell("No. Licencia");
            tabla.AddCell("Nombre y Apellido");
            tabla.AddCell("Placas de la Unidad");
            tabla.AddCell("Empresa");
            tabla.AddCell("Fecha y Hora de Ingreso");
            tabla.AddCell("No. Turno");

            foreach (var registro in registros)
            {
                tabla.AddCell(registro.NoLicencia);
                tabla.AddCell(registro.NombreApellido);
                tabla.AddCell(registro.Placas);
                tabla.AddCell(registro.Empresa);
                tabla.AddCell(registro.FechaHoraIngreso.ToString("dd/MM/yyyy hh:mm tt"));
                tabla.AddCell(FormatearNumeroIngreso(registro.Turno));
            }

            doc.Add(tabla);
            doc.Close();

            Console.WriteLine("Reporte PDF generado exitosamente.");
            
            Console.ReadLine();
        }

        static string FormatearNumeroIngreso(string numeroTurno)
        {
            return "Turno: " + numeroTurno.PadLeft(6, '0');
        }

        static bool ExisteRegistro(List<Registro> registros, Registro nuevoRegistro)
        {
            return registros.Any(registro => registro.NoLicencia == nuevoRegistro.NoLicencia && registro.Placas == nuevoRegistro.Placas && registro.Empresa == nuevoRegistro.Empresa);
        }

        static void AgregarRegistroAGoogleSheets(Registro registro, string spreadsheetId, string range)
        {
            List<IList<object>> registrosParaSheet = new List<IList<object>>();

            registrosParaSheet.Add(new List<object>
    {
        registro.NoLicencia,
        registro.NombreApellido,
        registro.Placas,
        registro.Empresa,
        registro.FechaHoraIngreso.ToString("yyyy-MM-dd HH:mm:ss"),
        registro.Turno
    });

            ValueRange body = new ValueRange
            {
                Values = registrosParaSheet
            };

            SpreadsheetsResource.ValuesResource.AppendRequest request =
                sheetsService.Spreadsheets.Values.Append(body, spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            AppendValuesResponse response = request.Execute();
        }


        static void AgregarRegistroAExcel(Registro registro, string excelFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Registros");

                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Registros");
                    worksheet.Cells["A1"].LoadFromCollection(new List<Registro> { registro }, true);
                }
                else
                {
                    int startRow = worksheet.Dimension.End.Row + 1;

                    if (!ExisteRegistroEnExcel(registro, worksheet))
                    {
                        worksheet.Cells[startRow, 1].LoadFromCollection(new List<Registro> { registro }, false);
                    }
                }

                package.Save();
            }
        }

        static bool ExisteRegistroEnExcel(Registro registro, ExcelWorksheet worksheet)
        {
            return worksheet.Cells["A:A"].Any(cell => cell.Value?.ToString() == registro.NoLicencia);
        }
    }
}


class Registro
        {
            public string NoLicencia { get; set; }
            public string NombreApellido { get; set; }
            public string Placas { get; set; }
            public string Empresa { get; set; }
            public DateTime FechaHoraIngreso { get; set; }
            public string Turno { get; set; }

            public Registro(string noLicencia, string nombreApellido, string placas, string empresa, DateTime fechaHoraIngreso, string turno)
            {
                NoLicencia = noLicencia;
                NombreApellido = nombreApellido;
                Placas = placas;
                Empresa = empresa;
                FechaHoraIngreso = fechaHoraIngreso;
                Turno = turno;
            }
        }
    

