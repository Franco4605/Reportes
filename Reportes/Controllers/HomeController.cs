using Microsoft.AspNetCore.Mvc;
using Reportes.Models;
using System.Diagnostics;

using System.Data;
using System.Data.SqlClient;

using ClosedXML.Excel;

namespace Reportes.Controllers
{
    public class HomeController : Controller
    {
        private readonly string  cadenaSQL;

        public HomeController(IConfiguration config)
        {
            cadenaSQL = config.GetConnectionString("cadenasql");
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Exportar_excel(string fechainicio, string fechafin)
        {
            DataTable tabla_cliente= new DataTable();
            using (var conexion = new SqlConnection(cadenaSQL))
            {
                conexion.Open();
                using (var adapter= new SqlDataAdapter())
                {
                    adapter.SelectCommand = new SqlCommand("sp_reporte_cliente", conexion);
                    adapter.SelectCommand.CommandType= CommandType.StoredProcedure;
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaInicio", fechainicio);
                    adapter.SelectCommand.Parameters.AddWithValue("@FechaFin",fechafin);

                    adapter.Fill(tabla_cliente);
                }
            }
            using (var libro= new XLWorkbook())
            {
                tabla_cliente.TableName = "Clientes";
                var hoja = libro.Worksheets.Add(tabla_cliente);
                hoja.ColumnsUsed().AdjustToContents();

                using (var memoria= new MemoryStream())
                {
                    libro.SaveAs(memoria);

                    var nombreExcel = string.Concat("Reporte cliente", DateTime.Now.ToString(), ".xlsx");
                    return File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                }
            }


          
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}