using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelSyncTool.Services
{
    public class ExcelReader
    {
        public List<Dictionary<string, string>> LeerExcel(string rutaArchivo)
        {
            var resultado = new List<Dictionary<string, string>>();

            // Habilita el soporte de licencias no comerciales (EPPlus)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(rutaArchivo));
            var hoja = package.Workbook.Worksheets[0]; // Lee la primera hoja

            int filaInicio = hoja.Dimension.Start.Row + 1; // Asume que la primera fila son cabeceras
            int filaFin = hoja.Dimension.End.Row;
            int columnaInicio = hoja.Dimension.Start.Column;
            int columnaFin = hoja.Dimension.End.Column;

            // Leer cabeceras
            var cabeceras = new List<string>();
            for (int col = columnaInicio; col <= columnaFin; col++)
            {
                cabeceras.Add(hoja.Cells[1, col].Text);
            }

            // Leer datos
            for (int fila = filaInicio; fila <= filaFin; fila++)
            {
                var filaDatos = new Dictionary<string, string>();
                for (int col = columnaInicio; col <= columnaFin; col++)
                {
                    string clave = cabeceras[col - 1];
                    string valor = hoja.Cells[fila, col].Text;
                    filaDatos[clave] = valor;
                }
                resultado.Add(filaDatos);
            }

            return resultado;
        }
    }
}
