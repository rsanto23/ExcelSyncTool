using ExcelSyncTool.Services;

var reader = new ExcelReader();
var datos = reader.LeerExcel("ruta/completa/a/tu/archivo.xlsx");

foreach (var fila in datos)
{
    Console.WriteLine(string.Join(" | ", fila.Select(kv => $"{kv.Key}: {kv.Value}")));
}
