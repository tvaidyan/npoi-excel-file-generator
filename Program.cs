using ExcelFileGenerationDemo;
using static Microsoft.AspNetCore.Http.Results;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.MapGet("/", () => "Hello World!");
app.MapGet("/excel", () =>
{
    var bytes = Demo.MakeExcelFile();
    return File(bytes, "application/vnd.ms-excel", "awesome-excel.xlsx");
});

app.Run();
