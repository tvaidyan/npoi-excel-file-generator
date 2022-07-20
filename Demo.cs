namespace ExcelFileGenerationDemo;
public static class Demo
{
    public static byte[] MakeExcelFile()
    {
        var employees = new List<Employee>();

        employees.Add(new Employee
        {
            EmployeeId = "ABC123",
            FirstName = "Tom",
            LastName = "Vaidyan",
            FavoriteColor = "Blue"
        });

        employees.Add(new Employee
        {
            EmployeeId = "DEF456",
            FirstName = "Jack",
            LastName = "Frost",
            FavoriteColor = "Red"
        });

        employees.Add(new Employee
        {
            EmployeeId = "GHI789",
            FirstName = "Jill",
            LastName = "Taylor",
            FavoriteColor = "Yellow"
        });
        var employeeTable = employees.ToDataTable();

        var excelMaker = new ExcelFileGenerator();
        return excelMaker.Create(employeeTable);
    }
}
