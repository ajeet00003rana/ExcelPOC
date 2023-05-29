namespace ExcelPOC.Models
{
    public class UploadedExcelColumnClassForDB
    {
        public int Id { get; set; }
        public string? Name { get; set; }
    }

    public class ColumnNameEnum
    {
        public const string ID = "ID";
        public const string Name = "Name";
    }
}
