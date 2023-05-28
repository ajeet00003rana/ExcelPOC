using System.Data;

namespace ExcelPOC.Models
{
    public class SampleData
    {
        public static DataTable GetDataTable()
        {
            DataTable table = new DataTable(); ;

            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add(0, "Lion");
            table.Rows.Add(1, "Tiger");
            table.Rows.Add(2, "Cat");
            table.Rows.Add(3, "Goat");
            table.Rows.Add(4, "Panther");
            table.Rows.Add(5, "Fox");
            table.Rows.Add(6, "Cheetah");

            return table;
        }
    }
}
