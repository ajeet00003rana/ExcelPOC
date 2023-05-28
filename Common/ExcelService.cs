using System.Data;

namespace ExcelPOC.Common
{
    public class ExcelService
    {

        public static async Task<byte[]> CreateExcel(DataSet ds, string filePath, Func<DataSet, string, Task> excelFunction)
        {
            DataSet data = ds;
            try
            {
                string excelFilename = Path.Combine(filePath);
                File.Create(excelFilename).Close();

                //callback function to create excel
                await excelFunction(data, excelFilename);

                if (excelFilename == null)
                    throw new Exception("filename not present");

                var memory = new MemoryStream();
                try
                {
                    using (var stream = new FileStream(excelFilename, FileMode.Open))
                    {
                        await stream.CopyToAsync(memory);
                    }
                    memory.Position = 0;
                }
                catch (Exception ex)
                {
                    throw;
                }

                return File.ReadAllBytes(excelFilename);
            }
            catch
            {
                throw;
            }
        }
    }
}

