using ExcelDataReader;
using ExcelPOC.Models;
using System.Data;
using System.Text;

namespace ExcelPOC.Common
{

    public interface IUtilityWrapper
    {
        string ExtractDataValueFromExcel(Dictionary<int, string> excelColumnDictionary, IDataRecord excelReader, string column);

        Task<TUploadBatchViewModel> BasicUploader<TUploadBatchViewModel>(
            byte[] batchFileData, string batchFileName,
            Func<TUploadBatchViewModel, Dictionary<int, string>, IExcelDataReader, HashSet<string>, TUploadBatchViewModel> transformFunc,
            Func<TUploadBatchViewModel, Task<TUploadBatchViewModel>> addFunc,
            params string[] columns)
            where TUploadBatchViewModel : IUploadBatchViewModel, new();
    }

    public class UtilityWrapper : IUtilityWrapper
    {
        public UtilityWrapper()
        {

        }

        public async Task<T> BasicUploader<T>(
           byte[] batchFileData,
           string batchFileName,
           Func<T, Dictionary<int, string>, IExcelDataReader, HashSet<string>, T> transformFunc,
           Func<T, Task<T>> addFunc,
           params string[] columns
       )
           where T : IUploadBatchViewModel, new()
        {
            if (batchFileData == null) return default;

            var fileExtension = Path.GetExtension(batchFileName);

            if (fileExtension == null ||
                (!fileExtension.Equals(".xls", StringComparison.OrdinalIgnoreCase) &&
                 !fileExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"Incorrect file type, please upload a Microsoft Excel file: {batchFileName}");
            }

            var batchViewModel = new T();

            try
            {
                await using (var stream = new MemoryStream(batchFileData))
                {
                    using var excelReader = CreateExcelReader(stream);
                    //4. DataSet - Create column names from first row
                    excelReader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        // Gets or sets a callback to obtain configuration options for a DataTable. 
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            // Gets or sets a value indicating whether to use a row from the 
                            // data as column names.
                            UseHeaderRow = true,

                            // Gets or sets a callback to determine which row is the header row. 
                            // Only called when UseHeaderRow = true.
                            ReadHeaderRow = (rowReader) =>
                            {
                                // F.ex skip the first row and use the 2nd row as column headers:
                                rowReader.Read();
                            }
                        }
                    });
                    var isFirstRow = true;

                    //5. Data Reader methods
                    var excelColumnDictionary = new Dictionary<int, string>();

                    batchViewModel.IsBatchValid = true;
                    batchViewModel.ValidationMessageList = new List<string>();
                    var nullRows = 0;
                    var hash = new HashSet<string>();

                    const int maxrows = 500;
                    var index = 0;

                    while (excelReader.Read() && batchViewModel.IsBatchValid)
                    {
                        if (index >= maxrows)
                        {
                            if (batchViewModel.IsBatchValid)
                            {
                                batchViewModel = await addFunc(batchViewModel);
                            }

                            batchViewModel.ResetContent();

                            index = 0;
                        }

                        if (isFirstRow)
                        {
                            batchViewModel.ExcelRowNumber++;
                            excelColumnDictionary = SetupColumnDictionary(excelColumnDictionary, excelReader);
                            isFirstRow = false;
                        }
                        else if (
                            columns.All(x =>
                                    string.IsNullOrEmpty(ExtractDataValueFromExcel(excelColumnDictionary,
                                        excelReader, x))))
                        {
                            ++nullRows;
                            if (nullRows >= 5)
                            {
                                break;
                            }
                        }
                        else
                        {
                            batchViewModel.ExcelRowNumber++;
                            batchViewModel = transformFunc(batchViewModel, excelColumnDictionary, excelReader, hash);
                            index++;
                        }
                    }
                }

                if (batchViewModel.IsBatchValid)
                {
                    return await addFunc(batchViewModel);
                }

                throw new Exception("Please upload a revised batch upload file, and run the process again.");
            }
            catch (Exception e)
            {
                var errorMessage = $"Batch Upload Error, the error has been logged. Row:{batchViewModel.ExcelRowNumber}";
                batchViewModel.IsBatchValid = false;
                throw;
            }
        }

        private static IExcelDataReader CreateExcelReader(Stream inputStream)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return ExcelReaderFactory.CreateReader(inputStream);
        }


        private static Dictionary<int, string> SetupColumnDictionary(Dictionary<int, string> excelColumnDictionary,
           IDataRecord excelReader)
        {
            for (var i = 0; i < excelReader.FieldCount; i++)
            {
                excelColumnDictionary.Add(i, excelReader.GetString(i));
            }

            return excelColumnDictionary;
        }

        public string ExtractDataValueFromExcel(Dictionary<int, string> excelColumnDictionary,
           IDataRecord excelReader, string column)
        {
            try
            {
                var val = excelReader.GetValue(excelColumnDictionary.First(p => p.Value == column).Key);
                return val != null ? val.ToString().TrimNullCheck() : string.Empty;
            }
            catch (Exception e)
            {
                throw new Exception(
                    $"Missing column: '{column}'. Please ensure that the column header is spelled exactly.", e);
            }
        }
    }
}
