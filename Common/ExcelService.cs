using ExcelDataReader;
using ExcelPOC.Models;
using System.Data;

namespace ExcelPOC.Common
{
    public interface IExcelService
    {
        Task<byte[]> CreateExcel(DataSet ds, string filePath, Func<DataSet, string, Task> excelFunction);
        Task<UploadBatchViewModel<UploadedExcelColumnClassForDB>> ExcelUploader(byte[] partsBatchFileData, string partsBatchFileName);
    }

    public class ExcelService : IExcelService
    {
        private readonly IUtilityWrapper _utilityWrapper;
        public ExcelService(IUtilityWrapper utilityWrapper)
        {
            _utilityWrapper = utilityWrapper;
        }

        public async Task<byte[]> CreateExcel(DataSet ds, string filePath, Func<DataSet, string, Task> excelFunction)
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

        public async Task<UploadBatchViewModel<UploadedExcelColumnClassForDB>> ExcelUploader(byte[] partsBatchFileData, string partsBatchFileName)
        {
            var partBatchViewModel = new UploadBatchViewModel<UploadedExcelColumnClassForDB>();

            if (partsBatchFileData != null)
            {
                try
                {
                    partBatchViewModel = await _utilityWrapper.BasicUploader<UploadBatchViewModel<UploadedExcelColumnClassForDB>>(
                            partsBatchFileData,
                            partsBatchFileName,
                            TransformExcelFileRowToModel,
                            AddItemListToDB,
                            ColumnNameEnum.ID,
                            ColumnNameEnum.Name
                        );

                    if (partBatchViewModel.Items.Count > 0)
                    {
                        partBatchViewModel.ValidationMessageList.Add(
                            $"Success: Part Uploader Completed Sucessfully.");
                    }
                }
                catch (Exception e)
                {
                    partBatchViewModel.ValidationMessageList.Add($"Part Batch Error: {e.Message}");
                }
            }
            else
            {
                partBatchViewModel.ValidationMessageList.Add("Failed To Upload Parts List File");
            }

            return partBatchViewModel;
        }

        private UploadBatchViewModel<UploadedExcelColumnClassForDB> TransformExcelFileRowToModel(UploadBatchViewModel<UploadedExcelColumnClassForDB> viewModel,
            Dictionary<int, string> excelColumnDictionary, IExcelDataReader excelReader, HashSet<string> seen)
        {
            var isValidAccount = GenerateDataFromExcel(viewModel, excelColumnDictionary, excelReader);

            if (!isValidAccount)
            {
                viewModel.IsBatchValid = false;
            }

            return viewModel;
        }

        private bool GenerateDataFromExcel(UploadBatchViewModel<UploadedExcelColumnClassForDB> viewModel,
        Dictionary<int, string> excelColumnDictionary, IDataRecord excelReader)
        {
            var partCode = string.Empty;

            try
            {
                var id = _utilityWrapper.ExtractDataValueFromExcel(excelColumnDictionary, excelReader, ColumnNameEnum.ID);
                var name = _utilityWrapper.ExtractDataValueFromExcel(excelColumnDictionary, excelReader, ColumnNameEnum.Name);

                if (string.IsNullOrEmpty(id))
                {
                    AddValidationMessage(viewModel, $"Error: Vendor Part Description is missing. Please update the file and try again.");
                    return false;
                }

                if (string.IsNullOrEmpty(name))
                {
                    AddValidationMessage(viewModel, $"Error: Vendor Part Code is missing. Please update the file and try again.");
                    return false;
                }

                var vendorPart = new UploadedExcelColumnClassForDB
                {
                    Id = Convert.ToInt32(id),
                    Name = name,
                };

                viewModel.Items.Add(vendorPart);

                return true;
            }
            catch (Exception e)
            {
                AddValidationMessage(viewModel, $"Error: PartCode {partCode} failed to be added, please try again.");
                return false;
            }
        }

        private static void AddValidationMessage(UploadBatchViewModel<UploadedExcelColumnClassForDB> viewModel, string message)
        {
            viewModel.ValidationMessageList.Add(message);
        }

        private Task<UploadBatchViewModel<UploadedExcelColumnClassForDB>> AddItemListToDB(UploadBatchViewModel<UploadedExcelColumnClassForDB> viewModel)
        {
            try
            {
                int itemCreated = 0;
                int itemUpdated = 0;

                foreach (var vendorPart in viewModel.Items)
                {
                    try
                    {
                        if (vendorPart.Id > 0)
                        {
                            //FindAndUpdate(vendorPart, false);
                            itemUpdated++;
                        }
                        else
                        {
                            //Insert(vendorPart, false);
                            itemCreated++;
                        }
                    }
                    catch (Exception e)
                    {
                        throw;
                    }
                }

                if (itemCreated == 0 && itemUpdated == 0)
                {
                    viewModel.IsBatchValid = false;
                    viewModel.ValidationMessageList.Add($"Error: No record found in excel file.");
                    return Task.FromResult(viewModel);
                }

                //_repository.SaveChanges();

                viewModel.ValidationMessageList.Add(itemCreated > 0
                    ? $"Success: {itemCreated} Vendor Parts added successfully."
                    : $"Success: 0 Vendor Parts added.");

                if (itemUpdated > 0)
                    viewModel.ValidationMessageList.Add($"Success: Vendor Part {itemUpdated} updated successfully.");

            }
            catch (Exception ex)
            {
                viewModel.ValidationMessageList
                    .Add($"Error: No Vendor Parts were added, please try again. If the problem continues please contact support.");
                throw;
            }

            return Task.FromResult(viewModel);
        }
    }
}

