namespace ExcelPOC.Models
{
    public interface IUploadBatchViewModel
    {
        bool IsBatchValid { get; set; }
        List<string> ValidationMessageList { get; set; }
        int ExcelRowNumber { get; set; }
        void ResetContent();
    }

    [Serializable]
    public class UploadBatchViewModel<T> : IUploadBatchViewModel
    {
        public List<T> Items { get; set; }
        public UploadBatchViewModel()
        {
            Items = new List<T>();
            ValidationMessageList = new List<string>();
        }

        public bool IsBatchValid { get; set; }
        public List<string> ValidationMessageList { get; set; }
        public int ExcelRowNumber { get; set; }

        public void ResetContent()
        {
            Items.Clear();
        }
    }
}
