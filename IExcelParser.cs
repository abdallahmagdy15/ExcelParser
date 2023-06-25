using ExcelParser.Models;

namespace ExcelParser
{
    public interface IExcelParser
    {
        public DomainModelResultList<T>? ImportExcelToDomainModelList<T>(string pathUploadedDownlod, out string updatedFilePath);
        public string? ExportDomainModelListToExcel<T>(List<T> modelList);
    }
}
