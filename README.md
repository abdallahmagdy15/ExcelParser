# Excel Parser Library
The Excel Parser library is a .NET library that provides functionality to import Excel data into domain models and export data of domain models into excel file. It allows you to easily parse Excel files and map the data to your domain model objects and parse data of domain models into excel file to be downloaded.

## Features
- Import Excel data into domain models
- Support for parent-child relationships
- Automatic data type conversion for common data types (string, char, int, decimal)
- Error handling and reporting for data conversion issues

## Installation
You can install the Excel Parser library using NuGet. Run the following command in the NuGet Package Manager Console:
```
Install-Package ExcelParser
```
## Usage
To use the Excel Parser library in your project, follow these steps:

1. Install the Excel Parser library using NuGet (see Installation section).
2. Add a reference to the Excel Parser library in your project.
3. Import the ExcelParser.Models and OfficeOpenXml namespaces in your code.
4. Create an instance of the ExcelParser class.
5. Call the ImportExcelToDomainModelList method, providing the path to the Excel file and an output parameter for the updated file path.
6. Check the returned DomainModelResultList<T> object for errors and imported domain models.

Example:

```
using ExcelParser.Models;
using OfficeOpenXml;
using System;

namespace YourNamespace
{
    class Program
    {
        static void Main(string[] args)
        {
            string uploadedFilePath = "path/to/your/excel/file.xlsx";
            string updatedFilePath;

            ExcelParser excelParser = new ExcelParser();
            DomainModelResultList<YourModel> result = excelParser.ImportExcelToDomainModelList<YourModel>(uploadedFilePath, out updatedFilePath);

            if (result != null)
            {
                // Import successful, handle the imported domain models
                foreach (DomainModelResult<YourModel> domainModelResult in result.ResultList)
                {
                    if (domainModelResult.Errors.Count == 0)
                    {
                        YourModel model = domainModelResult.Result;
                        // Process the imported model
                    }
                    else
                    {
                        // Handle errors for the current model
                        foreach (string error in domainModelResult.Errors)
                        {
                            Console.WriteLine(error);
                        }
                    }
                }
            }
            else
            {
                // Import failed, handle the errors
                foreach (string error in result.Errors)
                {
                    Console.WriteLine(error);
                }
            }

            // Use the updatedFilePath for further processing or download
        }
    }

    class YourModel
    {
        // Define your domain model properties
        // Make sure to provide appropriate data annotations or validations if needed
    }
}
```

## License
The Excel Parser library is licensed under the MIT License.

## Contributions
Contributions to the Excel Parser library are welcome! If you find any issues or have suggestions for improvements, please create an issue or submit a pull request on GitHub.

## Support
If you need any assistance or have any questions regarding the Excel Parser library, feel free to contact the library maintainer or create an issue on GitHub.

