//##################### Excel Parser - Domain Model - V 1.0 ########################


using ExcelParser.Models;
using OfficeOpenXml;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ElectronicInvoices.Services
{
    public class ExcelParser
    {
        public DomainModelResultList<T>? ImportExcelToDomainModelList<T>(string uploadedFilePath, out string updatedFilePath)
        {
            //errors = new List<string>();
            var parentDomainModelResultList = new DomainModelResultList<T>();
            updatedFilePath = null;//init

            try
            {
                var parentModelProperties = typeof(T).GetProperties();
                var parentIdProp = parentModelProperties.FirstOrDefault(p => p.Name.ToLowerInvariant().Trim() == "id");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(uploadedFilePath)))
                {
                    bool isTopParentSheet = true;

                    foreach (var sheet in package.Workbook.Worksheets)
                    {

                        //skipp errors sheet
                        if (sheet.Name.ToLowerInvariant().Trim().Contains("error"))
                            continue;

                        // Access and validate the column names in the first row
                        var columnNames = new List<string>();
                        for (int columnIndex = 1; columnIndex <= sheet.Dimension.Columns; columnIndex++)
                        {
                            var columnName = sheet.Cells[1, columnIndex].Value?.ToString();
                            if (string.IsNullOrEmpty(columnName))
                            {
                                parentDomainModelResultList.Errors.Add($"Column name is missing at column index {columnIndex}, sheet {sheet.Name}");
                            }
                            //skip errors column ... 
                            else if (columnName.ToLowerInvariant().Trim().Contains("error"))
                            {
                                continue;
                            }
                            else
                                columnNames.Add(columnName);

                        }
                        //end access and validate columns of sheet


                        if (isTopParentSheet)
                        {
                            //set values for each row
                            for (var rowIndex = 2; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                            {
                                var currentDomainModelResult = new DomainModelResult<T>()
                                {
                                    Result = (T)Activator.CreateInstance(typeof(T))
                                };

                                foreach (var columnName in columnNames)
                                {

                                    var property = parentModelProperties.FirstOrDefault(p => p.Name.ToLowerInvariant().Trim() == columnName.ToLowerInvariant().Trim());
                                    if (property != null)
                                    {
                                        // Set the property value based on the corresponding column value
                                        // You can implement any necessary data type conversions or validation here
                                        var columnIndex = columnNames.IndexOf(columnName) + 1;
                                        var cellValue = sheet.Cells[rowIndex, columnIndex].Value?.ToString();
                                        var errors = new List<string>();

                                        currentDomainModelResult.Result = (T?)SetValueConvertedToDatatype(ref property, currentDomainModelResult.Result, cellValue, rowIndex, typeof(T), ref errors);

                                        currentDomainModelResult.Errors.AddRange(errors);

                                    }
                                    else
                                    {
                                        currentDomainModelResult.Errors.Add($"No property found with the name '{columnName}' in the model");
                                    }
                                }

                                //end of row , add to the list 
                                currentDomainModelResult.RowIndex = rowIndex;
                                parentDomainModelResultList.ResultList.Add(currentDomainModelResult);
                            }

                        }
                        else
                        {

                            //set child current sheet types and props and list of objects
                            PropertyInfo[] currentModelProps;
                            Type currentType = typeof(T);
                            List<object?> currentModelList;
                            var isOneToOneEntity = false;
                            //

                            var nestedModelProp = parentModelProperties.FirstOrDefault(p => p.Name.Trim().ToLowerInvariant() == sheet.Name.Trim().ToLowerInvariant());
                            if (nestedModelProp != null)
                            {
                                //check if the property type is one-to-one entity type not list ..
                                if (nestedModelProp.PropertyType.IsClass && !nestedModelProp.PropertyType.IsGenericType)
                                    isOneToOneEntity = true;

                                currentType = isOneToOneEntity ? nestedModelProp.PropertyType : nestedModelProp.PropertyType.GetGenericArguments()[0];
                                currentModelProps = currentType.GetProperties();
                                currentModelList = new List<object?>();
                            }
                            else
                            {
                                parentDomainModelResultList.Errors.Add($"No property of parent model {typeof(T)}, found with the name '{sheet.Name}'");
                                break;
                            }
                            //end prepare current sheet types and data


                            //**** set values for each row and map instantly to parent
                            //prepare foreign key
                            var foreignParentIdKey = typeof(T).Name.Trim().ToLowerInvariant() + "id";
                            var foreignParentIdProp = currentModelProps.FirstOrDefault(p =>
                                 p.Name.ToLowerInvariant().Trim() == foreignParentIdKey.ToLowerInvariant().Trim());

                            for (var rowIndex = 2; rowIndex <= sheet.Dimension.End.Row; rowIndex++)
                            {

                                var currentModelObject = Activator.CreateInstance(currentType);
                                var errors = new List<string>();
                                foreach (var columnName in columnNames)
                                {

                                    var property = currentModelProps.FirstOrDefault(p => p.Name.ToLowerInvariant().Trim() == columnName.ToLowerInvariant().Trim());
                                    if (property != null)
                                    {
                                        // Set the property value based on the corresponding column value
                                        // You can implement any necessary data type conversions or validation here
                                        var columnIndex = columnNames.IndexOf(columnName) + 1;
                                        var cellValue = sheet.Cells[rowIndex, columnIndex].Value?.ToString();
                                        currentModelObject = SetValueConvertedToDatatype(ref property, currentModelObject, cellValue, rowIndex, currentType, ref errors);
                                    }
                                    else
                                    {
                                        parentDomainModelResultList.Errors.Add($"No property found with the name '{columnName}' in the model, child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                                    }
                                }

                                parentDomainModelResultList.Errors.AddRange(errors);
                                //end set values on the current object

                                //**** add the row object to its parent or skip

                                if (foreignParentIdProp != null)
                                {
                                    if (parentIdProp != null)
                                    {

                                        //map values of currentModelList to the each parent model object of List on parentid of parentIdKey = id of parent model

                                        //get parent id 
                                        var parentId = foreignParentIdProp.GetValue(currentModelObject);
                                        var foundParent = parentDomainModelResultList.ResultList.FirstOrDefault(x => ((Guid?)parentIdProp.GetValue(x.Result)).Equals((Guid?)parentId));
                                        if (foundParent != null && foundParent.Result != null)
                                        {
                                            if (isOneToOneEntity)
                                            {
                                                //get the prop of child in the parent model
                                                var childProp = parentModelProperties.FirstOrDefault(x =>
                                                x.Name.Trim().ToLowerInvariant() == currentType.Name.Trim().ToLowerInvariant());
                                                if (childProp != null)
                                                {
                                                    childProp.SetValue(foundParent.Result, currentModelObject);

                                                }
                                                else
                                                    foundParent.Errors.Add($"No child list found with name: '{currentType.Name.Trim().ToLowerInvariant() + "s"}' in the model name {typeof(T).Name}");
                                            }
                                            else
                                            {
                                                //get the prop of child list in the parent model
                                                var childListProp = parentModelProperties.FirstOrDefault(x =>
                                                x.Name.Trim().ToLowerInvariant() == currentType.Name.Trim().ToLowerInvariant() + "s");
                                                if (childListProp != null)
                                                {
                                                    //add modelobject to child list that list of same type of modelObject
                                                    //init
                                                    var childlist = childListProp.GetValue(foundParent.Result) as IList;
                                                    if (childlist == null)
                                                    {
                                                        var listType = typeof(List<>).MakeGenericType(currentType);
                                                        childlist = (IList)Activator.CreateInstance(listType);
                                                    }

                                                    childlist?.Add(currentModelObject);
                                                    //finally set the updated child to the parent
                                                    childListProp.SetValue(foundParent.Result, childlist);

                                                }
                                                else
                                                    foundParent.Errors.Add($"No child list found with name: '{currentType.Name.Trim().ToLowerInvariant() + "s"}' in the model name {typeof(T).Name}");
                                            }
                                        }
                                        else
                                            parentDomainModelResultList.Errors.Add($"No parent found with id : '{parentId}', child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                                    }
                                    else
                                        parentDomainModelResultList.Errors.Add($"No parentId found with the name (converted lowercase) 'id' in the model name {typeof(T).Name}");
                                }
                                else
                                    parentDomainModelResultList.Errors.Add($"No foreign parentId found with the name '{foreignParentIdKey}' in the model name {currentType.Name}");

                                //end map child item for its parent 


                            }

                        }

                        //in the end of each sheet check to set 
                        if (isTopParentSheet)
                            isTopParentSheet = false;
                    }
                }

                if (parentDomainModelResultList.Errors.Count > 0 || parentDomainModelResultList.ResultList.Any(x => x.Errors.Count > 0))
                {
                    //UPDATE THE CURRENT WORKBOOK FILE WITH NEXT UPDATES
                    //add column to the end of the parent sheet for the errors of each row (DomainModelResult.Errors) as list string in cell
                    //mark rows with red that conatins errors of parent sheet
                    //order the rows with errors to be first
                    //add another sheet of errors that is general (DomainModelResultList.Errors)
                    //in the created errors sheet add cell with link for each error to the corresponding first cell of row conatins the error if the error related to the row
                    //EXPORT THE UPDATED WORKBOOK FILE TO THE CLIENT TO BE DOWNLOADED
                    // Update the current workbook file with the necessary updates



                    using (var updatedPackage = new ExcelPackage(new FileInfo(uploadedFilePath)))
                    {
                        var parentSheet = updatedPackage.Workbook.Worksheets.FirstOrDefault();

                        // Add a column to the end of the parent sheet for the errors of each row (DomainModelResult.Errors) as a list of strings in a single cell
                        var errorsColumnIndex = parentSheet.Dimension.End.Column + 1;
                        var errorsColumnCell = parentSheet.Cells[1, errorsColumnIndex];
                        errorsColumnCell.Value = "Errors";
                        errorsColumnCell.Style.Font.Bold = true;

                        for (var rowIndex = 2; rowIndex <= parentSheet.Dimension.End.Row; rowIndex++)
                        {
                            var rowErrors = parentDomainModelResultList.ResultList.FirstOrDefault(x => x.RowIndex == rowIndex)?.Errors;
                            var rowErrorsCell = parentSheet.Cells[rowIndex, errorsColumnIndex];
                            rowErrorsCell.Value = string.Join("\n", rowErrors);
                            if (rowErrors.Count > 0)
                            {
                                //Mark errors with color                       
                                //rowErrorsCell.Style.Font.Color.SetColor(Color.PaleVioletRed);
                                parentSheet.Row(rowIndex).Style.Font.Color.SetColor(Color.PaleVioletRed);
                            }
                        }

                        if (parentDomainModelResultList.Errors.Count > 0)
                        {
                            // Create a new worksheet for general errors (DomainModelResultList.Errors)
                            var errorsSheet = updatedPackage.Workbook.Worksheets.Add("Errors");

                            // Add the general errors to the errors sheet
                            for (var i = 0; i < parentDomainModelResultList.Errors.Count; i++)
                            {
                                var error = parentDomainModelResultList.Errors[i];
                                var errorCell = errorsSheet.Cells[i + 1, 2];
                                errorCell.Value = error;

                                // Add a link to the corresponding first cell of the row containing the error
                                if (error.Contains("row No"))
                                {
                                    string pattern = @"child sheet name\s*:\s*<(.+?)>\s*,\s*row No\s*:\s*<(.+?)>";

                                    //extract row no and sheet name
                                    Match match = Regex.Match(error, pattern);

                                    if (match.Success)
                                    {
                                        string typeName = match.Groups[1].Value;
                                        string rowIndex = match.Groups[2].Value;

                                        Console.WriteLine("Type Name: " + typeName);
                                        Console.WriteLine("Row Index: " + rowIndex);

                                        // generate the cell link
                                        var linkCell = errorCell.Offset(0, -1);
                                        linkCell.Formula = $"HYPERLINK(\"#'{typeName}'!A{rowIndex}\", \"Go to Row\")";
                                        linkCell.Style.Font.Bold = true;
                                        linkCell.Style.Font.Color.SetColor(Color.BlueViolet);
                                    }
                                }
                            }
                        }
                        // Export the updated workbook file
                        updatedFilePath = Path.Combine(Path.GetDirectoryName(uploadedFilePath), (Guid.NewGuid().ToString()) + "_UpdatedWorkbook.xlsx");
                        updatedPackage.SaveAs(new FileInfo(updatedFilePath));

                        // Return the updated workbook file path to the client for download
                    }

                    return null;
                }

            }
            catch (Exception e)
            {
                parentDomainModelResultList.Errors.Add(e.Message);
            }

            // Return the imported domain model to insertupdate in db
            return parentDomainModelResultList;
        }


        private object SetValueConvertedToDatatype(ref PropertyInfo property, object currentModelObject, string? cellValue, int rowIndex, Type currentType, ref List<string> errors)
        {
            var propertyType = property.PropertyType;
            if (cellValue == null)
            {
                //errors.Add("cell value is null of property :" + property.Name+ $", child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                return currentModelObject;
            }
            if (propertyType == typeof(string) || Nullable.GetUnderlyingType(propertyType) == typeof(string))
            {
                property.SetValue(currentModelObject, cellValue.ToString());
            }
            else if (propertyType == typeof(char) || Nullable.GetUnderlyingType(propertyType) == typeof(char))
            {
                if (char.TryParse(cellValue.ToString(), out char charValue))
                {
                    property.SetValue(currentModelObject, charValue);
                }
                else
                {
                    // Handle conversion error
                    errors.Add($"Error converting value '{cellValue}' to type 'char' for property '{property.Name}' , child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                }

            }
            else if (propertyType == typeof(int) || Nullable.GetUnderlyingType(propertyType) == typeof(int))
            {
                if (int.TryParse(cellValue.ToString(), out int intValue))
                {
                    property.SetValue(currentModelObject, intValue);
                }
                else
                {
                    // Handle conversion error
                    errors.Add($"Error converting value '{cellValue}' to type 'int' for property '{property.Name}' , child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                }
            }
            else if (propertyType == typeof(decimal) || Nullable.GetUnderlyingType(propertyType) == typeof(decimal))
            {
                if (decimal.TryParse(cellValue.ToString(), out decimal decimalValue))
                {
                    property.SetValue(currentModelObject, decimalValue);
                }
                else
                {
                    // Handle conversion error
                    errors.Add($"Error converting value '{cellValue}' to type 'decimal' for property '{property.Name}', child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                }
            }
            else if (propertyType == typeof(Guid) || Nullable.GetUnderlyingType(propertyType) == typeof(Guid))
            {
                if (Guid.TryParse(cellValue.ToString(), out Guid value))
                {
                    property.SetValue(currentModelObject, value);
                }
                else
                {
                    // Handle conversion error
                    errors.Add($"Error converting value '{cellValue}' to type 'guid' for property '{property.Name}' , child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                }
            }
            else if (propertyType == typeof(DateTime) || Nullable.GetUnderlyingType(propertyType) == typeof(DateTime))
            {
                if (DateTime.TryParse(cellValue.ToString(), out DateTime dateTimeValue))
                {
                    property.SetValue(currentModelObject, dateTimeValue);
                }
                else
                {
                    // Handle conversion error
                    errors.Add($"Error converting value '{cellValue}' to type 'DateTime' for property '{property.Name}', child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
                }
            }
            else if (propertyType.IsEnum)
            {
                var enumValue = Enum.Parse(propertyType, cellValue);
                property.SetValue(currentModelObject, enumValue);
            }
            else
            {
                errors.Add($"No conversion for type '{propertyType}' for property '{property.Name}', child sheet name : <{currentType.Name.Trim().ToLowerInvariant()}> , row No : <{rowIndex}>");
            }
            return currentModelObject;
        }

        public string? ExportDomainModelListToExcel<T>(List<T> modelList)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Dictionary<string, int> currentRowIndex = new Dictionary<string, int>();
                using (var package = new ExcelPackage())
                {
                    // Create sheet
                    var parentSheetName = typeof(T).Name.Replace("Dto", "") + "s";
                    var parentSheet = package.Workbook.Worksheets.Add(parentSheetName);
                    var parentSheetProperties = typeof(T).GetProperties().Where(p => IsPrimitiveOrString(p.PropertyType)).ToList();

                    if (!currentRowIndex.ContainsKey(parentSheetName))
                        currentRowIndex[parentSheetName] = 1;

                    // Write column names in the sheet
                    for (var columnIndex = 1; columnIndex <= parentSheetProperties.Count; columnIndex++)
                    {
                        var propertyName = parentSheetProperties[columnIndex - 1].Name;
                        parentSheet.Cells[currentRowIndex[parentSheetName], columnIndex].Value = propertyName;
                    }

                    var childProperties = typeof(T).GetProperties().Where(p => !IsPrimitiveOrString(p.PropertyType)).ToList();

                    // Write data rows in the sheet
                    currentRowIndex[parentSheetName]++;

                    foreach (var model in modelList)
                    {
                        for (var columnIndex = 1; columnIndex <= parentSheetProperties.Count; columnIndex++)
                        {
                            var propertyValue = parentSheetProperties[columnIndex - 1].GetValue(model);
                            parentSheet.Cells[currentRowIndex[parentSheetName], columnIndex].Value = propertyValue;
                        }

                        currentRowIndex[parentSheetName]++;



                        // Create child sheets and write data rows for each child model
                        foreach (var childProperty in childProperties)
                        {
                            if (childProperty.PropertyType.IsGenericType && childProperty.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                            {
                                var childModels = (IList)childProperty.GetValue(model);
                                if (childModels != null && childModels.Count > 0)
                                {
                                    var childSheetName = childProperty.Name.Replace("Dto", "");

                                    var childSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == childSheetName) ?? package.Workbook.Worksheets.Add(childSheetName);

                                    // Write column names in the child sheet
                                    var childPropertiesList = childProperty.PropertyType.GenericTypeArguments[0].GetProperties().Where(p => IsPrimitiveOrString(p.PropertyType)).ToList();


                                    if (!currentRowIndex.ContainsKey(childSheetName))
                                    {
                                        currentRowIndex[childSheetName] = 1;//init

                                        for (var columnIndex = 1; columnIndex <= childPropertiesList.Count; columnIndex++)
                                        {
                                            var propertyName = childPropertiesList[columnIndex - 1].Name;
                                            childSheet.Cells[currentRowIndex[childSheetName], columnIndex].Value = propertyName;
                                        }

                                        currentRowIndex[childSheetName]++;
                                    }

                                    // Write data rows in the child sheet

                                    foreach (var childModel in childModels)
                                    {
                                        for (var columnIndex = 1; columnIndex <= childPropertiesList.Count; columnIndex++)
                                        {
                                            var propertyValue = childPropertiesList[columnIndex - 1].GetValue(childModel);
                                            childSheet.Cells[currentRowIndex[childSheetName], columnIndex].Value = propertyValue;
                                        }

                                        currentRowIndex[childSheetName]++;
                                    }
                                }
                            }
                            else if (childProperty.PropertyType.IsClass && !IsPrimitiveOrString(childProperty.PropertyType))
                            {
                                var childModel = childProperty.GetValue(model);
                                if (childModel != null)
                                {
                                    var childSheetName = childProperty.Name.Replace("Dto", "");
                                    var childSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == childSheetName) ?? package.Workbook.Worksheets.Add(childSheetName);

                                    // Write column names in the child sheet
                                    var childPropertiesList = childProperty.PropertyType.GetProperties().Where(p => IsPrimitiveOrString(p.PropertyType)).ToList();

                                    if (!currentRowIndex.ContainsKey(childSheetName))
                                    {
                                        currentRowIndex[childSheetName] = 1;//init

                                        for (var columnIndex = 1; columnIndex <= childPropertiesList.Count; columnIndex++)
                                        {
                                            var propertyName = childPropertiesList[columnIndex - 1].Name;
                                            childSheet.Cells[currentRowIndex[childSheetName], columnIndex].Value = propertyName;
                                        }

                                        currentRowIndex[childSheetName]++;
                                    }

                                    // Write data rows in the child sheet

                                    for (var columnIndex = 1; columnIndex <= childPropertiesList.Count; columnIndex++)
                                    {
                                        var propertyValue = childPropertiesList[columnIndex - 1].GetValue(childModel);
                                        childSheet.Cells[currentRowIndex[childSheetName], columnIndex].Value = propertyValue;
                                    }
                                    currentRowIndex[childSheetName]++;
                                }
                            }
                        }
                    }

                    // Save the Excel package to a file
                    var filePath = Path.Combine(Path.GetTempPath(), typeof(T).Name.Replace("Dto", "") + "s__" + Guid.NewGuid().ToString() + ".xlsx");
                    package.SaveAs(new FileInfo(filePath));

                    return filePath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        private static bool IsPrimitiveOrString(Type type)
        {
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                return IsPrimitiveOrString(Nullable.GetUnderlyingType(type));
            }
            return type.IsPrimitive || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime) || type == typeof(Guid);
        }

    }

}


