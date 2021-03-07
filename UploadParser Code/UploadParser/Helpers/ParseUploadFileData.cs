using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Text.RegularExpressions;
using UploadParser;

namespace UploadParser.Helpers
{
    class ParseUploadFileData
    {
        #region processFile
        public Object processFile(string fileName)
        {
            string res = "";
            ExcelFile excelFileOutput = new ExcelFile();
            CSVFile csvFileOutput = new CSVFile();
            try
            {
                string extension = Path.GetExtension(fileName);

                if (extension.Equals(".xls"))
                {
                    excelFileOutput = processXLS(fileName);
                    return excelFileOutput;
                }

                else if (extension.Equals(".csv"))
                {
                    csvFileOutput = processCSV(fileName);
                    return csvFileOutput;
                }

                return null;
            }

            catch (Exception e)
            {
                throw new Exception("Exception in ParseUploadFileData -> processFile()" + e);
            }
        }
        #endregion

        #region processXLS ( ExcelFile processXLS(string fileName))
        #region Main function
        public ExcelFile processXLS(string fileName)
        {
            try
            {
                //Load xls file to parse
                FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = new HSSFWorkbook(file);
                int numSheets = workbook.NumberOfSheets;

                // CREATE OBJECT
                ExcelFile excelFile = new ExcelFile(Path.GetFileName(fileName), numSheets);

                for (int i = 0; i < numSheets; i++)
                {
                    HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(i);
                    excelFile.sheetName[i] = sheet.SheetName;

                    //Get Sheet Data
                    string[][] result = processSheet(sheet);

                    //Save to ExcelFile object 
                    excelFile.sheetData[i] = CopyArrayBuiltIn(result);

                }
                return excelFile;
            }
            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> processXLS: " + e);
            }
        }
        #endregion

        #region Copy the retrieved jagged array to excel class member
        static string[][] CopyArrayBuiltIn(string[][] source)
        {
            var len = source.Length;
            var dest = new string[len][];
            try
            {
                for (var x = 0; x < len; x++)
                {
                    var inner = source[x];
                    if (inner == null)
                        continue;
                    var ilen = inner.Length;
                    var newer = new string[ilen];
                    Array.Copy(inner, newer, ilen);
                    dest[x] = newer;
                }
                return dest;
            }

            catch (Exception e)
            {
                throw new Exception("Exception in ParseUploadFileData -> CopyArrayBuiltIn : " + e.Message);
            }

        }
        #endregion

        #region Process Sheet
        public string[][] processSheet(HSSFSheet sheet)
        {
            int rowIter = 0;
            int colIter = 0;

            int rowNum = sheet.LastRowNum + 1;
            int nullRowCount = 0;
            TriggerParse trigger = new TriggerParse();

            string[][] data = new string[rowNum][];

            try
            {
                //Traverse rows until 50 null rows
                while (nullRowCount <= trigger.EMPTYROW_THRESHOLD)
                {
                    try
                    {
                        IRow irow = sheet.GetRow(rowIter);
                        if (irow == null || irow.LastCellNum == -1)
                        {
                            nullRowCount++;
                            rowIter++;
                            continue;
                        }
                        data[rowIter] = new string[irow.LastCellNum];
                        foreach (ICell cell in irow)
                        {
                            colIter = cell.ColumnIndex;
                            data[rowIter][colIter] = getCellValue(cell);
                        }
                        rowIter++;
                    }
                    catch (Exception ex)
                    { // Do Nothing
                    }
                }

                if (data != null)
                {
                    return data;
                }
                else
                    return null;
            }

            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> processSheet() : (Sheet : " + sheet.SheetName + ", Row: " + rowIter + ", Column : " + colIter + ") Exception: " + e);
            }
        }
        #endregion

        #region Identify and get cell value
        public String getCellValue(ICell cell)
        {
            String retVal = "";

            try
            {

                if (cell == null)
                {
                    return "";
                }
                string cellType = cell.CellType.ToString();
                switch (cellType)
                {
                    case "BLANK": retVal = "";
                        break;
                    case "BOOLEAN": retVal = "" + cell.BooleanCellValue.ToString();
                        break;
                    case "STRING": retVal = cell.StringCellValue.ToString();
                        break;
                    case "NUMERIC": retVal = isNumberOrDate(cell);
                        break;
                    case "FORMULA": retVal = processFormula(cell);
                        break;
                    default: retVal = "";
                        break;
                }
                return retVal;
            }

            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> getCellValue() : (Row: " + cell.RowIndex + ", Column : " + cell.ColumnIndex + ") Exception: " + e);
            }
        }
        #endregion

        #region Data processing helper methods
        private String isNumberOrDate(ICell cell)
        {
            String retVal;

            try
            {
                if (HSSFDateUtil.IsCellDateFormatted(cell))
                {
                    retVal = cell.DateCellValue.ToLongDateString();
                }
                else
                {
                    retVal = cell.NumericCellValue.ToString();
                }
                return retVal;
            }
            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> isNumberOrDate()  : (Row: " + cell.RowIndex + ", Column : " + cell.ColumnIndex + ") Exception: " + e);
            }
        }


        private String processFormula(ICell cell)
        {
            String retVal = "";
            try
            {
                IFormulaEvaluator evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                HSSFDataFormatter formatter = new HSSFDataFormatter();
                if (cell.CachedFormulaResultType == CellType.ERROR)
                {
                    retVal = "#VALUE!";
                }
                else
                {
                    retVal = formatter.FormatCellValue(cell, evaluator);
                    //if (retVal.matches("[0-9]+"))
                    //{
                    try
                    {
                        if (HSSFDateUtil.IsCellDateFormatted(cell))
                        {
                            retVal = cell.DateCellValue.ToLongDateString();
                        }

                        else
                        {
                            retVal = formatter.FormatCellValue(cell);
                        }
                    }

                    catch (InvalidOperationException ex)
                    {
                        //Do Nothing : XLS Sheet cell type not in conformity with the value inside
                    }

                    // }
                }
            }
            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> isNumberOrDate()  : (Row: " + cell.RowIndex + ", Column : " + cell.ColumnIndex + ") Exception: " + e);
            }
            return retVal;
        }
        #endregion
        #endregion

        #region processCSV
        public CSVFile processCSV(string filename)
        {
            CSVFile csvFile = null;
            string[][] data = null;
            int rowCounter = 0;

            try
            {
                //string[] lines = File.ReadAllLines(filename);
                string contents = File.ReadAllText(filename);

                //Clean bad data -> Remove \n (without preceded by a \r) so it is not is parsed with \r\n thus producing extra rows
                contents = Regex.Replace(contents, "(?<!\r)\n", "\\n");

                string[] lines = Regex.Split(contents, "(\r\n){1,}", RegexOptions.ExplicitCapture);

                int rowCount = lines.Count();

                csvFile = new CSVFile(Path.GetFileName(filename), rowCount);

                data = new string[rowCount][];

                foreach (string line in lines)
                {
                    int colCounter = 0;

                    // Split line with "," character outside the braces
                    var values = Regex.Matches(line, @"(((?<x>(?=[,\r\n]+))|""(?<x>([^""]|"""")+)""|(?<x>[^,\r\n]+)),?)", System.Text.RegularExpressions.RegexOptions.Multiline);

                    //Initialize firs row of jagged array -> data
                    data[rowCounter] = new string[values.Count];

                    // Insert row values to data
                    foreach (Match value in values)
                    {
                        data[rowCounter][colCounter] = value.Groups["x"].Value;
                        colCounter++;
                        Console.WriteLine(value.Groups["m"].Value);
                    }

                    rowCounter++;
                }
                csvFile.csvData = data;
            }

            catch (Exception e)
            {
                throw new Exception("\nException in ParseUploadFileData -> processCSV: " + e);
            }

            return csvFile;
        }
        #endregion
    }
}
