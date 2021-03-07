using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using UploadParser.Helpers;

namespace UploadParser
{
    
    public class TriggerParse
    {

        public int EMPTYROW_THRESHOLD = 50;

        public Object parseFile(Stream fileUpload, bool save, string path)
        {
            ExcelFile excelFileOutput = new ExcelFile();
            CSVFile csvFileOutput = new CSVFile();
            string fileName = "";          
            
            Object obj = null;

            //Save File
            SaveFile saveFile = new SaveFile();
          
            //Get File
            ParseUploadFileData parseFileUploadData = new ParseUploadFileData();
            
            try
            {
                fileName = saveFile.parseAndSave(fileUpload, save, path);
                string extension = Path.GetExtension(fileName);

                if (fileName == null || fileName == "")
                {
                    throw new Exception("File not retreived/invalid. Error in data");
                }

                if (extension == ".xls")
                {
                    excelFileOutput = (ExcelFile)parseFileUploadData.processFile(fileName);
                    obj = new ExcelFile();
                    obj = excelFileOutput;
                }

                else if (extension == ".csv")
                {
                    csvFileOutput = (CSVFile)parseFileUploadData.processFile(fileName);
                    obj = new CSVFile();
                    obj = csvFileOutput;
                }
                if (!save)
                {
                    if (!(fileName == null))
                    {
                        File.Delete(fileName);
                    }                    
                }

                if (path == "")
                {
                    string[] allFiles = Directory.GetFiles(UploadConstants.DEFAULT_SAVE_PATH);
                    if (allFiles.Length == 0)
                    {
                        Directory.Delete(UploadConstants.DEFAULT_SAVE_PATH);
                    }
                }
                return obj;
            }

            catch (Exception ex)
            {
                if (!(fileName == null))
                {
                    File.Delete(fileName);
                }

                if (path == "")
                {
                    string[] allFiles = Directory.GetFiles(UploadConstants.DEFAULT_SAVE_PATH);
                    if (allFiles.Length == 0)
                    {
                        Directory.Delete(UploadConstants.DEFAULT_SAVE_PATH);
                    }
                }
                return null;
            }
        }

    }
}
