using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace UploadParser.Helpers
{
    public class ExcelFile
    {
        public string fileName;
        public string[] sheetName;
        public string[][][] sheetData;  // x= sheet index, y= row number, z= cell value

        //Default Constructor
        public ExcelFile() { }

        //Parameterized Constructor
        public ExcelFile(string filename, int numSheets)
        {
            fileName = filename;
            sheetName = new string[numSheets];
            sheetData = new string[numSheets][][];
        }
    }

    public class CSVFile
    {
        public string fileName;
        public string[][] csvData;

        public CSVFile()
        { }

        public CSVFile(string name, int num)
        {
            fileName = name;
            csvData = new string[num][];
        }
    }
}
