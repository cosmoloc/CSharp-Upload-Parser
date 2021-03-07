# CSharp-Upload-Parser
Upload and Parse XLS and CSV Files in C#

### Description

Upload Parser is a library that can be used to parse an uploaded  XLS or CSV file. The parser receives the uploaded file stream along with some added parameters from the program, and returns a class object containing complete file data.


Following features are provided by the library:

Parse entire file and return a class with complete file data encapsulated in it.

An option for developer to save the uploaded file on the system/server or not.

An option for developer to set their own path where the uploaded files should be saved.

Currently the plugin supports excel files(.xls) and CSV(.csv) Files. Its functionality will be extended in future to support more file types.

### Plugin Usage

Plugin can be used in a code by implementing the following steps:

1. Add a reference of the library **UploadParser.dll** file in the program

2. Add a reference of the package : **NPOI.dll** in the program (to support xls workbook operations)

3. Include the UploadParser namespace in the code, i.e.: 

    `using UploadParser;`

4. Create an object of the class : TriggerParser, i.e.: 

    `TriggerParser <myTriggerObjectName> = new TriggerParser();`

5. Call the method parseFile() to trigger the parsing and get the desired result in Object type , i.e.:

    `Object <myExcelObjectName >= triggerParse.parseFile (<streamObj>, <save?>, <pathToSave>)`


### Plugin Defaults

1. Default directory to save uploaded file
UploadParser defines the default path to save the uploaded file to the server, in case the user passes an empty string to the called method.

This default value is :

 `public const string DEFAULT_SAVE_PATH = @"C:\\UPLOAD_PARSER\";`

Thus, if save= true and path = "", the uploaded file will be saved to *C:\\UPLOAD_PARSER\* directory.

To set own value, user would be required to pass the value of the desired path in the function : 

`triggerParse.parseFile (Stream fileUpload, bool save, string path)`

For Ex :   

`Object obj = triggerParse.parseFile(csvUpload, true, @"C:\\Loadload\");`


2. Default number of empty rows until the parser stops parsing the file and accepts EOF

Apart from identifying  the last row number programmatically,  it may be possible that an excel file may contain empty undefined rows in the end.

So, the UploadParser also defines a default integer value : 

`public int EMPTYROW_THRESHOLD = 50;`

i.e after 50 null rows, the parser would itself imagine EOF and thus stop parsing.

This value can be altered by the user. Before calling the function to parse the uploaded file, add the following line of code :

`triggerParse.EMPTYROW_THRESHOLD = 10;`

Thus the final code to call the plugin would be : 

`using UploadParser;`

`...`

`// Funtion in main program to call the plugin`

`public string parseUploadedFie(Stream xlsUpload)  `

`{`

 `        ...`
 
 `	TriggerParse triggerParse = new TriggerParse();`
 
  `       public int EMPTYROW_THRESHOLD = 50;`
 
 `       Object File = triggerParse.parseFile(fileUpload,true, @"C:\\MyDir\");`
 
 `        ...`
 
`}`

3) Default string to be replaced with “\n”

In UploadParser, it is assumed that a row terminates with “\r\n”.

But there may exist entries which contain “\n”, and thus may show up in excel file as a new line, even though it is a part of the same row entry.

To handle such issue, Upload Parser identifies the special character : “\n” and by default replaces it with “\\n”.

`contents = Regex.Replace(contents, "(?<!\r)\n", "");`


# CSharp-Upload-Parser
Parse an  XLS or CSV file to a generic class object
