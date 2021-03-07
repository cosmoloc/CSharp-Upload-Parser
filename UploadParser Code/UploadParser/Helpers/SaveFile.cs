using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace UploadParser.Helpers
{
    class SaveFile
    {

        // Save file to local system
        public string parseAndSave(Stream stream, bool save, string path)
        {
            byte[] data = new byte[32768];
            string content = "";
            string fileName = "";
            string extension = "";
            string ContentType = "";
            try
            {
                data = ToByteArray(stream);
                content = Encoding.UTF8.GetString(data);

                // The first line should contain the delimiter
                int delimiterEndIndex = content.IndexOf("\r\n");

                if (delimiterEndIndex > -1)
                {
                    string delimiter = content.Substring(0, content.IndexOf("\r\n"));

                    //Content-Type
                    Regex re = new Regex(@"(?<=Content\-Type:)(.*?)(?=\r\n\r\n)");
                    Match contentTypeMatch = re.Match(content);

                    //Filename
                    re = new Regex(@"(?<=filename\=\"")(.*?)(?=\"")");
                    Match filenameMatch = re.Match(content);

                    if (contentTypeMatch.Success && filenameMatch.Success)
                    {
                        // Set properties
                        ContentType = contentTypeMatch.Value.Trim();
                        fileName = filenameMatch.Value.Trim();
                        extension = Path.GetExtension(fileName);

                        // Covers cases : If user does not wish to save file (then temporarily save file in default plugin path), else if user wishes to save file but no path provided
                        if (String.IsNullOrEmpty(path))
                        {
                            path = UploadConstants.DEFAULT_SAVE_PATH;
                        }

                        // Get the start & end indexes of the file contents
                        int startIndex = contentTypeMatch.Index + contentTypeMatch.Length + "\r\n\r\n".Length;
                        int endIndex = 0;
                        byte[] delimiterBytes = Encoding.UTF8.GetBytes("\r\n" + delimiter);

                        //Calculate endIndex
                        int startPos = Array.IndexOf(data, delimiterBytes[0], startIndex);

                        if (startPos != -1)
                        {
                            while ((startPos + endIndex) < data.Length)
                            {
                                if (data[startPos + endIndex] == delimiterBytes[endIndex])
                                {
                                    endIndex++;
                                    if (endIndex == delimiterBytes.Length)
                                    {
                                        endIndex = startPos;
                                    }
                                }
                                else
                                {
                                    startPos = Array.IndexOf<byte>(data, delimiterBytes[0], startPos + endIndex);
                                    if (startPos != -1)
                                    {
                                        endIndex = 0;
                                    }
                                }
                            }
                        }

                        int contentLength = endIndex - startIndex;

                        // Extract the file contents from the byte array
                        byte[] fileData = new byte[contentLength];
                        Buffer.BlockCopy(data, startIndex, fileData, 0, contentLength);

                        //Save locally for parsing... delete after parsing completes in case the user doesnt wishes to save the file on the system)
                        fileName = ByteArrayToFile(fileName, fileData, path);                        
                    }
                }
                return fileName;
            }

            catch (Exception e)
            {
                return "Exception in parseAndSave: Could not parse Uploaded file ! Exception : " + e.Message;
            }

        }

        private byte[] ToByteArray(Stream stream)
        {
            byte[] buffer = new byte[32768];
            MemoryStream memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }

        private string ByteArrayToFile(string fileName, byte[] _ByteArray, string path)
        {
            try
            {

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                fileName = path + Path.GetFileNameWithoutExtension(fileName) + "_" + DateTime.Now.ToString("_yyyyMMdd_HHmmssfff") + Path.GetExtension(fileName);
                System.IO.FileStream _FileStream = new System.IO.FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);

                // Writes a block of bytes to this stream using data from a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // close file stream
                _FileStream.Close();

                return fileName;
            }
            catch (Exception e)
            {
                throw new Exception("Exception caught in ByteArrayToFile in creating backup process:" + e);
            }

        }

    }
}
