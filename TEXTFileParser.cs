using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication2
{
    class TEXTFileParser : Parsers
    {
        private string fileData = "";
        public string gettext(string filePath){
            try
            {
                
                // Create an instance of StreamReader to read from a file. 
                // The using statement also closes the StreamReader. 
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string line;
                    
                    // Read and display lines from the file until the end of  
                    // the file is reached. 
                    while ((line = sr.ReadLine()) != null)
                    {
                        this.fileData += line;
                    }
                }
            }catch (Exception e){
                // Let the user know what went wrong.
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
            return this.fileData;
        }
    }
}
