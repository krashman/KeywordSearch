using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Code7248.word_reader;

namespace WindowsFormsApplication2
{
    class DOCXFileParser : Parsers
    {
        //string filePath;

        public string gettext(string filePath){
            string fileData = "";

            Code7248.word_reader.TextExtractor extractor = new TextExtractor(filePath);

            fileData = extractor.ExtractText();

            return fileData;
        }
    }
}
