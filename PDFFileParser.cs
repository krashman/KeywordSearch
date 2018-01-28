using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    class PDFFileParser : Parsers
    {
        //string filePath;

        public string gettext(string filePath){
                 string textdata = getPDFtext(filePath);
                 return textdata;
            }

        public static string getPDFtext(string filePath){
            PDDocument pdfFile = null;
            try
            {
                pdfFile = PDDocument.load(filePath);
                PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(pdfFile);
            }
            catch (Exception e) {
                MessageBox.Show("error "+e);
                return "";
            }
            finally{
                if (pdfFile != null) {
                    pdfFile.close();
                }
            }
        }
    }
}
