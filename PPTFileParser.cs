using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    class PPTFileParser : Parsers
    {
        //string filePath;

        public string gettext(string filePath)
        {
            string fileData = "";
            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(@filePath, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
         
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                foreach (var item in presentation.Slides[i + 1].Shapes)
                {
                    var shape = (PowerPoint.Shape)item;
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            var text = textRange.Text;
                            fileData += text + " ";
                        }
                    }
                }
            }
            PowerPoint_App.Quit();
            return fileData;
        }
    }
}
