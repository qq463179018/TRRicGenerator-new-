using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PdfTronWrapper.TableBorder
{
    public class FormLineList:List<FormLine>
    {
        public FormLineList(List<FormLine> formLines)
        {
            formLines.ForEach(line => Add(line));
        }

        public FormLineList(FormLine formLine)
        {
            Add(formLine);
        }

        public FormLineList()
        {

        }
    }
}
