using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace TTTWordHelper
{
    public class Class1
    {
        public Microsoft.Office.Interop.Word.Application app;

        public test()
        {
            app = app.ApplicationClass();
            app.Documents.Open()
        }
    }
}
