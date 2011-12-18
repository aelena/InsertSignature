using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Utils.Office;

namespace InsertaFirma
{
    class Program
    {
        static void Main(string[] args)
        {
            OfficeSigner oSign = new OfficeSigner();
            //oSign.InsertScannedSignature(@"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Sample.docx",
            //    @"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Louis-xiv-signature.jpg", true);
            //oSign.InsertScannedSignature(@"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Sample.doc",
            //    @"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Louis-xiv-signature.jpg", true);
            oSign.InsertScannedSignature(@"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Sample.xlsx",
                @"D:\Projects\InsertaFirma\InsertaFirma\bin\Debug\Louis-xiv-signature.jpg", true);

        }
    }
}
