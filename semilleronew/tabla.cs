using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace semilleronew
{
    class tabla
    {
        void readtab()
        {
            string[] i = new string[1000];
            int valor = 5;

            string Direccion = "C:/Users/Legion-Pc/Desktop/WeibullViento.xslx";
            SLDocument sl = new SLDocument(Direccion);


            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 6)))
            {

                Console.WriteLine(sl.GetCellValueAsString(valor, 6));
                valor++;
            }
            Console.Read();
        }

        
    }
}
