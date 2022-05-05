using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automate_AR
{
    public class Invoice
    {
        public Invoice(string id = "", string extension = "")
        {
            this.id = id;
            this.extension = extension;
        }

        public string id = "";
        public string extension = "";
        public string file_id = "";
        public byte[] invoice_data = { };
    }
}
