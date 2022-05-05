using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automate_AR
{
    public class Job
    {
        public Job(string id = "", string name = "", bool late = false)
        {
            this.id = id;
            this.name = name;
            this.late = late;
        }

        public string id = "";
        public string name = "";
        public string contactID = "";
        public List<Invoice> invoices = new List<Invoice>();
        public bool late = false;
    }
}
