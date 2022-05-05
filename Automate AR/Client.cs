using System;
using System.Collections.Generic;

namespace Automate_AR {

    public class Client
    {
        public Client(string contactID = "", string email = "")
        {
            this.contactID = contactID;
            this.email = email;
        }

        public string contactID = "";
        public string email = "";
        public List<Job> jobs = new List<Job>();
    }
}
