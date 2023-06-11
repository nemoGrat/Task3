using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Classes
{
    internal class InfoAboutClients
    {
        public string clientCode;
        public string nameOfOrganization;
        public string adressOfOrganizaton;
        public string contactNameOfOrganization;
        public int numOfOrders;

        public InfoAboutClients(string clientCode, string nameOfOrganization, string adressOfOrganizaton, string contactNameOfOrganization, int numOfOrders)
        {
            this.clientCode = clientCode;
            this.nameOfOrganization = nameOfOrganization;
            this.numOfOrders = numOfOrders;
            this.adressOfOrganizaton = adressOfOrganizaton;
            this.contactNameOfOrganization = contactNameOfOrganization;
        }

    }
}
