using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_doc_compare
{
    class company
    {
        private string name;
        private string reg_nr;
        private string link;

        public company(string name, string reg_nr, string link)
        {

            this.name = name;
            this.reg_nr = reg_nr;
            this.link = link;
        }

        public string getName()
        {
            return this.name;
        }

        public string getLink()
        {
            return this.link;
        }

        public string getNr()
        {
            return this.reg_nr;
        }
    }
}
