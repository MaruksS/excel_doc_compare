using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;

namespace excel_doc_compare
{
    public partial class Form1 : Form
    {
        string name, reg_nr, link;
        List<company> companies = new List<company>();
        int matches;

        public Form1()
        {
            InitializeComponent();
            readSparrowExl();
            readDataExl();
        }


        /* 
         * Read sparrow export document
         */
        public void readSparrowExl()
        {
            //Define excel document
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\docs\document.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Limit loop to row.count
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                //Read value from cell
                var value = xlRange.Cells[i,1].Value2.ToString();

                if (value.Contains("Firma:"))
                {
                    name = value.Remove(0, 7);
                }
                else if (value.Contains("numurs:")){
                    reg_nr = Regex.Replace(value, "[^.0-9]", "");
                }
                else if (value.Contains(".lv"))
                {
                    link = value;
                }
                else if (value.Contains("-------"))
                {
                    //If dashes - record has ended, creating an object
                    company new_comp = new company(name, reg_nr, link);
                    companies.Add(new_comp);

                    name = null;
                    reg_nr = null;
                    link = null;
                } //End of if
            } //End of for loop
        } //End of readSparrowExl


        /* 
         * Read Dante export document
         */
        public void readDataExl()
        {
            //Define excel document
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\docs\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //Limit loop to row.count
            int rowCount = xlRange.Rows.Count;

            //amount of positive matches
            matches = 0;

            //Start writer to txt file
            StreamWriter sw = new StreamWriter(@"C:\docs\output.txt");

            for (int i = 1; i <= rowCount; i++)
            {
                string value = null;

                //Read 18th column, which is Identification
                if (xlRange.Cells[i, 18].Value2 != null)
                {
                    value = xlRange.Cells[i, 18].Value2.ToString();

                    //if Identification column is not empty, compare with each record from list
                    foreach (company comp in companies){

                        string reg_nr = comp.getNr();

                        /*
                         * if Identification column contains registration number from news update
                         * write it to the txt file with link and information
                         */
                        if (value.Contains(reg_nr))
                        {
                            string uid = xlRange.Cells[i, 1].Value2.ToString();
                            string name = comp.getName();
                            string link = comp.getLink();

                            sw.WriteLine("UID: "+ uid);
                            sw.WriteLine(name + " : "+ reg_nr);
                            sw.WriteLine(link);
                            sw.WriteLine("-------------");

                            matches++;
                        } //End of if
                    } //End of foreach
                } //End of if
            } //End of loop

            // end writer
            sw.Close();
            // update label text
            lbl_message.Text = "Updates for " + matches + " companies were found";
        } //End of readDataExl
    }


}
