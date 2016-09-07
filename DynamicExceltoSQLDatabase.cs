using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Linq;
using Map = System.Data.Linq.Mapping;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel.DataAnnotations.Schema;
using ExceltoSQLDatabaseApplication;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.Data.Common;

namespace ExceltoSQLDatabaseApplication
{
    public partial class DynamicExceltoSQLDatabase : Form
    {
      

        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook wkb;
        private Worksheet sheet;

        private bool _tab1confirmation;

        public bool tab1confirmation
        {
            get { return this._tab1confirmation; }
            set { this._tab1confirmation = value; }
        }
        private string _database;

        public string database
        {
            get { return this._database; }
            set { this._database = value; }
        }
        private string _tablename;

        public string tablename
        {
            get { return this._tablename; }
            set { this._tablename = value; }
        }
        private string excelfile;
        private int header_row_num;

        public string _excelfile
        {
            get { return this._excelfile; }
            set { this._excelfile = value; }
        }

        public DynamicExceltoSQLDatabase()
        {
            InitializeComponent();
            if (this.tab1confirmation)
            {


            }


        }
        private void SetDefault(System.Windows.Forms.Button button2)
        {
            this.AcceptButton = button2;
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            this.excelfile = @"" + openFileDialog1.FileName;



        }



        private void Browse(object sender, EventArgs e)
        {
            int size = -1;
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                Console.WriteLine(file);
                try
                {
                    //   string text = File.ReadAllText(file);
                    //  size = text.Length;
                }
                catch (Exception)
                {
                }
            }
            Console.WriteLine(size); // <-- Shows file size in debugging mode.
            Console.WriteLine(result); // <-- For debugging use.
        }

        private void button2_Click(object sender, EventArgs e)
        {

            this.tablename = textBox2.Text;

            DataContext db = new DataContext(@"*connection string" + textBox1.Text + "*connection string");
            bool exist = db.DatabaseExists();



            if (exist)
            {
                this.printtext.Text = "The Database you selected exists";
                this.Load += new System.EventHandler(this.Form1_Load);
                this.database = @"*connection string" + textBox1.Text + "*connection string";
                if (this.tablename != null)
                {
                    this.tab1confirmation = true;
                }


            }
            else
            {
                this.printtext.Text = "The Database you selected does not exist";
                this.Load += new System.EventHandler(this.Form1_Load);
                this.tab1confirmation = false;

            }

        }
 
        private void Form1_Load(object sender, EventArgs e)
        {
            this.tab1confirmation = false;
        }

        private void Opentab2(object sender, EventArgs e)
        {
            if (this.tab1confirmation == true && this.UpdateTable.SelectedTab == tabPage2)
            {
                this.label5.Text = "Good!";
                this.Load += new System.EventHandler(this.Form1_Load);
            }
            else
            {
                this.label5.Text = "Complete Database and Excel File Section before proceeding!";
                this.Load += new System.EventHandler(this.Form1_Load);
            }
        }

        private void Submit_Click(object sender, EventArgs e)
        {
          
            int NumofString = int.Parse(NumberofString.Text);
            int NumDate = int.Parse(NumofDate.Text);
            int NumDecimal = int.Parse(NumofDecimal.Text);
            int NumInt = int.Parse(NumofInt.Text);

            int[] ColumnnumString = null;
            int[] ColumnnumInt = null; ;
            int[] ColumnnumDecimal = null;
            int[] ColumnnumDate = null;
         
     

            if (NumofString != 1)
            {
                var values = this.StringColumns.Text.Split(',');
                ColumnnumString = Array.ConvertAll(values, int.Parse);             
            }
            else if (NumofString == 1)
            {
                ColumnnumString = new int[1];
                ColumnnumString[0] = int.Parse(StringColumns.Text);
            }


            if (NumInt != 1)
            {
                var valuess = this.IntColumns.Text.Split(',');
                ColumnnumInt = Array.ConvertAll(valuess, int.Parse);
            }
            else if (NumInt == 1)
            {
                ColumnnumInt = new int[1];
                ColumnnumInt[0] = int.Parse(IntColumns.Text);
            }

            if (NumDecimal != 1)
            {
                var valuesss = this.DecimalColumns.Text.Split(',');
                ColumnnumDecimal = Array.ConvertAll(valuesss, int.Parse);
            }
            else if (NumDecimal == 1)
            {
                ColumnnumDecimal = new int[1];
                ColumnnumDecimal[0] = int.Parse(DecimalColumns.Text);
            }

            if (NumDate != 1)
            {
                var valuessss = this.DateColumns.Text.Split(',');
                ColumnnumDate = Array.ConvertAll(valuessss, int.Parse);
            }
            else if (NumDate == 1)
            {
                ColumnnumDate = new int[1];
                ColumnnumDate[0] = int.Parse(DateColumns.Text);
            }



          
            int headerrownum = Int32.Parse(header_row_number.Text);
            int NumberofRows = Int32.Parse(NumofRows.Text);
            int NumberofColumns = Int32.Parse(NumofColumns.Text);
          
          
                CreateTableDesign(NumofString, NumDate, NumDecimal, NumInt, ColumnnumString, ColumnnumInt, ColumnnumDecimal, ColumnnumDate, PrimaryKeyColumnNumber, headerrownum, NumberofRows, NumberofColumns);
            
         
          
        }

   

        private void CreateTableDesign(int NumofString, int NumDate, int NumDecimal, int NumInt, int[] ColumnNumsString, int[] ColumnNumInt, int[] ColumnnumDecimal, int[] ColumnnumDate, int PrimaryKeyColNum, int headerrownum, int NumofRows, int NumberofColumns)
        {




            excel = new Microsoft.Office.Interop.Excel.Application();
            wkb = null;
            sheet = null;
            wkb = excel.Workbooks.Open(this.excelfile,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
           Type.Missing, Type.Missing);
            sheet = wkb.Sheets[ExcelWorksheetName.Text] as Worksheet;

            Range range;



            Cell[] cells = new Cell[NumberofColumns + 1];


            bool match = false;
            for (int i = 1; i < NumberofColumns + 1; i++, match = false)  //Initialize Cell objects
            {
                range = sheet.Cells[header_row_number.Text, i];

                if (NumofString != 0 && match == false)
                {
                    for (int s = 0; s < NumofString; s++)
                    {
                        if (ColumnNumsString[s] == i)
                        {
                            cells[i] = new Cell
                            {
                                columnname = range.Text.ToString(),
                                columndatatype = "string",
                                columnnum = i
                            };
                            match = true;
                        }

                    }
                }
                if (NumInt != 0 && match == false)
                {
                    for (int intt = 0; intt < NumInt; intt++)
                    {
                        if (ColumnNumInt[intt] == i)
                        {
                            cells[i] = new Cell
                            {
                                columnname = range.Text.ToString(),
                                columndatatype = "int",
                                columnnum = i
                            };
                            match = true;
                        }

                    }
                }
                if (NumDecimal != 0 && match == false)
                {
                    for (int dec = 0; dec < NumDecimal; dec++)
                    {
                        if (ColumnnumDecimal[dec] == i)
                        {
                            cells[i] = new Cell
                            {
                                columnname = range.Text.ToString(),
                                columndatatype = "decimal",
                                columnnum = i
                            };
                            match = true;
                        }

                    }
                }
                if (NumDate != 0 && match == false)
                {
                    for (int date = 0; date < NumDate; date++)
                    {
                        if (ColumnnumDate[date] == i)
                        {
                            cells[i] = new Cell
                            {
                                columnname = range.Text.ToString(),
                                columndatatype = "date",
                                columnnum = i
                            };
                            match = true;
                        }

                    }
                }

            }


            /////////////////////////////////////////////////////////////////////////////////////// Cell objects Initialized, Create SQL Query


            string sqlstatement = "INSERT INTO " + tablename + " ";
            string sqlparams = "VALUES(@";




            for (int y = 1; y < NumberofColumns + 1; y++)
            {

                sqlparams += cells[y].columnname;


                if (y != NumberofColumns)
                {                   
                    sqlparams += ", @";
                }
                else
                {
                    sqlparams += ")";
                }
            }
            sqlstatement += sqlparams;

            PopulateTable(headerrownum, NumofRows, NumberofColumns, cells, sqlstatement, sheet);

        }


        private void PopulateTable(int headerrownum, int NumofRows, int NumberofColumns, Cell[] cells, string sqlstatement, Worksheet sheet)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = @"*connection string" + textBox1.Text + "*connection string";
            Range range;

            conn.Open();
            SqlCommand cmd = new SqlCommand(sqlstatement, conn);
            for (int y = 1; y < NumberofColumns + 1; y++)
            {

                if (cells[y].columndatatype == "string")
                {

                    cmd.Parameters.Add("@" + cells[y].columnname, SqlDbType.VarChar, 100);

                }
                if (cells[y].columndatatype == "int")
                {


                    cmd.Parameters.Add("@" + cells[y].columnname, SqlDbType.Int);
                }

                if (cells[y].columndatatype == "date")
                {


                    cmd.Parameters.Add("@" + cells[y].columnname, SqlDbType.DateTime);
                }
                if (cells[y].columndatatype == "decimal")
                {


                    cmd.Parameters.Add("@" + cells[y].columnname, SqlDbType.Decimal);
                }
            }



            int a = 0;
            for (int i = headerrownum + 1, k = 1; i < NumofRows + 1; k++)
            {
                if (k != NumberofColumns + 1)
                {
                    if (cells[k].columndatatype == "string")
                    {
                        range = sheet.Cells[i, k];
                        cmd.Parameters["@" + cells[k].columnname].Value = range.Text.ToString();

                    }
                    if (cells[k].columndatatype == "int")
                    {
                        range = sheet.Cells[i, k];

                        cmd.Parameters["@" + cells[k].columnname].Value = Int32.Parse(range.Text.ToString());
                    }

                    if (cells[k].columndatatype == "date")
                    {
                        range = sheet.Cells[i, k];

                        cmd.Parameters["@" + cells[k].columnname].Value = DateTime.Parse(range.Text.ToString());
                    }
                    if (cells[k].columndatatype == "decimal")
                    {
                        range = sheet.Cells[i, k];

                        cmd.Parameters["@" + cells[k].columnname].Value = System.Convert.ToDecimal(range.Value);
                    }
                    if (k == NumberofColumns)
                    {
                        a = cmd.ExecuteNonQuery();
                        k = 0;
                        i++;
                    }

                }



            }

            conn.Close();

         








        }


      
    
    }
}
   