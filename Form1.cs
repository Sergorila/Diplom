using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelDataReader;
using MaterialSkin.Controls;

namespace WinForm
{

    public partial class Form1 : Form
    {

        private DataSet ds;
        IExcelDataReader reader = null;
        OpenFileDialog openFileDialog = new OpenFileDialog();
        

        public Form1()
        {
            InitializeComponent();
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fName = openFileDialog.FileName;
                var file = new FileInfo(fName);
                try
                {
                    var col1 = new DataGridViewTextBoxColumn();
                    var col2 = new DataGridViewTextBoxColumn();
                    var col3 = new DataGridViewTextBoxColumn();
                    var col4 = new DataGridViewTextBoxColumn();
                    var col5 = new DataGridViewTextBoxColumn();
                    var col6 = new DataGridViewTextBoxColumn();
                    var col7 = new DataGridViewTextBoxColumn();
                    var col8 = new DataGridViewTextBoxColumn();
                    var col9 = new DataGridViewTextBoxColumn();
                    var col10 = new DataGridViewTextBoxColumn();
                    var col11 = new DataGridViewTextBoxColumn();
                    var col12 = new DataGridViewTextBoxColumn();

                    col1.HeaderText = "левый-правый";
                    col1.Name = "col1";

                    col2.HeaderText = "угол и покрытие";
                    col2.Name = "col2";

                    col3.HeaderText = "угол";
                    col3.Name = "col3";

                    col4.HeaderText = "покрытие лог";
                    col4.Name = "col4";

                    col5.HeaderText = "покрытие";
                    col5.Name = "col5";

                    col6.HeaderText = "ножка просвет";
                    col6.Name = "col6";

                    col7.HeaderText = "другая нога";
                    col7.Name = "col7";

                    col8.HeaderText = "Сужение суставной щели другой ноги";
                    col8.Name = "col8";

                    col9.HeaderText = "код Сужение";
                    col9.Name = "col9";

                    col10.HeaderText = "суст поверхности";
                    col10.Name = "col10";

                    col11.HeaderText = "центрирование головки";
                    col11.Name = "col11";

                    col12.HeaderText = "заключение";
                    col12.Name = "col12";

                    using (var stream = new FileStream(fName, FileMode.Open))
                    {
                        if (reader != null) { reader = null; }

                        // Judge it is .xls or .xlsx
                        if (file.Extension == ".xls") { reader = ExcelReaderFactory.CreateBinaryReader(stream); }
                        else if (file.Extension == ".xlsx") { reader = ExcelReaderFactory.CreateOpenXmlReader(stream); }

                        if (reader == null) { return; }
                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            UseColumnDataType = true,
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = false,
                                ReadHeaderRow = (rowReader) => {
                                    rowReader.Read();
                                },

                                // Gets or sets a callback to determine whether to include the 
                                // current row in the DataTable.
                                FilterRow = (rowReader) => {
                                    return true;
                                },
                            }
                        });

                        var tablenames = ds.Tables;
                        dataGridView1.DataSource = tablenames[0];
                        

                        dataGridView1.Columns.AddRange(new DataGridViewColumn[] { col1, col2, col3, col4, col5, col6, col7
                                                                                    , col8, col9, col10, col11, col12});

                        

                    }

                    using (StreamReader stream = new StreamReader(@"C:\Users\Sergey\Desktop\Kursovaya\prediction.txt"))
                    {
                        int position = 0;
                        string temp;
                        while ((temp = stream.ReadLine()) != null)
                        {
                            string info = dataGridView1[1, position + 1].Value.ToString();
                            info = info.Replace("{", "");
                            info = info.Replace("}", "");
                            string[] lines = info.Split('.');
                            temp = temp.Replace("[", "");
                            temp = temp.Replace("]", "");
                            string[] tempSplit = temp.Split(' ');


                            int[] indexes = new int[tempSplit.Length];

                            int k = 0;
                            foreach (var item in tempSplit)
                            {
                                indexes[k] = int.Parse(item);
                                k++;
                            }

                            for (int i = 0; i < indexes.Length; i++)
                            {
                                if (lines[i].Contains("аключение"))
                                {
                                    dataGridView1[13, position + 1].Value += lines[i];
                                }
                                else
                                {
                                    if (indexes[i] <= 2)
                                    {
                                        dataGridView1[indexes[i] + 1, position + 1].Value = lines[i];

                                        if (indexes[i] == 2)
                                        {
                                            string[] details = lines[i].Split(',');
                                            dataGridView1[indexes[i] + 4, position + 1].Value = details[1];

                                            int value;
                                            int.TryParse(string.Join("", details[0].Where(c => char.IsDigit(c))), out value);

                                            dataGridView1[indexes[i] + 2, position + 1].Value = value.ToString();

                                            if (details[1].Contains("неполная") || details[1].Contains("не полная"))
                                            {
                                                dataGridView1[indexes[i] + 3, position + 1].Value = "0";
                                            }
                                            else
                                            {
                                                dataGridView1[indexes[i] + 3, position + 1].Value = "1";
                                            }
                                        }
                                    }

                                    if (indexes[i] >= 3 && indexes[i] <= 5)
                                    {
                                        dataGridView1[indexes[i] + 3, position + 1].Value = lines[i];

                                        if (lines[i].Contains("неравномерно"))
                                        {
                                            if (lines[i].Contains("несколько"))
                                            {
                                                if (lines[i].Contains("несколько неравномерно"))
                                                {
                                                    dataGridView1[9, position + 1].Value = "несколько неравномерно";
                                                }

                                                dataGridView1[9, position + 1].Value = "несколько";

                                                dataGridView1[10, position + 1].Value = "1";
                                            }
                                            else if (lines[i].Contains("незначительно"))
                                            {
                                                if (lines[i].Contains("незначительно неравномерно"))
                                                {
                                                    dataGridView1[9, position + 1].Value = "незначительно неравномерно";
                                                }

                                                dataGridView1[9, position + 1].Value = "незначительно";

                                                dataGridView1[10, position + 1].Value = "1";
                                            }
                                            else if (lines[i].Contains("умеренно"))
                                            {
                                                if (lines[i].Contains("умеренно неравномерно"))
                                                {
                                                    dataGridView1[9, position + 1].Value = "умеренно неравномерно";
                                                }

                                                dataGridView1[9, position + 1].Value = "умеренно";

                                                dataGridView1[10, position + 1].Value = "2";
                                            }
                                            else if (lines[i].Contains("значительно"))
                                            {
                                                if (lines[i].Contains("значительно неравномерно"))
                                                {
                                                    dataGridView1[9, position + 1].Value = "значительно неравномерно";
                                                }

                                                dataGridView1[9, position + 1].Value = "значительно";

                                                dataGridView1[10, position + 1].Value = "3";
                                            }
                                            else if (lines[i].Contains("резко"))
                                            {
                                                if (lines[i].Contains("резко неравномерно"))
                                                {
                                                    dataGridView1[9, position + 1].Value = "резко неравномерно";
                                                }

                                                dataGridView1[9, position + 1].Value = "резко";

                                                dataGridView1[10, position + 1].Value = "4";
                                            }
                                        }
                                        else
                                        {
                                            dataGridView1[9, position + 1].Value = "не сужена";

                                            dataGridView1[10, position + 1].Value = "0";
                                        }
                                    }

                                    if (indexes[i] >= 6 && indexes[i] <= 12)
                                    {
                                        dataGridView1[indexes[i] + 5, position + 1].Value = lines[i];
                                    }
                                }
                            }

                            position += 1;
                        }
                    }

                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    

                    this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                    this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    dataGridView1.Height = 780;
                    dataGridView1.Width = 1180;
                    this.ClientSize = new System.Drawing.Size(1200, 900);
                }
                catch (Exception ex)
                {
                    //tbPath.Text = "";
                    //cbSheet.Enabled = false;
                    //btnOpen.Enabled = true;
                    //MessageBox.Show(ex.Message);
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
            ExcelWorkSheet.Cells.Style.WrapText = true;

        }
    }
}
