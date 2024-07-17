using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.VisualBasic.FileIO;
using Org.BouncyCastle.Bcpg.Sig;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace WinformCsvFile
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "All Files|*.*";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filepath = openFileDialog.FileName;
                    textBox1.Text = filepath;
                    if (filepath.ToLower().EndsWith(".pdf"))
                    {
                        DataFromPDF(filepath);
                      
                    }
                    else if (filepath.ToLower().EndsWith(".xls") || filepath.ToLower().EndsWith(".xlsx"))
                    {
                        Data(filepath, ".xlsx");
                    }
                    else if (filepath.ToLower().EndsWith(".csv"))
                    {
                        LoadCSV(filepath);
                    }
                    else if (filepath.ToLower().EndsWith(".txt"))
                    {
                        LoadTextFile(filepath);
                    }
                }
            }
        }
        // Excel
        public void Data(string filepath, string ext)
        {
            string connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filepath};Extended Properties='Excel 12.0;'";

            OleDbConnection conn = new OleDbConnection(connString);

            conn.Open();

            // Retrieves the schema table of excel
            DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);


            string sheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();

            string query = $"SELECT * FROM [{sheetName}]";

            //data retrieved by the adapter from the Excel sheet. 

            OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            dataGridView1.DataSource = dataTable;

        }

        

        public DataTable ConvertCSVtoDataTable(string strFilePath)
        {

            DataTable dtCsv = new DataTable();
            string Fulltext;

            using (StreamReader sr = new StreamReader(strFilePath))
            {
                while (!sr.EndOfStream)
                {
                    //read the full file text
                    Fulltext = sr.ReadToEnd().ToString();
                    //split the full file text into rows
                    string[] rows = Fulltext.Split('\n');
                    for (int i = 0; i < rows.Count() - 1; i++)
                    {
                        //split each row with comma to get the individual values
                        string[] rowValues = rows[i].Split(',');
                        {
                            if (i == 0)
                            {
                                for (int j = 0; j < rowValues.Count(); j++)
                                {
                                    //add headers
                                    dtCsv.Columns.Add(rowValues[j]);
                                }
                            }
                            else
                            {
                                DataRow dr = dtCsv.NewRow();
                                for (int k = 0; k < rowValues.Count(); k++)
                                {
                                    dr[k] = rowValues[k].ToString();
                                }
                                //add other rows
                                dtCsv.Rows.Add(dr);
                            }
                        }
                    }
                }
            }
            return dtCsv;
        }


        private void DataFromPDF(string filepath)
        {
            try
            {
                DataTable dataTable = new DataTable();

                // dataTable = ConvertCSVtoDataTable(filepath);

                using (PdfReader reader = new PdfReader(filepath))
                {
                    // Iterate through each page of the PDF
                    for (int page = 1; page <= reader.NumberOfPages; page++)
                    {
                        // Extract text from the page
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(reader, page, strategy);

                        // Split text into lines
                        string[] lines = currentText.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                        // Assuming the first line contains column headers (modify as needed)
                        if (page == 1 && dataTable.Columns.Count == 0)
                        {
                            string[] headers = lines[0].Split(' ');
                            foreach (string header in headers)
                            {
                                dataTable.Columns.Add(new DataColumn(header));
                            }
                        }

                        try
                        {
                            // Add data rows
                            for (int i = 0; i < lines.Length; i++) // Start from 1 to skip headers
                            {
                                string[] fields = lines[i].Split(' ');

                                DataRow toInsert = dataTable.NewRow();

                                // insert in the desired place
                                dataTable.Rows.InsertAt(toInsert, i );
                                for (int k = 0; k < fields.Length; k++)
                                {
                                    try
                                    {
                                        //dataTable.Rows.Add(fields[k]);
                                        dataTable.Rows[i][k] = fields[k];
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                    
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                        }

                       
                    }
                }

                // Display in DataGridView
                dataGridView1.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading PDF file: {ex.Message}");
            }
        }


        
        private void LoadTextFile(string filepath)
        {
            try
            {
                string[] lines = File.ReadAllLines(filepath);

                // Assuming the text file contains tabular data separated by tabs or commas
                // You can adjust this based on your text file's format
                string[] headers = lines[0].Split('\t'); // Assuming tabs separate columns

                // Clear existing data and columns
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();

                // Add columns to DataGridView
                foreach (string header in headers)
                {
                    dataGridView1.Columns.Add(header, header);
                }

                // Add rows to DataGridView
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] fields = lines[i].Split('\t'); // Assuming tabs separate columns
                    dataGridView1.Rows.Add(fields);
                }

                MessageBox.Show("Text File Loaded Successfully", "Info");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading text file: {ex.Message}", "Error");
            }
        }


        //CSV


        private void LoadCSV(string filepath)
        {
            try
            {
                // TextFieldParser is reading and parsing delimited text files, such as CSV files.

                using (TextFieldParser parser = new TextFieldParser(filepath))
                {
                 //   parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    DataTable dataTable = new DataTable();

                    if (!parser.EndOfData)
                    {
                        //start read data 
                        string[] fields = parser.ReadFields();
                        foreach (string field in fields)
                        {
                          // dataTable.Columns.Add(new DataColumn(field));
                            dataTable.Columns.Add(field);
                        }
                    }

                    // Add rows
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        dataTable.Rows.Add(fields);
                    }

                    // Display in DataGridView
                    dataGridView1.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading CSV file: {ex.Message}");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        //file to convert to excel
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectAll();
            DataObject copydata = dataGridView1.GetClipboardContent();
            if (copydata != null) Clipboard.SetDataObject(copydata);
            Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            xlapp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook xlWbook;
            Microsoft.Office.Interop.Excel.Worksheet xlsheet;
            object miseddata = System.Reflection.Missing.Value;
            xlWbook = xlapp.Workbooks.Add(miseddata);

            xlsheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWbook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)xlsheet.Cells[1, 1];
            xlr.Select();

            xlsheet.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

        }

        //Database to fetch gridview 
        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection sqlConnection = new SqlConnection("Server=DESKTOP-0OPK228\\SQLEXPRESS;Database=Testdb;User Id=sa;Password=admin@123;");
            SqlCommand cmd = new SqlCommand("select * from t_student", sqlConnection);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sqlDataAdapter.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void CSV_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToCSV(dataGridView1, "DataGridViewExport.csv");
        }
        private void ExportDataGridViewToCSV(DataGridView dataGridView, string filename)
        {
            if (dataGridView.Rows.Count > 0)
            {
                StringBuilder csvContent = new StringBuilder();

                // Add column headers
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    csvContent.Append((dataGridView.Columns[i].HeaderText));
                    if (i < dataGridView.Columns.Count - 1)
                    {
                        csvContent.Append(",");
                    }
                } 
                csvContent.AppendLine();

                // Add rows
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        csvContent.Append((row.Cells[i].Value?.ToString()));
                        if (i < dataGridView.Columns.Count - 1)
                        {
                            csvContent.Append(",");
                        }
                    }
                    csvContent.AppendLine();
                }

                // Write to file
                try
                {
                    using (StreamWriter sw = new StreamWriter(filename, false, Encoding.UTF8))
                    {
                        sw.Write(csvContent.ToString());
                    }
                    MessageBox.Show("CSV Data Exported Successfully", "Info");

                    // Open the file after ensuring it's closed
                    using (Process process = new Process())
                    {
                        process.StartInfo.FileName = filename;
                        process.StartInfo.UseShellExecute = true;
                        process.Start();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting data: {ex.Message}", "Error");
                }
            }

            else
            {
                MessageBox.Show("No Record To Export", "Info");
            }
        }






        private void Pdf_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToPDF(dataGridView1, "DataGridViewExport.pdf");
        }
        private void ExportDataGridViewToPDF(DataGridView dataGridView, string filename)
         {
            if (dataGridView.Rows.Count > 0)
            {
                Document pdfDocument = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                PdfPTable pdfTable = new PdfPTable(dataGridView.Columns.Count);
              //  pdfTable.DefaultCell.Padding = 3;
                //pdfTable.WidthPercentage = 100;
                //pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                // Add headers
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                 //   cell.BackgroundColor = new BaseColor(240, 240, 240);
                    pdfTable.AddCell(cell);
                }

                // Add rows
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        pdfTable.AddCell(cell.Value?.ToString());
                    }
                }

                using (FileStream stream = new FileStream(filename, FileMode.Create))
                {
                    PdfWriter.GetInstance(pdfDocument, stream);
                    pdfDocument.Open();
                    pdfDocument.Add(pdfTable);
                    pdfDocument.Close();
                    stream.Close();
                }

                MessageBox.Show("Pdf Data Exported Successfully", "Info");
                // Automatically open the PDF file
                System.Diagnostics.Process.Start(filename);
            }
            else
            {
                MessageBox.Show("No Record To Export", "Info");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToTextFile(dataGridView1, "DataGridViewExport.txt");
        }
        private void ExportDataGridViewToTextFile(DataGridView dataGridView, string filename)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filename))
                {
                    // Write column headers
                    for (int i = 0; i < dataGridView.Columns.Count; i++)
                    {
                        writer.Write(dataGridView.Columns[i].HeaderText);
                        if (i < dataGridView.Columns.Count - 1)
                        {
                            writer.Write("\t"); // Tab separator
                        }
                    }
                    writer.WriteLine();

                    // Write rows
                    foreach (DataGridViewRow row in dataGridView.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            for (int i = 0; i < dataGridView.Columns.Count; i++)
                            {
                                writer.Write(row.Cells[i].Value?.ToString());
                                if (i < dataGridView.Columns.Count - 1)
                                {
                                    writer.Write("\t"); // Tab separator
                                   
                                }
                            }
                            writer.WriteLine();
                        }
                    }
                }

                MessageBox.Show(" Text Data Exported Successfully", "Info");
                // Automatically open the text file
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filename) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data to text file: {ex.Message}", "Error");
            }
        }
    }
}
