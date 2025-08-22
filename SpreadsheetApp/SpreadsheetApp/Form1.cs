using System;
using System.Windows.Forms;
using Simulator;

namespace SpreadsheetApp
{
    /// <summary>
    /// The main form of the spreadsheet GUI application.
    /// It allows viewing and editing a SharableSpreadSheet instance using a DataGridView.
    /// </summary>
    public partial class Form1 : Form
    {
        // Backend spreadsheet object, supports thread-safe operations

        private SharableSpreadSheet sheet;

        /// <summary>
        /// Initializes the form and its components.
        /// Creates a new spreadsheet with default size (10x5),
        /// clears its contents, and populates the UI grid.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            sheet = new SharableSpreadSheet(10, 5);  // Instantiate spreadsheet with 10 rows and 5 columns
            ClearSheet();  // Ensure all cells are empty initially
            InitGrid();   // Load grid layout and contents

        }


        /// <summary>
        /// Clears all spreadsheet cells by setting their content to an empty string.
        /// This is a full-table iteration used on initialization only.
        /// </summary>
        private void ClearSheet()
        {
            var size = sheet.getSize();
            int rows = size.Item1;
            int cols = size.Item2;

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.setCell(r, c, "");  // Set each cell to empty string
                }
            }
        }

        /// <summary>
        /// Converts a zero-based column index into an Excel-style column name.
        /// For example: 0 → A, 1 → B, ..., 25 → Z, 26 → AA, etc.
        /// </summary>
        /// <param name="columnIndex">Zero-based column index</param>
        /// <returns>Excel-style column name</returns>
        private string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>
        /// Initializes the DataGridView based on the spreadsheet's current size.
        /// Sets headers, adds rows, and populates with current cell values.
        /// </summary>
        private void InitGrid()
        {
            var size = sheet.getSize();
            int rows = size.Item1;
            int cols = size.Item2;

            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // Add column headers (Excel-style)
            for (int c = 0; c < cols; c++)
            {
                string header = GetExcelColumnName(c);
                dataGridView1.Columns.Add(header, header);
            }
            // Add row headers and row entries
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.RowHeadersWidth = 60;
            dataGridView1.Rows.Add(rows);
            for (int r = 0; r < rows; r++)
            {
                dataGridView1.Rows[r].HeaderCell.Value = (r + 1).ToString();  // Display row numbers starting from 1
            }

            RefreshGrid();  // Populate with actual cell values
        }


        /// <summary>
        /// Updates the DataGridView cell values based on the current spreadsheet state.
        /// Should be called after data is loaded or modified.
        /// </summary>
        private void RefreshGrid()
        {
            var size = sheet.getSize();
            for (int r = 0; r < size.Item1; r++)
            {
                for (int c = 0; c < size.Item2; c++)
                {
                    dataGridView1[c, r].Value = sheet.getCell(r, c);  // Fetch value from spreadsheet
                }
            }
        }


        /// <summary>
        /// Called when a cell edit is committed in the DataGridView.
        /// Updates the underlying spreadsheet model with the new value.
        /// </summary>
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            string newValue = dataGridView1[col, row].Value?.ToString() ?? "";
            sheet.setCell(row, col, newValue);    // Reflect UI change in backend
        }

        /// <summary>
        /// Opens a CSV file using an OpenFileDialog and loads it into the spreadsheet.
        /// Reinitializes the DataGridView with the new dimensions and values.
        /// </summary>
        private void btnLoad_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "CSV files|*.csv";   // Restrict to .csv files
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    sheet.load(dlg.FileName);    // Load into backend spreadsheet
                    InitGrid();                 // Rebuild the UI grid
                }
            }
        }

        /// <summary>
        /// Saves the current spreadsheet to a CSV file using a SaveFileDialog.
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "CSV files|*.csv";   // Restrict to .csv files
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    sheet.save(dlg.FileName);   // Save backend spreadsheet
                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
