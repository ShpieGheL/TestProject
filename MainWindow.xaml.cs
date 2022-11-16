using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Data;
using System.Data.OleDb;
using System.IO;
using ExcelLibrary;

namespace TestProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string path;
        OleDbConnection connection = new();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Find(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new();
            openFileDialog.Title = "Выберите файл";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xls)|*.xls|Text files (*.csv)|*.csv|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
                path = openFileDialog.FileName;
            else
                return;
            DataSet ds = new();
            if (System.IO.Path.GetExtension(path) == ".xlsx")
            {
                connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={@path};Extended Properties='Excel 8.0;HDR=Yes'");
                connection.Open();
                var dtable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                OleDbDataAdapter oleda = new();
                oleda.SelectCommand = new OleDbCommand("SELECT * FROM [" + dtable.Rows[0]["TABLE_NAME"].ToString() + "]", connection);
                oleda.Fill(ds);
            }
            else if (System.IO.Path.GetExtension(path) == ".csv")
            {
                DataTable dt = new();
                string[] lines = File.ReadAllLines(path);
                if (lines.Length > 0)
                {
                    string firstLine = lines[0];
                    string[] headerLabels;
                    headerLabels = firstLine.Split(';');
                    foreach (string headerWord in headerLabels)
                    {
                        dt.Columns.Add(new DataColumn(headerWord));
                    }
                    for (int i = 1; i < lines.Length; i++)
                    {
                        string[] dataWords;
                        dataWords = lines[i].Split(';');
                        DataRow dr = dt.NewRow();
                        int columnIndex = 0;
                        foreach (string headerWord in headerLabels)
                        {
                            dr[headerWord] = dataWords[columnIndex++];
                        }
                        dt.Rows.Add(dr);
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    ds.Tables.Add(dt);
                }
            }
            Table.ItemsSource = ds.Tables[0].DefaultView;
        }

        private void Table_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Table.SelectedIndex == -1)
                return;
            list.Items.Clear();
            chart.DataContext = null;
            double[] data = new double[4];
            data[0] = Convert.ToDouble(((DataRowView)Table.SelectedItem)["Distance"].ToString().Replace('.',','));
            data[1] = Convert.ToDouble(((DataRowView)Table.SelectedItem)["Angle"].ToString().Replace('.', ','));
            data[2] = Convert.ToDouble(((DataRowView)Table.SelectedItem)["Width"].ToString().Replace('.', ','));
            data[3] = Convert.ToDouble(((DataRowView)Table.SelectedItem)["Hegth"].ToString().Replace('.', ','));
            string def = ((DataRowView)Table.SelectedItem)["IsDefect"].ToString();
            list.Items.Add("Name: " + ((DataRowView)Table.SelectedItem)["Name"].ToString());
            list.Items.Add($"Distance: {data[0]}");
            list.Items.Add($"Angle: {data[1]}");
            list.Items.Add($"Width: {data[2]}");
            list.Items.Add($"Heigth: {data[3]}");
            list.Items.Add($"IsDefect: " + ((DataRowView)Table.SelectedItem)["IsDefect"].ToString());
            double[] x = new double[] { data[0] - data[2] / 2, data[0] - data[2] / 2, data[0] + data[2] / 2, data[0] + data[2] / 2};
            double[] y = new double[] { data[1] - data[3] / 2, data[1] + data[3] / 2, data[1] + data[3] / 2, data[1] - data[3] / 2 };
            List<KeyValuePair<double, double>> valueList = new();
            for (int i = 0; i < 4; i++)
                valueList.Add(new KeyValuePair<double, double>(x[i], y[i]));
            chart.DataContext = valueList;
        }

        private string saveFileDialog ()
        {
            SaveFileDialog saveFileDialog = new();
            saveFileDialog.Title = "Выберите назначение";
            saveFileDialog.Filter = "All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
                return saveFileDialog.FileName;
            return null;
        }

        private void ExCSV(object sender, RoutedEventArgs e)
        {
            string path = saveFileDialog();
            if (path == null)
                return;
            IEnumerable<string> columnNames = ((DataView)Table.ItemsSource).ToTable().Columns.Cast<DataColumn>().
                                  Select(column => column.ColumnName);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Join(";", columnNames));
            foreach (DataRow row in ((DataView)Table.ItemsSource).ToTable().Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                ToArray();
                sb.AppendLine(string.Join(";", fields));
            }
            File.WriteAllText(path + ".csv", sb.ToString());
        }

        private void ExXLSX(object sender, RoutedEventArgs e)
        {
            string path = saveFileDialog();
            if (path == null)
                return;
            Export(".xlsx");
        }
        private void ExXLS(object sender, RoutedEventArgs e)
        {
            string path = saveFileDialog();
            if (path == null)
                return;
            Export(".xls");
        }

        private void Export (string extension)
        {
            DataSet ds = new();
            DataTable dt = ((DataView)Table.ItemsSource).Table.Copy();
            ds.Tables.Add(dt);
            DataSetHelper.CreateWorkbook(path + extension, ds);
        }
    }
}
