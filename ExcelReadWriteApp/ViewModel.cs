

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Data.OleDb;
using System.Data;
using Prism.Events;
using Spire.Xls;
using System.Windows.Controls;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Diagnostics;
using ClassLibrary1;

namespace ExcelReadWriteApp
{
    public class ViewModel : INotifyPropertyChanged
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        

        private string? _filepath;
        public string? filepath
        {
            get => _filepath;
            set
            {
                _filepath = value;
                OnpropertyChanged("filepath");
            }
        }

        private DataTable? _dataTable;
        public DataTable? dataTable
        {
            get { return _dataTable; }
            set
            {
                _dataTable = value;
                OnpropertyChanged("dataTable");
            }
        }

        private DataRow? _row;
        public DataRow? row
        {
            get
            {
                return _row;
            }
            set
            {
                _row = value;
                OnpropertyChanged("row");
            }
        }

        
        public Employee? Employee { get; set; }
        private IEventAggregator? eventAggregator;

        public ICommand BrowseCmd { get; set; }
        public ICommand ImportCmd { get; set; }
        public ICommand AddCmd { get; set; }
        public ICommand EnterCmd { get; set; }
        public ICommand ExportCmd { get; set; }
        //OleDbConnection _connection;
        public Class1 obj;
        public ViewModel(IEventAggregator eventAggregator)
        {
            this.eventAggregator = eventAggregator;
            this.eventAggregator.GetEvent<EmployeeTransferEvent>().Subscribe((_employee) => enter(this.Employee = _employee));
            BrowseCmd = new DelegateCommand(browse, () => canExecuteBrowse);
            ImportCmd = new DelegateCommand(import, () => canExecuteImport);
            AddCmd = new DelegateCommand(addRec);
            ExportCmd = new DelegateCommand(export);
            //EnterCmd = new DelegateCommand(enter);
        }

        public void addRec()
        {
            Add addRecordForm = new Add();
            addRecordForm.ShowDialog();


        }
        public void enter(object e)
        {

            row = dataTable.NewRow();
            Employee = (Employee)e;
            row["Employee ID"] = Employee.EmployeeId;
            row["Employee Name"] = Employee.EmployeeName;
            row["Salary"] = Employee.Salary;
            row["Department"] = Employee.Department;
            dataTable.Rows.Add(row);

        }
        public bool canExecuteBrowse => true;
        public bool canExecuteImport
        {
            get
            {
                if (string.IsNullOrEmpty(_filepath)) return false;
                return true;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnpropertyChanged(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }
        public void browse()
        {
            log.Info("Hello logging world! Browsing a file..");

            OpenFileDialog OpenFile = new OpenFileDialog();
            OpenFile.Title = "Select File";
            OpenFile.Filter = "Excel Sheet (*.xls)|*.xls|All files(*.*)|*.*";
            bool? response = OpenFile.ShowDialog();
            if (response == true)
            {
                filepath = OpenFile.FileName;
            }

        }
        public void import()
        {
            string excelConnectionString = @"provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            try
            {  
               obj = new Class1();
               dataTable=obj.read(excelConnectionString);

                //log.Info("Establishing connection for import..");
                //throw new Exception();
                /* _connection = new OleDbConnection(excelConnectionString);
                 _connection.Open();

                 OleDbDataAdapter adapter = new OleDbDataAdapter("Select * from [Sheet1$]", _connection);
                 DataSet ds = new DataSet();

                 adapter.Fill(ds);
                 dataTable = ds.Tables[0];
                 if (dataTable == null || dataTable.Columns.Count == 0)
                 {
                     throw new Exception("Blank file selected! Choose another file.");
                 }*/
                
                log.Info("Import Successful..");
            }

            catch (Exception exception)
            {
                log.Error(exception);

                MessageBox.Show(exception.ToString());
            }
        }


        public void export()
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                excelworkBook = excel.Workbooks.Add(Type.Missing);

                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;

                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        if (rowcount == 3)
                        {
                            excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
                        }

                        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                    }


                }


                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "Excel File (*.xls)|*.xls|Show All Files (*.*)|*.*";
                bool? result = save.ShowDialog();


                if (result == true)
                {

                    filepath = save.FileName;
                }
                excelworkBook.SaveAs(filepath); ;
                excelworkBook.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }


    }

}

