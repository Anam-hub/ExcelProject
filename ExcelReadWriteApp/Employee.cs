using Prism.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadWriteApp
{
    public class Employee: PubSubEvent<Employee>, INotifyPropertyChanged ,IDataErrorInfo
    {

        public int? EmployeeId { get; set; }
        private string employeeName = "";

        

        public string EmployeeName { 
            get { return employeeName; }
            set
            {
                employeeName = value;
                OnpropertyChanged("EmployeeName");
            }
        }
        
            public string Error
            { 
            get
            {
                return string.Empty;
            }
        }

        public string this[string columnName] 
        {
            get
            {
                string result = String.Empty;
                if (columnName == "EmployeeName")
                {
                    if (EmployeeName.Length < 2||EmployeeName.Length > 12)
                    {
                        result = "Name should be between range 2-12";
                    }
                }

                return result;
            }
        }


        public double? Salary { get; set; }
        public string? Department { get; set; }

        

        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnpropertyChanged(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }

    }
}
