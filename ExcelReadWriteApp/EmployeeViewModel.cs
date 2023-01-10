using Microsoft.Win32;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelReadWriteApp
{
    public class EmployeeViewModel : INotifyPropertyChanged
    {
        private Employee? _employee;
        private IEventAggregator _eventAggregator;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Employee? Employee
        {
            get { return _employee; }
            set
            {
                _employee = value;
                OnpropertyChanged("Employee");
            }
        }



        public ICommand SaveCmd
        {
            get; set;
        }

        public EmployeeViewModel(IEventAggregator eventAggregator)
        {
            this._eventAggregator = eventAggregator;
            Employee = new Employee();
            SaveCmd = new DelegateCommand(save);

        }

        public void save()
        {
            _eventAggregator.GetEvent<EmployeeTransferEvent>().Publish(_employee);
            initLog4Net();
            log.Info("Your record is added in the list");
            msg = "Data recorded successfully!";
        }
        private string? _msg;
        public string? msg
        {
            get { return _msg; }
            set { _msg = value; OnpropertyChanged("msg"); }
        }





        public event PropertyChangedEventHandler? PropertyChanged;
        public void OnpropertyChanged(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }


        public void initLog4Net()
        {
            try
            {
                var hierarchy = (log4net.Repository.Hierarchy.Hierarchy)log4net.LogManager.GetRepository();
                hierarchy.Configured = true;
                var rollingAppender = new log4net.Appender.RollingFileAppender
                {
                    File = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + "\\" +
                  "LogFiles\\" + Employee.EmployeeName+"_" + DateTime.Now.ToString("yyyyMMddTHHmm") + ".log",
                    AppendToFile = true,
                    LockingModel = new log4net.Appender.FileAppender.MinimalLock(),
                    Layout = new log4net.Layout.PatternLayout("%date [%thread] %level %logger - %message%newline")
                };
                var traceAppender = new log4net.Appender.TraceAppender()
                {
                    Layout = new log4net.Layout.PatternLayout("%date [%thread] %level %logger - %message%newline")
                };
                hierarchy.Root.AddAppender(rollingAppender);
                hierarchy.Root.AddAppender(traceAppender);
                rollingAppender.ActivateOptions();
                hierarchy.Root.Level = log4net.Core.Level.All;
            }
            catch (global::System.Exception ex)
            {
                new Exception(ex.Message);
            }
        }
    }

    public class EmployeeTransferEvent:PubSubEvent<Employee>
    {
    }
}
