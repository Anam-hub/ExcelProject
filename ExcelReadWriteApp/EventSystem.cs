
using ExcelReadWriteApp;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Prism
{
    public class EventSystem : PubSubEvent<Employee>
    {
        public sealed class Event
        {
            private IEventAggregator eventAggregator;
            internal IEventAggregator EventAggregator
            {
                get
                {
                    if (eventAggregator == null)
                    {
                        eventAggregator = new EventAggregator();
                    }

                    return eventAggregator;
                }
            }

            private static readonly Event eventInstance = new Event();
            internal static Event EventInstance
            {
                get
                {
                    return eventInstance;
                }
            }

           
            private Event() { }
               

        }


    }
}
