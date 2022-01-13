using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeHelper.Models
{
    public class Data
    {
        public string Name { get; set; }
        public ObservableCollection<string> Values { get; set; }
    }
}
