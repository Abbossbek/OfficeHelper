using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeHelper.Models
{
    public class Column:DispatchedBindableBase
    {
        private string name;
        private bool isChecked;

        public string Name { 
            get => name;
            set => SetProperty(ref name, value);
        }
        public bool IsChecked
        {
            get => isChecked;
            set => SetProperty(ref isChecked, value);
        }
    }
}
