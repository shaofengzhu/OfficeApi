using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeExtension
{
    public interface ITrackedObject
    {
        string _ReferenceId
        {
            get;
            set;
        }

        void _KeepReference();

    }
}
