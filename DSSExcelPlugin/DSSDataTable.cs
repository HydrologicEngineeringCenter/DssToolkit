using Hec.Dss;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSSExcelPlugin
{
    class DSSDataTable
    {
        public DataTable table = new DataTable();
        public RecordType Type
        {
            get
            {
                return _recordType;
            }
        }
        private RecordType _recordType;

        public DSSDataTable(RecordType type)
        {
            _recordType = type;
        }
    }
}
