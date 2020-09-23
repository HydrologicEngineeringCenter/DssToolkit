using Hec.Dss;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DSSExcel
{
    public class ReviewPageVM
    {
        public UserControl PreviousPage;
        public RecordType currentRecordType;
        private bool tsPathGenerated = false;
        private bool pdPathGenerated = false;
        private DssPath tsPath = new DssPath();
        private DssPath pdPath = new DssPath();

        private string aPart = "";
        private string bPart = "";
        private string cPart = "";
        private string dPart = "";
        private string ePart = "";
        private string fPart = "";
        public string Apart
        {
            get { return aPart; }
        }
        public string Bpart
        {
            get { return bPart; }
        }
        public string Cpart
        {
            get { return cPart; }
        }
        public string Dpart
        {
            get { return dPart; }
        }
        public string Epart
        {
            get { return ePart; }
        }
        public string Fpart
        {
            get { return fPart; }
        }

        public ReviewPageVM()
        {

        }
    }
}
