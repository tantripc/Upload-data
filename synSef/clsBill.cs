using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace synSef
{
    public class clsBill
    {
        public clsBill()
        {

        }
        string index, se, atr_where, atr_when, atr_who;

        public string Index
        {
            get { return index; }
            set { index = value; }
        }

        public string Atr_who
        {
            get { return atr_who; }
            set { atr_who = value; }
        }

        public string Atr_when
        {
            get { return atr_when; }
            set { atr_when = value; }
        }

        public string Atr_where
        {
            get { return atr_where; }
            set { atr_where = value; }
        }

        public string SE
        {
            get { return se; }
            set { se = value; }
        }
    }
}
