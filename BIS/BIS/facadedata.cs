using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BIS
{
    public class facadedata
    {
        public string BIN;
        public string Num;
        public string Strt;
        public string Boro;
        public string Zip;
        public string NumStory;
        public List<cycle> Cycles = new List<cycle>();
        //public string Cycle;
        //public string CurrentStatus;

        public facadedata() { }
    }
}
