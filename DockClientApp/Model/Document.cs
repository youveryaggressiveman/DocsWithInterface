using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DockClientApp.Model
{
    public class Document
    {
        public string Post { get; set; }
        public string MainFio { get; set; }
        public string Group { get; set; }
        public string Period { get; set; }
        public string DateOfPublication { get; set; }
        public string NameOfPublication { get; set; }
        public string Place { get; set; }
        public string NameOfDirection { get; set; }
        public string WorkDirection { get; set; }
        public string Authors { get; set; }
        public string Year { get; set; }
    }
}
