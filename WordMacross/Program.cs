using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordMacross;

namespace WordMacross
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var wordMacross = new WordHeaders();
            wordMacross.AddHeaderRange();
        }
    }
}
