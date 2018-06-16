using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace PresentationCommandInterface
{
    class Program
    {
        static void Main(string[] args)
        {
            PresentationOperation a = new PresentationOperation();
            a.GotoImage();
        }
    }
}
