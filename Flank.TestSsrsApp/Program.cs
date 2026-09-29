using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Flank.TestSsrsApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            new Flank.Ssrs.SsrsClient()
                .TestConnectionAsync()
                .GetAwaiter()
                .GetResult();

            Console.ReadLine();
        }
    }
}
