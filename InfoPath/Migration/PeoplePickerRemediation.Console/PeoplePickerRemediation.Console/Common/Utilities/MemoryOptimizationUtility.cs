using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeoplePickerRemediation.Console.Common.Utilities
{
    public static class MemoryOptimizationUtility
    {
        public static long GetMemoryUsage { get { return (System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)); } }
    }
}
