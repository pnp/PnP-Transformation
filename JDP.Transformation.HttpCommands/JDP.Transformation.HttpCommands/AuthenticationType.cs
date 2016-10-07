using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDP.Transformation.HttpCommands
{
    /// <summary>
    /// Used to differentiate the authentication scheme details during execution
    /// </summary>
    public enum AuthenticationType
    {
        DefaultCredentials,
        NetworkCredentials,
        Office365
    }
}

