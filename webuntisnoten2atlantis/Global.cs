// Published under the terms of GPLv3 Stefan Bäumer 2019.

using System;
using System.Collections.Generic;

namespace webuntisnoten2atlantis
{
    internal static class Global
    {
        public static List<string> Output { get; internal set; }

        internal static void PrintMessage(string message)
        {
            Output.Add("");
            Output.Add("/* " + (message.PadRight(60)).PadLeft(78) + " */");
            Output.Add("");
        }
    }
}