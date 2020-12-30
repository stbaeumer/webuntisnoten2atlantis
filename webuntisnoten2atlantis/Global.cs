// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System.Collections.Generic;

namespace webuntisnoten2atlantis
{
    internal static class Global
    {
        public static List<string> Output { get; internal set; }

        internal static void PrintMessage(string message)
        {
            if (message.ToLower().Contains("zu"))
            {
                Output.Add("");
            }
            Output.Add("/* " + message.PadRight(97) + " */");
            if (message.ToLower().Contains("zu"))
            {
                Output.Add("");
            }
        }
    }
}