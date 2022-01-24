// Published under the terms of GPLv3 Stefan Bäumer 2021.

using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace webuntisnoten2atlantis
{
    internal class Lehrer
    {
        public Lehrer()
        {
        }

        public string Kuerzel { get; internal set; }
        public string Mail { get; internal set; }

        
    }
}