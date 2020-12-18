// Published under the terms of GPLv3 Stefan Bäumer 2019.

namespace webuntisnoten2atlantis
{
    public class Abwesenheit
    {
        public int NotenkopfId { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public double StundenAbwesend { get; internal set; }
        public double StundenAbwesendUnentschuldigt { get; internal set; }
        public int StudentId { get; internal set; }
        public string Zeugnisart { get; internal set; }
        public string HzJz { get; internal set; }
    }
}