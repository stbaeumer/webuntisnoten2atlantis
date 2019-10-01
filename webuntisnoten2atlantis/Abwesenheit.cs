namespace webuntisnoten2atlantis
{
    public class Abwesenheit
    {
        public int StudentId { get; internal set; }
        public string Name { get; internal set; }
        public string Klasse { get; internal set; }
        public double StundenAbwesend { get; internal set; }
        public double StundenAbwesendUnentschuldigt { get; internal set; }
    }
}