// Published under the terms of GPLv3 Stefan Bäumer 2020.

namespace webuntisnoten2atlantis
{
    public class Zuordnung
    {
        public string Quellklasse { get; internal set; }
        public string Quellfach { get; internal set; }
        public string Zielklasse { get; internal set; }
        public string Zielfach { get; internal set; }

        public Zuordnung(string quellklasse, string quellfach)
        {
            Quellklasse = quellklasse;
            Quellfach = quellfach;
        }

        public Zuordnung(string quellklasse, string quellfach, string zielfach)
        {
            Quellklasse = quellklasse;
            Quellfach = quellfach;
            Zielfach = zielfach;
        }

        public Zuordnung()
        {
        }
    }
}