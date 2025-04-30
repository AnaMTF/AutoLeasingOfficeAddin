using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordClientNou.Models
{
    public class Oferta
    {
        public int IdOferta { get; set; }
        public string NumeMasina { get; set; }      // „X5 30d xDrive”
        public string Marca { get; set; }      // „BMW”
        public string TipCarburant { get; set; }      // „Diesel”
        public int DurataMinima { get; set; }      // zile
        public decimal PretZi { get; set; }      // tarif / zi
        public int Kilometraj { get; set; }      // km / zi incluși (null = nelimitat)
        public int NrLocuri { get; set; }      // 2‑9
        public string StatusOferta { get; set; }      // Disponibil, Indisponibil
        public string NumePoza { get; set; }      // numele imaginii
        public string Culoare { get; set; }      // culoarea mașinii
        public int AnFabricatie { get; set; }      // anul fabricatiei

        public Oferta() { }
        public Oferta(string numeMasina, string marca, string tipCarburant,
            int durataMinima, decimal pretZi, int kilometraj, int nrLocuri,
            string statusOferta, string culoare, int anFabricatie)
        {
            NumeMasina = numeMasina;
            Marca = marca;
            TipCarburant = tipCarburant;
            DurataMinima = durataMinima;
            PretZi = pretZi;
            Kilometraj = kilometraj;
            NrLocuri = nrLocuri;
            StatusOferta = statusOferta;
            Culoare = culoare;
            AnFabricatie = anFabricatie;
        }

        public override string ToString() =>
            $"{Marca} {NumeMasina} ({PretZi:C} / zi)";
    }

}
