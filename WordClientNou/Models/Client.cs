using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordClientNou.Models
{
    public class Client
    {
        public int Id { get; set; }
        public string Nume { get; set; }
        public string Prenume { get; set; }
        public string Telefon { get; set; }
        public string Judet { get; set; }
        public string CNP { get; set; }
        public string Email { get; set; }
        public string Adresa { get; set; }
        public string LocSec { get; set; }
        public string Venit { get; set; }
        public string DataCreare { get; set; }

        public Client(string nume, string prenume, string telefon, string judet, string cnp, string email, string adresa, string locSec, string venit)
        {
            Nume = nume;
            Prenume = prenume;
            Telefon = telefon;
            Judet = judet;
            CNP = cnp;
            Email = email;
            Adresa = adresa;
            LocSec = locSec;
            Venit = venit;
            DataCreare = DateTime.Now.ToString("yyyy-MM-dd");
        }
    }
}
