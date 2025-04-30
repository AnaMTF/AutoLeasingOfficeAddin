using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordClientNou
{
    public partial class FormSelectieContract : Form
    {
        private readonly string connString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

        public FormSelectieContract()
        {
            InitializeComponent();

            getClienti();

            /* când utilizatorul schimbă clientul ⇒ reîncarcă ofertele */
            comboBoxClienti.SelectedIndexChanged += (_, __) => getOferta();
        }

        /*──────────────────────── clienți ────────────────────────*/
        private void getClienti()
        {
            comboBoxClienti.Items.Clear();

            var connection = new OleDbConnection(connString);
            connection.Open();

            const string query = "SELECT Nume, Prenume FROM Clienti ORDER BY Nume, Prenume";
            var command = new OleDbCommand(query, connection);
            var reader = command.ExecuteReader();

            while (reader.Read())
            {
                string nume = reader.GetString(0);
                string prenume = reader.GetString(1);
                comboBoxClienti.Items.Add($"{nume} {prenume}");
            }
        }

        /*──────────────────────── oferte pentru clientul selectat ────────────────────────*/
        private void getOferta()
        {
            comboBoxOferta.Items.Clear();

            if (comboBoxClienti.SelectedIndex == -1) return;

            // desparte "Nume Prenume" -> pot exista mai multe prenume, de aceea Join
            string[] parts = comboBoxClienti.SelectedItem.ToString().Split(' ');
            string nume = parts[0];
            string prenume = string.Join(" ", parts.Skip(1));

            var connection = new OleDbConnection(connString);
            connection.Open();

            const string sql =
                @"SELECT o.Marca, o.NumeMasina, o.Kilometraj
                FROM  Oferte  AS o
                INNER JOIN Clienti AS c ON c.Id = o.IdClient
                WHERE c.Nume = ?  AND c.Prenume = ?";

            var cmd = new OleDbCommand(sql, connection);
            cmd.Parameters.AddWithValue("?", nume);
            cmd.Parameters.AddWithValue("?", prenume);

            var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string oferta = $"{rdr["Marca"]} {rdr["NumeMasina"]} {rdr["Kilometraj"]}km";
                comboBoxOferta.Items.Add(oferta);
            }
        }

        /*──────────────────────── butoane ────────────────────────*/
        private void buttonConfirma_Click(object sender, EventArgs e)
        {
            Word.Application appWord = new Word.Application();
            Word.Document doc = appWord.Documents.Add();

        }

        private void buttonRenunta_Click(object sender, EventArgs e)
        {
            Close();
        }
    }

}