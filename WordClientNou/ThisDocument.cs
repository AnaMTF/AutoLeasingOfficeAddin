using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using WordClientNou.Models;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordClientNou
{
    public partial class ThisDocument
    {
        Button btnAdaugareClient, btnUltimulClient;
        //string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
        //        System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IPLAproiect.accdb");
        string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            btnAdaugareClient = new Button();
            btnAdaugareClient.Text = "Adauga Client";
            btnAdaugareClient.Click += new EventHandler(btnAdaugareClient_Click);

            btnUltimulClient = new Button();
            btnUltimulClient.Text = "Ultimul Client";
            btnUltimulClient.Click += new EventHandler(btnUltimulClient_Click);

            Globals.ThisDocument.ActionsPane.Controls.Add(btnAdaugareClient);
            Globals.ThisDocument.ActionsPane.Controls.Add(btnUltimulClient);
        }

        private void btnAdaugareClient_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(rtcNume.Text) || string.IsNullOrEmpty(rtcPrenume.Text) || string.IsNullOrEmpty(rtcTelefon.Text)
                || string.IsNullOrEmpty(rtcJudet.Text) || string.IsNullOrEmpty(rtcCNP.Text) || string.IsNullOrEmpty(rtcEmail.Text)
                || string.IsNullOrEmpty(rtcAdresa.Text) || string.IsNullOrEmpty(rtcLocSec.Text) || string.IsNullOrEmpty(dlcVenit.Text))
            {
                MessageBox.Show("Vă rugăm să completați toate câmpurile.");
                return;
            }
            else if (rtcCNP.Text.Length != 13)
            {
                MessageBox.Show("CNP-ul trebuie să aibă 13 caractere.");
                return;
            }
            else if (rtcTelefon.Text.Length != 10)
            {
                MessageBox.Show("Numărul de telefon trebuie să aibă 10 caractere.");
                return;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(rtcEmail.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Adresa de email nu este validă.");
                return;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(rtcCNP.Text, @"^\d+$"))
            {
                MessageBox.Show("CNP-ul trebuie să conțină doar cifre.");
                return;
            }
            else
            {
                OleDbConnection conn = new OleDbConnection(connString);
                conn.Open();
                Client client = new Client(rtcNume.Text, rtcPrenume.Text, rtcTelefon.Text, rtcJudet.Text, rtcCNP.Text, rtcEmail.Text, rtcAdresa.Text, rtcLocSec.Text, dlcVenit.Text);
                string maxIdQuery = "SELECT MAX(Id) FROM Clienti";
                OleDbCommand maxIdCommand = new OleDbCommand(maxIdQuery, conn);
                object result = maxIdCommand.ExecuteScalar();
                int maxId = (result == DBNull.Value) ? 0 : System.Convert.ToInt32(result);
                client.Id = maxId + 1;
                string insertQuery = "INSERT INTO Clienti (Nume, Prenume, Telefon, CNP, Email, Judet, LocSec, Adresa, Venit, DataCreare) " +
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                OleDbCommand insertCommand = new OleDbCommand(insertQuery, conn);
                insertCommand.Parameters.AddWithValue("?", client.Nume);
                insertCommand.Parameters.AddWithValue("?", client.Prenume);
                insertCommand.Parameters.AddWithValue("?", client.Telefon);
                insertCommand.Parameters.AddWithValue("?", client.CNP);
                insertCommand.Parameters.AddWithValue("?", client.Email);
                insertCommand.Parameters.AddWithValue("?", client.Judet);
                insertCommand.Parameters.AddWithValue("?", client.LocSec);
                insertCommand.Parameters.AddWithValue("?", client.Adresa);
                insertCommand.Parameters.AddWithValue("?", client.Venit);
                insertCommand.Parameters.AddWithValue("?", client.DataCreare);

                insertCommand.ExecuteNonQuery();
                MessageBox.Show("Client adăugat cu succes!");

                // Clear the text boxes after adding the client
                rtcNume.Text = "";
                rtcPrenume.Text = "";
                rtcTelefon.Text = "";
                rtcJudet.Text = "";
                rtcCNP.Text = "";
                rtcEmail.Text = "";
                rtcAdresa.Text = "";
                rtcLocSec.Text = "";
                conn.Close();
            }

        }

        private void btnUltimulClient_Click(object sender, EventArgs e)
        {
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IPLAproiect.accdb");

            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();
                string query = "SELECT TOP 1 * FROM Clienti ORDER BY Id DESC";

                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string mesaj = $"ID: {reader["Id"]}\n" +
                                   $"Nume: {reader["Nume"]}\n" +
                                   $"Prenume: {reader["Prenume"]}\n" +
                                   $"Telefon: {reader["Telefon"]}\n" +
                                   $"CNP: {reader["CNP"]}\n" +
                                   $"Email: {reader["Email"]}\n" +
                                   $"Județ: {reader["Judet"]}\n" +
                                   $"Loc Secundar: {reader["LocSec"]}\n" +
                                   $"Adresa: {reader["Adresa"]}\n" +
                                   $"Venit: {reader["Venit"]}\n" +
                                   $"Data creării: {reader["DataCreare"]}";

                    MessageBox.Show(mesaj, "Ultimul client adăugat");
                }
                else
                {
                    MessageBox.Show("Baza de date este goală.");
                }

                reader.Close();
                conn.Close();
            }
        }


        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
