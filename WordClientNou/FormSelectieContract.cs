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

namespace WordClientNou
{
    public partial class FormSelectieContract : Form
    {
        string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";
        public FormSelectieContract()
        {
            InitializeComponent();
            getClienti();
            getOferta();
        }

        private void getClienti()
        {
            using (var connection = new OleDbConnection(connString))
            {
                connection.Open();
                string query = "SELECT * FROM Clienti";
                using (var command = new OleDbCommand(query, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string nume = reader["Nume"].ToString();
                            string prenume = reader["Prenume"].ToString();
                            comboBoxClienti.Items.Add($"{nume} {prenume}");
                        }
                    }
                }
                connection.Close();
            }
        }

        private void getOferta()
        {
            using (var connection = new OleDbConnection(connString))
            {
                connection.Open();
                if (comboBoxClienti.SelectedIndex != -1)
                {
                    string client = comboBoxClienti.SelectedItem.ToString();
                    string[] numePrenume = client.Split(' ');
                    string nume = numePrenume[0];
                    string prenume = numePrenume[1];
                    string query = $"SELECT * FROM Oferte WHERE Nume='{nume}' AND Prenume='{prenume}'";
                    using (var command = new OleDbCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string oferta = reader["Marca"].ToString() + " " + reader["NumeMasina"].ToString() + " " + reader["Kilometraj"].ToString() + "km";
                                comboBoxOferta.Items.Add(oferta);
                            }
                        }
                    }
                }
                connection.Close();
            }

        }

        private void buttonConfirma_Click(object sender, EventArgs e)
        {

        }

        private void buttonRenunta_Click(object sender, EventArgs e)
        {

        }
    }
}