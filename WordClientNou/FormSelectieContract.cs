using System;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordClientNou
{
    public partial class FormSelectieContract : Form
    {
        private readonly string connString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

        /*──────── helper pentru ComboBox ofertă ────────*/
        private sealed class OfertaCombo
        {
            public int IdOferta;
            public string Text;
            public override string ToString() => Text;   // aşa rămâne afişat
        }

        public FormSelectieContract()
        {
            InitializeComponent();
            getClienti();
            comboBoxClienti.SelectedIndexChanged += (_, __) => getOferta();
        }

        /*──────── clienţi ────────*/
        private void getClienti()
        {
            comboBoxClienti.Items.Clear();
            var con = new OleDbConnection(connString);
            con.Open();

            const string q = "SELECT Nume, Prenume FROM Clienti ORDER BY Nume, Prenume";
            var cmd = new OleDbCommand(q, con);
            var r = cmd.ExecuteReader();
            while (r.Read())
            {
                comboBoxClienti.Items.Add($"{r.GetString(0)} {r.GetString(1)}");
            }
        }

        /*──────── oferte pentru clientul selectat ────────*/
        private void getOferta()
        {
            comboBoxOferta.Items.Clear();
            if (comboBoxClienti.SelectedIndex == -1) return;

            string[] parts = comboBoxClienti.Text.Split(' ');
            string nume = parts[0];
            string prenume = string.Join(" ", parts.Skip(1));

            var con = new OleDbConnection(connString);
            con.Open();

            const string sql =
              @"SELECT o.IdOferta, o.Marca, o.NumeMasina, o.Kilometraj
                  FROM  Oferte AS o
                  INNER JOIN Clienti AS c ON c.Id = o.IdClient
                  WHERE c.Nume = ? AND c.Prenume = ?";

            var cmd = new OleDbCommand(sql, con);
            cmd.Parameters.AddWithValue("?", nume);
            cmd.Parameters.AddWithValue("?", prenume);

            var rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                var of = new OfertaCombo
                {
                    IdOferta = rdr.GetInt32(0),
                    Text = $"{rdr["Marca"]} {rdr["NumeMasina"]} {rdr["Kilometraj"]} km"
                };
                comboBoxOferta.Items.Add(of);
            }
        }

        /*──────── Confirmă ────────*/
        private void buttonConfirma_Click(object sender, EventArgs e)
        {
            if (comboBoxClienti.SelectedIndex == -1 || comboBoxOferta.SelectedIndex == -1)
            {
                MessageBox.Show("Selectaţi un client şi o ofertă!"); return;
            }

            string clientFull = comboBoxClienti.Text;
            var ofertaSel = (OfertaCombo)comboBoxOferta.SelectedItem;

            if (MessageBox.Show($"Generaţi contractul pentru {clientFull}?",
                                "Confirmare", MessageBoxButtons.OKCancel)
                != DialogResult.OK) return;

            /* 1️⃣ despărţim nume / prenume */
            string[] p = clientFull.Split(' ');
            string nume = p[0];
            string prenume = string.Join(" ", p.Skip(1));

            /* 2️⃣ citim date din Access */
            var db = new OleDbConnection(connString);
            db.Open();

            var cmdC = new OleDbCommand("SELECT * FROM Clienti WHERE Nume=? AND Prenume=?", db);
            cmdC.Parameters.AddWithValue("?", nume);
            cmdC.Parameters.AddWithValue("?", prenume);
            var rc = cmdC.ExecuteReader();
            if (!rc.Read()) { MessageBox.Show("Client inexistent."); return; }

            var cmdO = new OleDbCommand("SELECT * FROM Oferte WHERE IdOferta=?", db);
            cmdO.Parameters.AddWithValue("?", ofertaSel.IdOferta);
            var ro = cmdO.ExecuteReader();
            if (!ro.Read()) { MessageBox.Show("Oferta inexistentă."); return; }

            /* 3️⃣ Word Interop – document gol */
            var word = new Word.Application { Visible = false };
            Word.Document doc = word.Documents.Add();

            void Para(string text, int size = 10, bool bold = false,
                      Word.WdParagraphAlignment align = Word.WdParagraphAlignment.wdAlignParagraphJustify,
                      int spaceAfter = 6)
            {
                var pgh = doc.Content.Paragraphs.Add();
                pgh.Range.Text = text;
                pgh.Range.Font.Name = "Times New Roman";
                pgh.Range.Font.Size = size;
                pgh.Range.Font.Bold = bold ? 1 : 0;
                pgh.Format.Alignment = align;
                pgh.Format.SpaceAfter = spaceAfter;
                pgh.Range.InsertParagraphAfter();
            }

            /* PARTIES */
            Para("VEHICLE LEASE AGREEMENT", 18, true,
                 Word.WdParagraphAlignment.wdAlignParagraphCenter, 12);

            DateTime effective = Convert.ToDateTime(rc["DataCreare"]);
            Para($"This Vehicle Lease Agreement (\"Agreement\") is entered into on {effective:dd.MM.yyyy} "
               + $"(“Effective Date”) by and between COMPANIA ANA LEASING, "
               + $"București (\"Lessor\") and {rc["Nume"]} {rc["Prenume"]}, CNP {rc["CNP"]}, "
               + $"domiciliat în {rc["Judet"]}, {rc["LocSec"]}, {rc["Adresa"]} "
               + $"(\"Lessee\").", 10, false);

            /* VEHICLE SPECIFICATIONS –  table 2 col x 7 rânduri */
            Para("VEHICLE SPECIFICATIONS", 12, true, Word.WdParagraphAlignment.wdAlignParagraphCenter , 3);

            var tbl = doc.Tables.Add(doc.Content.Paragraphs.Add().Range, 6, 2);
            tbl.Columns[1].Width = 150; tbl.Columns[2].Width = 150;
            void Cell(int r, string key, string val)
            { tbl.Cell(r, 1).Range.Text = key; tbl.Cell(r, 2).Range.Text = val; }

            Cell(1, "Marca", ro["Marca"].ToString());
            Cell(2, "Model", ro["NumeMasina"].ToString());
            Cell(3, "Tip de Carburant", ro["TipCarburant"].ToString());
            Cell(4, "Body Color", ro["Culoare"].ToString());
            Cell(5, "Year", ro["AnFabricatie"].ToString());
            Cell(6, "Mileage", $"{ro["Kilometraj"]} km");

            tbl.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            tbl.Borders.Enable = 1;

            //swicth case in functie de venit se schimba avansul
            string tipVenit = rc["Venit"].ToString();
            decimal avans = 0.1m;
            if (tipVenit == "Sub 5.000 RON")
            {
                avans = 0.2m;
            }
            else if (tipVenit == "5.000 RON - 10.000 RON")
            {
                avans = 0.25m;
            }
            else if (tipVenit == "10.000 RON - 20.000 RON")
            {
                avans = 0.35m;
            }
            else if (tipVenit == "Peste 20.000 RON")
            {
                avans = 0.45m;
            }

            /* PAYMENT AND TERMS */
            Para("PAYMENT AND PAYMENT TERMS", 12, true);
            decimal price = Convert.ToDecimal(ro["PretZi"]);
            int zile = Convert.ToInt32(ro["DurataMinima"]);
            decimal pretTotal = price * zile;
            Para($"Leasing price: {price:0} lei/zi reprezentând: {pretTotal:0}; plata în numerar sau transfer bancar catre banca ING cu ibanu RO31INGB0000999908030304."
               + $"  Depozit de garanţie (non-refundable): {pretTotal * avans:0} lei.");

            /* AGREEMENT DURATION */
            Para("AGREEMENT DURATION", 12, true);
            
            DateTime end = effective.AddDays(zile);
            Para($"Perioada începe la {effective:dd.MM.yyyy} şi se termină la {end:dd.MM.yyyy}.");

            /* VEHICLE USAGE */
            Para("VEHICLE USAGE", 12, true);
            Para("Vehiculul va fi folosit exclusiv în scop personal; este interzisă subînchirierea.");

            /* LATE RETURN */
            Para("LATE RETURN", 12, true);
            Para($"Taxă de întârziere: {price:0} lei pentru fiecare zi peste perioada contractată.");

            /* WARRANTIES & GUARANTEES – doar subtitluri + bullets generice */
            Para("WARRANTIES AND GUARANTEES", 12, true);
            Para("- Lessor garantează proprietatea şi lipsa oricăror sarcini.\n"
               + "- Vehiculul este asigurat parţial.\n"
               + "- Lessee se obligă să plătească integral costurile cu combustibilul.", 10, false);

            /* LIABILITY + DISPUTE + GOVERNING LAW */
            Para("LIABILITY", 12, true);
            Para("Părţile nu răspund pentru daune indirecte, cu excepţia celor rezultate din neglijenţă.");

            Para("DISPUTE RESOLUTION", 12, true);
            Para("Orice dispută se soluţionează pe cale amiabilă; în caz contrar, prin arbitraj CCI.");

            Para("GOVERNING LAW", 12, true);
            Para("Prezentul acord este guvernat de legea română.");

            /* ENTIRE AGREEMENT + SEVERABILITY */
            Para("ENTIRE AGREEMENT", 12, true);
            Para("Acordul conţine înţelegerea completă a Părţilor şi înlocuieşte orice acord anterior.");

            Para("SEVERABILITY", 12, true);
            Para("Dacă o clauză devine inaplicabilă, restul contractului rămâne valabil.");

            /* SIGNATURES */
            Para("\n\nLESSOR ________________________    LESSEE ________________________", 10, false,
                 Word.WdParagraphAlignment.wdAlignParagraphLeft, 0);
            Para($"DATE: {DateTime.Now:dd.MM.yyyy}", 10, false,
                 Word.WdParagraphAlignment.wdAlignParagraphLeft, 0);

            /* 4️⃣ salvează în folderul proiectului */
            string marca = ro["Marca"].ToString();
            string model = ro["NumeMasina"].ToString();
            string data = DateTime.Now.ToString("dd-MM-yyyy");
            string fileName = $"{nume}_{prenume}_{marca}_{model}_{data}.docx"
                               .Replace(' ', '_');

            string folder = "C:\\Users\\anama\\source\\repos\\ProiectIPLA\\Contracte";
            string fullPath = Path.Combine(folder, fileName);

            doc.SaveAs2(fullPath);
            word.Visible = true;

            MessageBox.Show(new Form { TopMost = true }, $"Contract salvat:\n{fullPath}");
        }


        private void buttonRenunta_Click(object sender, EventArgs e) => Close();
    }
}
