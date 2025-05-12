using System;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordClientNou
{
    public partial class FormSelectieContract : Form
    {
        private readonly string connString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";
        private sealed class OfertaCombo
        {
            public int IdOferta;
            public string Text;
            public override string ToString() => Text;   // aşa apare în listă
        }

        public FormSelectieContract()
        {
            InitializeComponent();
            getClienti();
            comboBoxClienti.SelectedIndexChanged += (_, __) => getOferta();
        }

        //lsta cu clientii
        private void getClienti()
        {
            comboBoxClienti.Items.Clear();

            var con = new OleDbConnection(connString);
            con.Open();

            const string sql = "SELECT Nume, Prenume FROM Clienti ORDER BY Nume, Prenume";
            var cmd = new OleDbCommand(sql, con);
            var rdr = cmd.ExecuteReader();

            while (rdr.Read())
                comboBoxClienti.Items.Add($"{rdr.GetString(0)} {rdr.GetString(1)}");

            con.Close();
        }

        //oferte pentru clientul selectat
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
            con.Close();
        }

        private void buttonConfirma_Click(object sender, EventArgs e)
        {
            if (comboBoxClienti.SelectedIndex == -1 || comboBoxOferta.SelectedIndex == -1)
            {
                MessageBox.Show("Selectaţi un client şi o ofertă!");
                return;
            }

            string clientFull = comboBoxClienti.Text;
            var ofertaSel = (OfertaCombo)comboBoxOferta.SelectedItem;

            if (MessageBox.Show($"Generaţi contractul pentru {clientFull}?",
                "Confirmare", MessageBoxButtons.OKCancel)!= DialogResult.OK) return;

            string[] p = clientFull.Split(' ');
            string nume = p[0];
            string prenume = string.Join(" ", p.Skip(1));

            var db = new OleDbConnection(connString);
            db.Open();

            var cmdC = new OleDbCommand("SELECT * FROM Clienti WHERE Nume=? AND Prenume=?", db);
            cmdC.Parameters.AddWithValue("?", nume);
            cmdC.Parameters.AddWithValue("?", prenume);
            var rc = cmdC.ExecuteReader();
            if (!rc.Read()) { MessageBox.Show("Client inexistent."); db.Close(); return; }

            var cmdO = new OleDbCommand("SELECT * FROM Oferte WHERE IdOferta=?", db);
            cmdO.Parameters.AddWithValue("?", ofertaSel.IdOferta);
            var ro = cmdO.ExecuteReader();
            if (!ro.Read()) { MessageBox.Show("Oferta inexistentă."); db.Close(); return; }

            // generare word cu contract
            var word = new Word.Application { Visible = false };
            Word.Document doc = word.Documents.Add();

            // titlu
            Word.Paragraph titlu = doc.Content.Paragraphs.Add();
            titlu.Range.Text = "\tCONTRACT DE LEASING AUTO";
            titlu.Range.Font.Name = "Aptos";
            titlu.Range.Font.Size = 18;
            titlu.Range.Font.Bold = 1;
            titlu.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titlu.Format.SpaceAfter = 12;
            titlu.Range.InsertParagraphAfter();

            //parti contractante
            Word.Paragraph subtParti = doc.Content.Paragraphs.Add();
            subtParti.Range.Text = "\tArt. 1. PĂRȚILE CONTRACTANTE";
            subtParti.Range.Font.Name = "Aptos";
            subtParti.Range.Font.Size = 12;
            subtParti.Range.Font.Bold = 1;
            subtParti.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtParti.Format.SpaceAfter = 12;
            subtParti.Range.InsertParagraphAfter();

            DateTime effective = Convert.ToDateTime(rc["DataCreare"]);
            Word.Paragraph partile = doc.Content.Paragraphs.Add();
            partile.Range.Text =
                $"\tPrezentul contract este încheiat la data de {effective:dd.MM.yyyy} între societatea comercială COMPANIA ANA LEASING, " +
                $"cu sediul în Românie București Sector 3 str. Florilor nr. 25 tel. 0205551166 înmatriclată în Registrul Comerțului București " +
                $" sub nr. J23/493/2005 având contul nr. RO31INGB0000999908030304 la Banca ING și codul fiscal nr. 255 prin reprezentantul legal Anamaria-Florentina Țițeche " +
                $"având funcția director general cetățean român posesor al actului de identitate/pașaport nr. 6554488220039 (\"LOCATOR\") şi \n " +
                $"civilul {rc["Nume"]} {rc["Prenume"]} cetățean român posesor al actului de identitate/pașaport nr.  {rc["CNP"]}, " +
                $"domiciliat în {rc["Judet"]}, {rc["LocSec"]}, {rc["Adresa"]} (\"UTILIATOR\"). \n S-a încheiat prezentul contract.";
            partile.Range.Font.Name = "Times New Roman";
            partile.Range.Font.Size = 10;
            partile.Format.SpaceAfter = 12;
            partile.Range.InsertParagraphAfter();

            // subtitluy cu specificații vehicul
            Word.Paragraph subtVeh = doc.Content.Paragraphs.Add();
            subtVeh.Range.Text = "\tArt. 2. SPECIFICAȚII VEHICUL";
            subtVeh.Range.Font.Name = "Aptos";
            subtVeh.Range.Font.Size = 12;
            subtVeh.Range.Font.Bold = 1;
            subtVeh.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtVeh.Format.SpaceAfter = 3;
            subtVeh.Range.InsertParagraphAfter();

            // teabel cu detalii vehicul
            Word.Table tbl = doc.Tables.Add(doc.Content.Paragraphs.Add().Range, 6, 2);
            tbl.Columns[1].Width = 150; tbl.Columns[2].Width = 350;
            tbl.Cell(1, 1).Range.Text = "Marca"; tbl.Cell(1, 2).Range.Text = ro["Marca"].ToString();
            tbl.Cell(2, 1).Range.Text = "Model"; tbl.Cell(2, 2).Range.Text = ro["NumeMasina"].ToString();
            tbl.Cell(3, 1).Range.Text = "Tip carburant"; tbl.Cell(3, 2).Range.Text = ro["TipCarburant"].ToString();
            tbl.Cell(4, 1).Range.Text = "Culoare caroserie"; tbl.Cell(4, 2).Range.Text = ro["Culoare"].ToString();
            tbl.Cell(5, 1).Range.Text = "An fabricaţie"; tbl.Cell(5, 2).Range.Text = ro["AnFabricatie"].ToString();
            tbl.Cell(6, 1).Range.Text = "Kilometraj"; tbl.Cell(6, 2).Range.Text = $"{ro["Kilometraj"]} km";
            tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            tbl.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tbl.Borders.Enable = 1;

            //doua randuri goale dupa tabel
            Word.Range afterTbl = tbl.Range;
            afterTbl.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterTbl.InsertParagraphAfter();
            afterTbl.InsertParagraphAfter();

            //calcularea avansului
            string tipVenit = rc["Venit"].ToString();
            decimal avansPct = 0.1m;
            if (tipVenit == "Sub 5.000 RON")
                avansPct = 0.2m;
            else if (tipVenit == "5.000 RON - 10.000 RON")
                avansPct = 0.25m;
            else if (tipVenit == "10.000 RON - 20.000 RON")
                avansPct = 0.35m;
            else if (tipVenit == "Peste 20.000 RON")
                avansPct = 0.45m;

            decimal tarifZi = Convert.ToDecimal(ro["PretZi"]);
            int zile = Convert.ToInt32(ro["DurataMinima"]);
            decimal total = tarifZi * zile;

            // obiectul contractului
            Word.Paragraph subtObiect = doc.Content.Paragraphs.Add();
            subtObiect.Range.Text = "\tArt. 3. OBIECTUL CONTRACTULUI";
            subtObiect.Range.Font.Name = "Aptos";
            subtObiect.Range.Font.Size = 12;
            subtObiect.Range.Font.Bold = 1;
            subtObiect.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtObiect.Format.SpaceAfter = 12;
            subtObiect.Range.InsertParagraphAfter();

            Word.Paragraph txtObiect = doc.Content.Paragraphs.Add();
            txtObiect.Range.Text = $"\tLOCATORUL COMPANIA ANA LEASING închiriază UTILIZATORULUI {rc["Nume"]} {rc["Prenume"]} vehiculul specificat la art. 2," +
                $" iar UTILIZATORUL {rc["Nume"]} {rc["Prenume"]} se obligă să preia acest bun și să plătească " +
                $"tariful stabilit la art. 4. LOCATORUL COMPANIA ANA LEASING se obligă să transmită UTILIZATORULUI {rc["Nume"]} {rc["Prenume"]} dreptul de folosință " +
                $"și posesia unui automobil marca {ro["Marca"]} - {ro["NumeMasina"]}, având caracteristicile prezentate în art. 2." +
                $" Tariful include TVA. Tariful nu include costurile de întreţinere, reparaţii, asigurări, " +
                $"impozite, taxe şi alte cheltuieli legate de utilizarea vehiculului.";
            txtObiect.Range.Font.Name = "Times New Roman";
            txtObiect.Range.Font.Size = 10;
            txtObiect.Format.SpaceAfter = 12;
            txtObiect.Range.InsertParagraphAfter();

            // plata si conditii
            Word.Paragraph subtPlata = doc.Content.Paragraphs.Add();
            subtPlata.Range.Text = "\tArt. 4. PLATĂ ŞI CONDIŢII";
            subtPlata.Range.Font.Name = "Aptos";
            subtPlata.Range.Font.Size = 12;
            subtPlata.Range.Font.Bold = 1;
            subtPlata.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtPlata.Format.SpaceAfter = 12;
            subtPlata.Range.InsertParagraphAfter();

            Word.Paragraph txtPlata = doc.Content.Paragraphs.Add();
            txtPlata.Range.Text =
                $"\tTarif: {tarifZi:0} lei/zi (total {total:0} lei). Plata se face numerar sau prin transfer bancar la ING – RO31INGB0000999908030304. " +
                $"Depozit de garanţie: {(total * avansPct):0} lei. \n Transformarea obligațiilor de plată din valută în lei se face la cursul zilei afișat de BNR." +
                $" Diferențele de curs dintre cel utilizat în facturi și cel din ziua încasării de către LOCATOR mai mari de 2% se regularizează de părți prin achitarea diferenței către partea dezavantajată.";
            txtPlata.Range.Font.Name = "Times New Roman";
            txtPlata.Range.Font.Size = 10;
            txtPlata.Format.SpaceAfter = 12;
            txtPlata.Range.InsertParagraphAfter();

            // durata contractului
            DateTime end = effective.AddDays(zile);

            Word.Paragraph subtDur = doc.Content.Paragraphs.Add();
            subtDur.Range.Text = "\tArt. 5. DURATA CONTRACTULUI";
            subtDur.Range.Font.Name = "Aptos";
            subtDur.Range.Font.Size = 12;
            subtDur.Range.Font.Bold = 1;
            subtDur.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtDur.Format.SpaceAfter = 12;
            subtDur.Range.InsertParagraphAfter();

            Word.Paragraph txtDur = doc.Content.Paragraphs.Add();
            txtDur.Range.Text = $"\tContract valabil de la {effective:dd.MM.yyyy} până la {end:dd.MM.yyyy}.";
            txtDur.Range.Font.Name = "Times New Roman";
            txtDur.Range.Font.Size = 10;
            txtDur.Format.SpaceAfter = 12;
            txtDur.Range.InsertParagraphAfter();

            //termen livrare
            Word.Paragraph subtLiv = doc.Content.Paragraphs.Add();
            subtLiv.Range.Text = "\tArt. 6. TERMEN DE LIVRARE";
            subtLiv.Range.Font.Name = "Aptos";
            subtLiv.Range.Font.Size = 12;
            subtLiv.Range.Font.Bold = 1;
            subtLiv.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtLiv.Format.SpaceAfter = 12;
            subtLiv.Range.InsertParagraphAfter();

            Word.Paragraph txtLiv = doc.Content.Paragraphs.Add();
            txtLiv.Range.Text = $"\tUTILIZATORUL și-a ales obiectul leasingului în conformitate cu propriile necesități funcționale." +
                $"\r\nTermenul de livrare este de cel mult 15 zile de la încasarea de către LOCATOR a ratei 1.";
            txtLiv.Range.Font.Name = "Times New Roman";
            txtLiv.Range.Font.Size = 10;
            txtLiv.Format.SpaceAfter = 12;
            txtLiv.Range.InsertParagraphAfter();

            //predare - primire

            Word.Paragraph subtPred = doc.Content.Paragraphs.Add();
            subtPred.Range.Text = "\tArt. 7. PREDARE - PRIMIRE";
            subtPred.Range.Font.Name = "Aptos";
            subtPred.Range.Font.Size = 12;
            subtPred.Range.Font.Bold = 1;
            subtPred.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtPred.Format.SpaceAfter = 12;
            subtPred.Range.InsertParagraphAfter();

            Word.Paragraph txtPred = doc.Content.Paragraphs.Add();
            txtPred.Range.Text = $"\tUTILIZATORUL va prelua obiectul leasingului în termenul prevăzut la art. 6. data exactă a predării-primirii fiind comunicată în scris de către LOCATOR.\r\n" +
                $"Locul predării-primirii va fi București-Porsche Leasing Șos. PIPERA-TUNARI, CLĂDIREA PORSCHE 2 Voluntari\r\n\r\nPredarea-primirea obiectului leasingului se va face numai către reprezentantul autorizat al " +
                $"UTILIZATORULUI la prezentarea de către acesta a împuternicirii de reprezentare.\r\nPredarea-primirea obiectului leasingului se consemnează într-un proces-verbal care atestă condiția conformă cu caracteristicile " +
                $"tehnice și funcționalitățile obiectului leasingului.\r\nPerioada de leasing începe odată cu predarea-primirea. Dacă UTILIZATORUL întârzie preluarea obiectului leasingului, perioada de leasing va începe " +
                $"în ziua în care trebuia efectuată predarea-primirea. În cazul neefectuării predării-primirii din vina uneia din părți, aceasta va suporta penalități de 0,5%.";

            txtPred.Range.Font.Name = "Times New Roman";
            txtPred.Range.Font.Size = 10;
            txtPred.Format.SpaceAfter = 12;
            txtPred.Range.InsertParagraphAfter();

            // utilizarea vehiculului
            Word.Paragraph subtUtil = doc.Content.Paragraphs.Add();
            subtUtil.Range.Text = "\tArt. 8. UTILIZAREA VEHICULULUI";
            subtUtil.Range.Font.Name = "Aptos";
            subtUtil.Range.Font.Size = 12;
            subtUtil.Range.Font.Bold = 1;
            subtUtil.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtUtil.Format.SpaceAfter = 12;
            subtUtil.Range.InsertParagraphAfter();

            Word.Paragraph txtUtil = doc.Content.Paragraphs.Add();
            txtUtil.Range.Text = "\tVehiculul poate fi folosit exclusiv în scop personal. Subînchirierea este interzisă.LOCATORUL rămâne proprietarul obiectului leasingului pe toată durata contractului de leasing.\r\n" +
                "Riscurile și beneficiile aferente dreptului de proprietate trec asupra UTILIZATORULUI din momentul predării-primirii obiectului leasingului.";
            txtUtil.Range.Font.Name = "Times New Roman";
            txtUtil.Range.Font.Size = 10;
            txtUtil.Format.SpaceAfter = 12;
            txtUtil.Range.InsertParagraphAfter();

            //intarziere plata

            Word.Paragraph subtIntarziere = doc.Content.Paragraphs.Add();
            subtIntarziere.Range.Text = "\tArt. 9. ÎNTÂRZIEREA PLĂȚILOR";
            subtIntarziere.Range.Font.Name = "Aptos";
            subtIntarziere.Range.Font.Size = 12;
            subtIntarziere.Range.Font.Bold = 1;
            subtIntarziere.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtIntarziere.Format.SpaceAfter = 12;
            subtIntarziere.Range.InsertParagraphAfter();

            Word.Paragraph txtIntarziere = doc.Content.Paragraphs.Add();
            txtIntarziere.Range.Text = $"\tÎn cazul întârzierii plății, UTILIZATORUL va plăti penalități de 0,5% pe zi de întârziere. " +
                $"După 30 zile de întârziere, LOCATORUL poate rezilia contractul și solicita returnarea vehiculului.";
            txtIntarziere.Range.Font.Name = "Times New Roman";
            txtIntarziere.Range.Font.Size = 10;
            txtIntarziere.Format.SpaceAfter = 12;
            txtIntarziere.Range.InsertParagraphAfter();

            // întârziere returnare
            Word.Paragraph subtLate = doc.Content.Paragraphs.Add();
            subtLate.Range.Text = "\tArt. 10. ÎNTÂRZIERE LA RETURNARE";
            subtLate.Range.Font.Name = "Aptos";
            subtLate.Range.Font.Size = 12;
            subtLate.Range.Font.Bold = 1;
            subtLate.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtLate.Format.SpaceAfter = 12;
            subtLate.Range.InsertParagraphAfter();

            Word.Paragraph txtLate = doc.Content.Paragraphs.Add();
            txtLate.Range.Text = $"\tPenalizare: {tarifZi:0} lei pentru fiecare zi de întârziere a returnării vehiculului.";
            txtLate.Range.Font.Name = "Times New Roman";
            txtLate.Range.Font.Size = 10;
            txtLate.Format.SpaceAfter = 12;
            txtLate.Range.InsertParagraphAfter();

            // garanții și responsabilități
            Word.Paragraph subtGar = doc.Content.Paragraphs.Add();
            subtGar.Range.Text = "\tArt. 11. GARANŢII ŞI RESPONSABILITĂŢI";
            subtGar.Range.Font.Name = "Aptos";
            subtGar.Range.Font.Size = 12;
            subtGar.Range.Font.Bold = 1;
            subtGar.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtGar.Format.SpaceAfter = 12;
            subtGar.Range.InsertParagraphAfter();

            Word.Paragraph gr1 = doc.Content.Paragraphs.Add();
            gr1.Range.Text = "\t- Locatorul garantează proprietatea vehiculului.";
            gr1.Range.Font.Name = "Times New Roman";
            gr1.Range.Font.Size = 10;
            gr1.Format.SpaceAfter = 6;
            gr1.Range.InsertParagraphAfter();

            Word.Paragraph gr2 = doc.Content.Paragraphs.Add();
            gr2.Range.Text = "\t- Vehiculul este asigurat CASCO.";
            gr2.Range.Font.Name = "Times New Roman";
            gr2.Range.Font.Size = 10;
            gr2.Format.SpaceAfter = 6;
            gr2.Range.InsertParagraphAfter();

            Word.Paragraph gr3 = doc.Content.Paragraphs.Add();
            gr3.Range.Text = "\t- Locatarul suportă costul combustibilului.";
            gr3.Range.Font.Name = "Times New Roman";
            gr3.Range.Font.Size = 10;
            gr3.Format.SpaceAfter = 12;
            gr3.Range.InsertParagraphAfter();

            // răspundere
            Word.Paragraph subtResp = doc.Content.Paragraphs.Add();
            subtResp.Range.Text = "\tArt. 12. RĂSPUNDERE";
            subtResp.Range.Font.Name = "Aptos";
            subtResp.Range.Font.Size = 12;
            subtResp.Range.Font.Bold = 1;
            subtResp.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtResp.Format.SpaceAfter = 12;
            subtResp.Range.InsertParagraphAfter();

            Word.Paragraph txtResp = doc.Content.Paragraphs.Add();
            txtResp.Range.Text = "\tUTILIZATORUL își asumă riscul în cazurile deteriorării, daunei totale sau furtului neacoperite de polița de asigurare din vina sa, " +
                "cât și în cazul uzurii premature a obiectului leasingului cu urmare a utilizării și întreținerii neconforme dispozițiilor art. 8.\r\nÎn caz de daună totală sau furt, " +
                "LOCATORUL va returna UTILIZATORULUI, din suma totală a plăților efectuate de către acesta în cadrul contractului de leasing și contravaloarea despăgubirii " +
                "de la societatea de asigurări, diferența ce depășește valoarea contractuală de leasing și însoțește încheierii unui nou contract de leasing.";
            txtResp.Range.Font.Name = "Times New Roman";
            txtResp.Range.Font.Size = 10;
            txtResp.Format.SpaceAfter = 12;
            txtResp.Range.InsertParagraphAfter();

            // litigii
            Word.Paragraph subtLit = doc.Content.Paragraphs.Add();
            subtLit.Range.Text = "\tArt. 13. DISPOZIȚII FINALE";
            subtLit.Range.Font.Name = "Aptos";
            subtLit.Range.Font.Size = 12;
            subtLit.Range.Font.Bold = 1;
            subtLit.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtLit.Format.SpaceAfter = 12;
            subtLit.Range.InsertParagraphAfter();

            Word.Paragraph txtLit = doc.Content.Paragraphs.Add();
            txtLit.Range.Text = "\tLitigiile apărute în legătură cu formarea, validitatea, interpretarea sau executarea obligațiilor contractuale se soluționează pe cale amiabilă sau arbitrală" +
                ". În lipsa unei clauze compromisorii, litigiile vor fi soluționate la instanțele judecătorești competente.\r\nPrezentul contract, inclusiv anexele, a fost încheiat în două exemplare, câte unul pentru fiecare parte.";
            txtLit.Range.Font.Name = "Times New Roman";
            txtLit.Range.Font.Size = 10;
            txtLit.Format.SpaceAfter = 12;
            txtLit.Range.InsertParagraphAfter();

            // legea aplicabilă
            Word.Paragraph subtLege = doc.Content.Paragraphs.Add();
            subtLege.Range.Text = "\tArt. 14. LEGEA APLICABILĂ";
            subtLege.Range.Font.Name = "Aptos";
            subtLege.Range.Font.Size = 12;
            subtLege.Range.Font.Bold = 1;
            subtLege.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtLege.Format.SpaceAfter = 12;
            subtLege.Range.InsertParagraphAfter();

            Word.Paragraph txtLege = doc.Content.Paragraphs.Add();
            txtLege.Range.Text = "\tPrezentul contract este guvernat de legea română.";
            txtLege.Range.Font.Name = "Times New Roman";
            txtLege.Range.Font.Size = 10;
            txtLege.Format.SpaceAfter = 12;
            txtLege.Range.InsertParagraphAfter();

            // contract complet
            Word.Paragraph subtComp = doc.Content.Paragraphs.Add();
            subtComp.Range.Text = "\tArt. 15. CONTRACT COMPLET";
            subtComp.Range.Font.Name = "Aptos";
            subtComp.Range.Font.Size = 12;
            subtComp.Range.Font.Bold = 1;
            subtComp.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtComp.Format.SpaceAfter = 12;
            subtComp.Range.InsertParagraphAfter();

            Word.Paragraph txtComp = doc.Content.Paragraphs.Add();
            txtComp.Range.Text = "\tContractul reprezintă acordul integral al părţilor şi înlocuieşte orice înţelegere anterioară. Modificarea contractului se face de comun acord prin act adițional semnat de ambele părți.";
            txtComp.Range.Font.Name = "Times New Roman";
            txtComp.Range.Font.Size = 10;
            txtComp.Format.SpaceAfter = 12;
            txtComp.Range.InsertParagraphAfter();

            // divizibilitate
            Word.Paragraph subtDiv = doc.Content.Paragraphs.Add();
            subtDiv.Range.Text = "\tArt. 16. DIVIZIBILITATE";
            subtDiv.Range.Font.Name = "Aptos";
            subtDiv.Range.Font.Size = 12;
            subtDiv.Range.Font.Bold = 1;
            subtDiv.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            subtDiv.Format.SpaceAfter = 12;
            subtDiv.Range.InsertParagraphAfter();

            Word.Paragraph txtDiv = doc.Content.Paragraphs.Add();
            txtDiv.Range.Text = "\tDacă o clauză devine invalidă, restul contractului rămâne în vigoare.";
            txtDiv.Range.Font.Name = "Times New Roman";
            txtDiv.Range.Font.Size = 10;
            txtDiv.Format.SpaceAfter = 12;
            txtDiv.Range.InsertParagraphAfter();

            // finalul contractului
            Word.Paragraph final = doc.Content.Paragraphs.Add();
            final.Range.Text = "\tPrezentul contract, inclusiv anexele, a fost încheiat în două exemplare, câte unul pentru fiecare parte.";
            final.Range.Font.Name = "Times New Roman";
            final.Range.Font.Size = 10;
            final.Range.Font.Italic = 1;
            final.Format.SpaceAfter = 12;
            final.Range.InsertParagraphAfter();


            Word.Paragraph logoPara = doc.Content.Paragraphs.Add();
            Word.Range logoRange = logoPara.Range;
            logoRange.InlineShapes.AddPicture(@"C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\Resources\Logo.png");
            logoRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
           
            logoRange.InsertParagraphAfter();
            Word.Paragraph sign1 = doc.Content.Paragraphs.Add();
            sign1.Range.Text = "\tLOCATOR ________________________               UTILIZATOR ________________________";
            sign1.Range.Font.Name = "Times New Roman";
            sign1.Range.Font.Size = 10;
            sign1.Format.SpaceAfter = 0;
            sign1.Range.InsertParagraphAfter();

            Word.Paragraph sign2 = doc.Content.Paragraphs.Add();
            sign2.Range.Text = $"\tDATA: {DateTime.Now:dd.MM.yyyy}";
            sign2.Range.Font.Name = "Times New Roman";
            sign2.Range.Font.Size = 10;
            sign2.Format.SpaceAfter = 0;
            sign2.Range.InsertParagraphAfter();

            // colorare text care vine din baza de date
            string[] dinamice =
            {
        rc["Nume"].ToString(), rc["Prenume"].ToString(), rc["CNP"].ToString(),
        rc["Judet"].ToString(), rc["LocSec"].ToString(), rc["Adresa"].ToString(),
        ro["Marca"].ToString(), ro["NumeMasina"].ToString(),
        ro["TipCarburant"].ToString(), ro["Culoare"].ToString(),
        ro["AnFabricatie"].ToString(), ro["Kilometraj"].ToString()
    };

            foreach (var val in dinamice)
            {
                Word.Range rng = doc.Content;
                if (rng.Find.Execute(val))
                {
                    rng.Font.Color = Word.WdColor.wdColorViolet;
                    rng.Font.Italic = 1;
                }
            }

            //salvare document
            string marca = ro["Marca"].ToString();
            string model = ro["NumeMasina"].ToString();
            string dataStr = DateTime.Now.ToString("dd-MM-yyyy");
            string fileName = $"{nume}_{prenume}_{marca}_{model}_{dataStr}.docx".Replace(' ', '_');

            string folder = @"C:\Users\anama\source\repos\ProiectIPLA\Contracts";
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            string fullPath = Path.Combine(folder, fileName);

            doc.SaveAs2(fullPath);
            word.Visible = true;

            MessageBox.Show(new Form { TopMost = true }, $"Contract salvat:\n{fullPath}");
            db.Close();
        }


        private void buttonRenunta_Click(object sender, EventArgs e) => Close();
    }
}
