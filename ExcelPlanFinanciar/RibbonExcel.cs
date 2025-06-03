using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPlanFinanciar
{
    public partial class RibbonExcel
    {
        private Excel.Application excelApp;
        private const string ConnString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";
        private static readonly Dictionary<string, double> VenitCategorieDobanda = new Dictionary<string, double>
        {
            { "Sub 5.000 RON",          0.11 },   // 11 % anual
            { "5.000 RON - 10.000 RON", 0.15 },
            { "10.000 RON - 20.000 RON",0.25 },
            { "Peste 20.000 RON",       0.35 }
        };
        

        private void RibbonExcel_Load(object sender, RibbonUIEventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
        }

        private void btnVizualizareGrafic_Click(object sender, RibbonControlEventArgs e)
        {
            var f = new FormGraficSelector();
            if (f.ShowDialog() != DialogResult.OK) return;

            Excel.Worksheet ws;
            try { ws = (Excel.Worksheet)excelApp.Worksheets["Plan Financiar Plata"]; }
            catch { MessageBox.Show("Întâi generaţi planul financiar."); return; }

            foreach (Excel.ChartObject c in ws.ChartObjects())
                if (c.Name.StartsWith("Grafic"))
                    c.Delete();

            int lastRow = ws.Cells[ws.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row;
            if (lastRow < 3) { MessageBox.Show("Tabelul pare gol."); return; }

            Excel.Range rngLuna = ws.Range[$"A2:A{lastRow}"];
            Excel.Range rngSold = ws.Range[$"E2:E{lastRow}"];
            Excel.Range rngDobanda = ws.Range[$"C2:C{lastRow}"];
            Excel.Range rngPrincipal = ws.Range[$"B2:B{lastRow}"];

            switch (f.AlegereGrafic)
            {
                case FormGraficSelector.TipGrafic.Linie:
                    var g1 = ws.ChartObjects().Add(300, 40, 480, 280);
                    g1.Name = "GraficLinie";
                    var ch1 = g1.Chart;
                    ch1.ChartType = Excel.XlChartType.xlLine;
                    ch1.SetSourceData(rngSold);
                    ch1.SeriesCollection(1).XValues = rngLuna;
                    ch1.SeriesCollection(1).Name = "Sold rămas";
                    var s2 = ch1.SeriesCollection().NewSeries();
                    s2.Values = rngDobanda; s2.XValues = rngLuna; s2.Name = "Dobândă lunară";
                    ch1.HasTitle = true;
                    ch1.ChartTitle.Text = "Evoluţia soldului şi a dobânzii";
                    break;

                case FormGraficSelector.TipGrafic.StructuraPct:
                    var g2 = ws.ChartObjects().Add(300, 40, 480, 280);
                    g2.Name = "GraficPct";
                    var ch2 = g2.Chart;
                    ch2.ChartType = Excel.XlChartType.xlColumnStacked100;
                    ch2.SetSourceData(ws.Range[$"B2:C{lastRow}"]);
                    ch2.SeriesCollection(1).Name = "Principal";
                    ch2.SeriesCollection(2).Name = "Dobândă";
                    ch2.SeriesCollection(1).XValues = rngLuna;
                    ch2.HasTitle = true;
                    ch2.ChartTitle.Text = "% Structură plată lunară";
                    break;

                case FormGraficSelector.TipGrafic.Combo:
                    var g3 = ws.ChartObjects().Add(300, 40, 480, 280);
                    g3.Name = "GraficCombo";
                    var ch3 = g3.Chart;
                    ch3.ChartType = Excel.XlChartType.xlColumnClustered;
                    ch3.SetSourceData(rngDobanda);
                    ch3.SeriesCollection(1).XValues = rngLuna;
                    ch3.SeriesCollection(1).Name = "Dobândă lunară";
                    var soldLine = ch3.SeriesCollection().NewSeries();
                    soldLine.Values = rngSold;
                    soldLine.XValues = rngLuna;
                    soldLine.Name = "Sold rămas";
                    soldLine.ChartType = Excel.XlChartType.xlLine;
                    soldLine.AxisGroup = Excel.XlAxisGroup.xlSecondary;
                    ch3.HasTitle = true;
                    ch3.ChartTitle.Text = "Dobândă vs Sold (axă secundară)";
                    break;
            }
        }

        //private void btnVizualizareGrafic_Click(object sender, RibbonControlEventArgs e)
        //{
        //    if (excelApp == null) return;

        //    Excel.Worksheet ws;
        //    try { ws = (Excel.Worksheet)excelApp.Worksheets["Plan Financiar Plata"]; }
        //    catch { MessageBox.Show("Întâi generaţi planul financiar."); return; }

        //    //reset grafice
        //    foreach (Excel.ChartObject c in ws.ChartObjects())
        //    {
        //        string n = c.Name;
        //        if (n == "GraficPlan" || n == "GraficPct" || n == "GraficCombo")
        //            c.Delete();
        //    }

        //    int lastRow = ws.Cells[ws.Rows.Count, "A"].End(Excel.XlDirection.xlUp).Row;
        //    if (lastRow < 3) { MessageBox.Show("Tabelul pare gol."); return; }

        //    Excel.Range rngLuna = ws.Range[$"A2:A{lastRow}"];
        //    Excel.Range rngSold = ws.Range[$"E2:E{lastRow}"];   // sold rămas
        //    Excel.Range rngDobanda = ws.Range[$"C2:C{lastRow}"];   // dobândă lunară
        //    Excel.Range rngPrincipal = ws.Range[$"B2:B{lastRow}"]; // principal lunar

        //    //grafic linie sold + dobânda
        //    Excel.ChartObject g1 = ((Excel.ChartObjects)ws.ChartObjects())
        //                           .Add(300, 40, 480, 280);
        //    g1.Name = "GraficPlan";
        //    Excel.Chart ch1 = g1.Chart;
        //    ch1.ChartType = Excel.XlChartType.xlLine;
        //    ch1.SetSourceData(rngSold);
        //    ch1.SeriesCollection(1).XValues = rngLuna;
        //    ch1.SeriesCollection(1).Name = "Sold rămas";
        //    Excel.Series s2 = ch1.SeriesCollection().NewSeries();
        //    s2.Values = rngDobanda; s2.XValues = rngLuna; s2.Name = "Dobândă lunară";

        //    ch1.HasTitle = true;
        //    ch1.ChartTitle.Text = "Evoluţia soldului şi a dobânzii";

        //    var axX = ch1.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
        //    axX.HasTitle = true;
        //    axX.AxisTitle.Caption = "Lună";          

        //    var axY = ch1.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
        //    axY.HasTitle = true;
        //    axY.AxisTitle.Caption = "Lei";

        //    //grafic coloane suprapuse (procent principal vs dobândă)
        //    Excel.ChartObject g2 = ((Excel.ChartObjects)ws.ChartObjects())
        //                           .Add(800, 40, 420, 280);
        //    g2.Name = "GraficPct";
        //    Excel.Chart ch2 = g2.Chart;
        //    ch2.ChartType = Excel.XlChartType.xlColumnStacked100;
        //    ch2.SetSourceData(ws.Range[$"B2:C{lastRow}"]);   // principal + dobândă
        //    ch2.SeriesCollection(1).Name = "Principal";
        //    ch2.SeriesCollection(2).Name = "Dobândă";
        //    ch2.SeriesCollection(1).XValues = rngLuna;
        //    ch2.HasTitle = true;
        //    ch2.ChartTitle.Text = "% structură plată lunară";

        //    //grafic Combo: coloane dobândă + linie sold pe axa secundara
        //    Excel.ChartObject g3 = ((Excel.ChartObjects)ws.ChartObjects())
        //                           .Add(800, 340, 420, 280);
        //    g3.Name = "GraficCombo";
        //    Excel.Chart ch3 = g3.Chart;
        //    ch3.ChartType = Excel.XlChartType.xlColumnClustered;
        //    ch3.SetSourceData(rngDobanda);                    
        //    ch3.SeriesCollection(1).XValues = rngLuna;
        //    ch3.SeriesCollection(1).Name = "Dobândă lunară";

        //    Excel.Series soldLine = ch3.SeriesCollection().NewSeries();
        //    soldLine.Values = rngSold;
        //    soldLine.XValues = rngLuna;
        //    soldLine.Name = "Sold rămas";
        //    soldLine.ChartType = Excel.XlChartType.xlLine;
        //    soldLine.AxisGroup = Excel.XlAxisGroup.xlSecondary; 

        //    ch3.HasTitle = true;
        //    ch3.ChartTitle.Text = "Dobândă vs sold (axă secundară)";
        //}

        private void btnPaymentPlan_Click(object sender, RibbonControlEventArgs e)
        {
            GenereazaPlan();
        }
        private void GenereazaPlan(double sumaInAvans = 0.0)
        {
 
            int idClient = -1, idOferta = -1;
            using (var con = new OleDbConnection(ConnString))
            {
                con.Open();
                var cmd = new OleDbCommand(
                    "SELECT IdClientActual, IdOfertaActuala FROM DateActuale WHERE ID=1", con);
                var reader = cmd.ExecuteReader();
                if (reader.Read()) { idClient = reader.GetInt32(0); idOferta = reader.GetInt32(1); }
                
                if (idClient <= 0 || idOferta <= 0)
                {
                    MessageBox.Show("Nu există client / ofertă curentă."); return;
                }

     
                string nume = "", prenume = "", venitCat = "";
                double pretZi = 0; int durataZile = 0;

                var cCmd = new OleDbCommand("SELECT Nume,Prenume,Venit FROM Clienti WHERE Id=?", con);
                cCmd.Parameters.AddWithValue("?", idClient);
                var rc = cCmd.ExecuteReader();
                if (rc.Read())
                {
                    nume = rc.GetString(0); prenume = rc.GetString(1); venitCat = rc.GetString(2);
                }

                var oCmd = new OleDbCommand("SELECT PretZi,DurataMinima FROM Oferte WHERE IdOferta=?", con);
                oCmd.Parameters.AddWithValue("?", idOferta);
                var ro = oCmd.ExecuteReader();
                if (ro.Read())
                {
                    pretZi = Convert.ToDouble(ro.GetDecimal(0));
                    durataZile = ro.GetInt32(1);
                }


    
                int luni = (int)Math.Ceiling(durataZile / 30.0);
                double principal = pretZi * durataZile - sumaInAvans;     
                if (principal <= 0) { MessageBox.Show("Plata în avans depăşeşte costul total."); return; }

                double rataAnuala = VenitCategorieDobanda.TryGetValue(venitCat, out var r) ? r : 0.1;
                double rataLunara = rataAnuala / 12.0;

                double plataLunara = principal * rataLunara / (1 - Math.Pow(1 + rataLunara, -luni));

                Excel.Worksheet ws;
                try { ws = (Excel.Worksheet)excelApp.Worksheets["Plan Financiar Plata"]; ws.Cells.Clear(); }
                catch { ws = excelApp.Worksheets.Add(); ws.Name = "Plan Financiar Plata"; }

                // header
                string[] h = { "Luna", "Principal", "Dobândă", "Plata", "Sold rămas" };
                for (int i = 0; i < h.Length; i++) ws.Cells[1, i + 1] = h[i];
                ws.Range["A1:E1"].Font.Bold = true;

    
                double sold = principal;
                int row = 2;
                for (int luna = 1; luna <= luni; luna++)
                {
                    double dobanda = sold * rataLunara;
                    double principalPlat = plataLunara - dobanda;
                    if (luna == luni)                      
                    {
                        principalPlat = sold;
                        plataLunara = principalPlat + dobanda;
                    }
                    sold -= principalPlat;

                    ws.Cells[row, 1] = luna;
                    ws.Cells[row, 2] = principalPlat;
                    ws.Cells[row, 3] = dobanda;
                    ws.Cells[row, 4] = plataLunara;
                    ws.Cells[row, 5] = Math.Max(sold, 0);
                    row++;
                }

                ws.Range[$"B2:D{row - 1}"].NumberFormat = "#,##0";
                ws.Columns["A:E"].AutoFit();


                int lastDataRow = row - 1;
                Excel.Range tblRange = ws.Range[$"A1:E{lastDataRow}"];

                Excel.ListObject lo;
                try
                {
                    lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                                                           tblRange, null, Excel.XlYesNoGuess.xlYes);
                }
                catch { lo = ws.ListObjects[1]; lo.Resize(tblRange); }  // deja există

                lo.Name = "tblPlan";
                lo.TableStyle = "TableStyleMedium12";

                Excel.Range header = ws.Range["A1:E1"];
                header.Font.Size = 12;
                header.Font.Bold = true;
                header.Interior.Color = 0xa48ed6;

                tblRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                tblRange.Borders.Color = 0x7c87e8;
                tblRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                tblRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                tblRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                MessageBox.Show("Plan financiar generat / actualizat.");
            }
        }

        private void btnPlataInAvans_Click(object sender, RibbonControlEventArgs e)
        {
            var f = new FormPlataAvans();
            f.StartPosition = FormStartPosition.CenterParent;

            if (f.ShowDialog() == DialogResult.OK)
                GenereazaPlan(f.Avans);
        }

        private void btnExportPDF_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Range selectie = (Excel.Range)excelApp.Selection;

                if (selectie == null || selectie.Rows.Count < 1)
                {
                    MessageBox.Show("Selectați celulele corespunzătoare perioadei dorite.");
                    return;
                }

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string fileName = $"RaportFinanciar_{DateTime.Now:dd-MM-yyyy_HH-mm}.pdf";
                string fullPath = Path.Combine(desktopPath, fileName);

                selectie.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    fullPath,
                    Excel.XlFixedFormatQuality.xlQualityStandard
                );

                MessageBox.Show($"Raport exportat cu succes pe Desktop:\n{fullPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la export PDF: " + ex.Message);
            }
        }

        private void btnGraficSelectie_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Worksheet ws = (Excel.Worksheet)excelApp.ActiveSheet;
                Excel.Range selectie = (Excel.Range)excelApp.Selection;

                if (selectie == null || selectie.Rows.Count < 3)
                {
                    MessageBox.Show("Selectați cel puțin 3 rânduri: etichete + 2 seturi de valori.");
                    return;
                }

                int firstCol = selectie.Column;
                int lastCol = firstCol + selectie.Columns.Count - 1;

                int labelRow = selectie.Row;
                int serie1Row = labelRow + 1;
                int serie2Row = labelRow + 2;

                int nrPuncte = lastCol - firstCol + 1;
                string[] etichete = new string[nrPuncte];
                double[] serie1 = new double[nrPuncte];
                double[] serie2 = new double[nrPuncte];

                for (int i = 0; i < nrPuncte; i++)
                {
                    var labelCell = ws.Cells[labelRow, firstCol + i];
                    var val1Cell = ws.Cells[serie1Row, firstCol + i];
                    var val2Cell = ws.Cells[serie2Row, firstCol + i];

                    etichete[i] = labelCell.Text.ToString();

                    if (double.TryParse(val1Cell.Value2?.ToString(), out double val1))
                        serie1[i] = val1;
                    if (double.TryParse(val2Cell.Value2?.ToString(), out double val2))
                        serie2[i] = -Math.Abs(val2); 
                }

                Excel.ChartObjects charts = (Excel.ChartObjects)ws.ChartObjects();
                Excel.ChartObject chartObj = charts.Add(300, 100, 600, 350);
                chartObj.Name = "GraficBipolar";
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlBarClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Comparație Bidirecțională";

                Excel.SeriesCollection series = (Excel.SeriesCollection)chart.SeriesCollection();

                Excel.Series s1 = series.NewSeries();
                s1.Values = serie1;
                s1.XValues = etichete;
                s1.Name = "Luna 1";

                Excel.Series s2 = series.NewSeries();
                s2.Values = serie2;
                s2.XValues = etichete;
                s2.Name = "Luna 2";

                chart.Axes(Excel.XlAxisType.xlValue).HasTitle = true;
                chart.Axes(Excel.XlAxisType.xlValue).AxisTitle.Text = "Valori (lei)";
                chart.Axes(Excel.XlAxisType.xlCategory).ReversePlotOrder = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la generarea graficului bipolar: " + ex.Message);
            }

        }
    }
}
