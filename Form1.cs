using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Data.OleDb;
using System.Windows.Forms.DataVisualization.Charting;


namespace Projekt3_Stankiewicz53518
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            rsStyl_linii("Ciągła");
            rschart1.Titles.Add("Wykres funkcji");
            rsradioButton2.Checked = true;
        }

        private void button_koloru_linii_Click(object sender, EventArgs e)
        {
            rsKolor_linii();
        }

        private void button_koloru_tla_Click(object sender, EventArgs e)
        {
            rsKolor_tla();
        }

        private void grubosc_linii_Scroll(object sender, EventArgs e)
        {
            rsGrubosc_linii(rssuwak_grubosc_linii.Value);
        }

        private void grubosc_linii_liczba_ValueChanged(object sender, EventArgs e)
        {
            rsGrubosc_linii((int)rsgrubosc_linii_liczba.Value);
            rschart1.Series["Wartość funkcji f(x)"].BorderWidth = rssuwak_grubosc_linii.Value;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void kolorTłaWykresuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsKolor_tla();
        }

        private void kolorLiniiWykresuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsKolor_linii();
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            rsGrubosc_linii(1);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            rsGrubosc_linii(2);
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            rsGrubosc_linii(3);
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            rsGrubosc_linii(4);
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            rsGrubosc_linii(5);
        }

        private void ciągłaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsStyl_linii("Ciągła");
        }

        private void kropkowaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsStyl_linii("Kropkowa");
        }

        private void kreskowaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsStyl_linii("Kreskowa");
        }

        private void kreskowokropkowaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsStyl_linii("Kreskowo-kropkowana");
        }
        private void styl_lini_SelectedIndexChanged(object sender, EventArgs e)
        {
            rsZmianaLiniiWykresu();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            rsZmianaOpisuWykresu();
        }

        public bool rsXWarunki()
        {
            double rsZmienna;
            if (rswartosc_x.Text == "")
            {
                errorProvider1.SetError(rswartosc_x, "Nie może być pusty!");
                return false;
            }
            else if (!rsCzy_liczba(rswartosc_x.Text, out rsZmienna))
            {
                errorProvider1.SetError(rswartosc_x, "Podana wartość nie jest liczbą!");
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool rsEpsWarunki()
        {
            double rsEps;
            if (rswartosc_eps.Text == "")
            {
                errorProvider1.SetError(rswartosc_eps, "Nie może być pusty!");
                return false;
            }
            else if (!rsCzy_liczba(rswartosc_eps.Text, out rsEps))
            {
                errorProvider1.SetError(rswartosc_eps, "Podana wartość nie jest liczbą!");
                return false;
            }
            else if (rsEps <= 0.0F || rsEps >= 1.0F)
            {
                errorProvider1.SetError(rswartosc_eps, "Podana wartość nie spełnia założeń!");
                return false;
            }
            else
            {
                return true;
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();    
            if(rsXWarunki() && rsEpsWarunki())
            {
                double rsX, rsEps; 
                rsCzy_liczba(rswartosc_eps.Text, out rsX);
                rsCzy_liczba(rswartosc_eps.Text, out rsEps);
                rsCysc_srodek();
                rslabel16.Visible = true;
                double rsWynik = rsSumaSzeregu(rsX, rsEps);
                rslabel16.Image = null;
                rslabel16.Text = "Wartość F(X) z podanych danych dla danego szeregu wynosi: " + rsWynik;
            }
            else
            {
                rslabel16.Image = Properties.Resources.Przechwytywanie;
                rslabel16.Text = "";
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Reset pól wpisywania
            rswartosc_x.Text = "";
            rswartosc_eps.Text = "";
            rswartosc_xd.Text = "";
            rswartosc_xg.Text = "";
            rswartosc_h.Text = "";

            //Reset całkowania
            rscalki_dokladnosc.Text = "";
            rscalki_dolna_granica.Text = "";
            rscalki_gorna_granica.Text = "";
            rscalki_wynik.Text = "";

            //Grobosc lini
            rsGrubosc_linii(1);

            //Wygląd wykresu
            rswziernik_kolor_tla.BackColor= Color.FromArgb(255, 224, 192);
            rswziernik_kolor_linii.BackColor = Color.Blue;
            rsStyl_linii("Ciągła");

            //Opcje radio
            rsradioButton1.Checked = false;
            rsradioButton2.Checked = true;

            //Reset środkowego labela
            rslabel16.Image = Properties.Resources.Przechwytywanie;
            rslabel16.Text = "";
            rslabel16.Visible = true;
            rstabela.Visible = false;
            rschart1.Visible = false;

            //Restet w menu
            pogrubionaIKursywaToolStripMenuItem.Checked = false;
            kursywaToolStripMenuItem.Checked = false;
            pogrubionaToolStripMenuItem.Checked = false;

            //Czyszczenie błędów
            errorProvider1.Clear();

            //Reset koloru przycisków
            rsbutton1.ForeColor = Color.Black;
            rsbutton2.ForeColor = Color.Black;
            rsbutton3.ForeColor = Color.Black;
            rsbutton4.ForeColor = Color.Black;
            rsbutton5.ForeColor = Color.Black;
            rsbutton_koloru_tla.ForeColor = Color.Black;
            rsbutton_koloru_linii.ForeColor = Color.Black;
        }

        public bool rsXdXgWarunki()
        {
            double rsXd, rsXg;
            if(rswartosc_xd.Text == "")
            {
                errorProvider1.SetError(rswartosc_xd, "Nie może być pusty!");
                return false;
            }
            else if(!rsCzy_liczba(rswartosc_xd.Text, out rsXd))
            {
                errorProvider1.SetError(rswartosc_xd, "Podana wartość nie jest liczbą!");
                return false;
            }

            else if (rswartosc_xg.Text == "")
            {
                errorProvider1.SetError(rswartosc_xg, "Nie może być pusty!");
                return false;
            }
            else if (!rsCzy_liczba(rswartosc_xg.Text, out rsXg))
            {
                errorProvider1.SetError(rswartosc_xg, "Podana wartość nie jest liczbą!");
                return false;
            }

            else if(rsXd > rsXg)
            {
                errorProvider1.SetError(rswartosc_xd, "Podana wartość nie spełnia założeń!");
                errorProvider1.SetError(rswartosc_xg, "Podana wartość nie spełnia założeń!");
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool rsHWarunki()
        {
            double rsH;
            if (rswartosc_h.Text == "")
            {
                errorProvider1.SetError(rswartosc_h, "Nie może być pusty!");
                return false;
            }
            else if (!rsCzy_liczba(rswartosc_h.Text, out rsH))
            {
                errorProvider1.SetError(rswartosc_h, "Podana wartość nie jest liczbą!");
                return false;
            }
            else if(rsH <= 0.0F || rsH >= 1.0F)
            {
                errorProvider1.SetError(rswartosc_h, "Podana wartość nie spełnia założeń!");
                return false;
            }
            else
            {
                return true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            double rsEps, rsXd, rsXg, rsH;
            errorProvider1.Clear();
            if(rsCzy_liczba(rswartosc_eps.Text, out rsEps) || rswartosc_eps.Text=="")
            {
                rsEps = 0.000001;
            }

            if (rsXdXgWarunki() && rsHWarunki())
            {
                rsCzy_liczba(rswartosc_xg.Text, out rsXg);
                rsCzy_liczba(rswartosc_xd.Text, out rsXd);
                rsCzy_liczba(rswartosc_h.Text, out rsH);

                rsCysc_srodek();
                if (!pogrubionaIKursywaToolStripMenuItem.Checked && !kursywaToolStripMenuItem.Checked)
                {
                    pogrubionaToolStripMenuItem.Checked = true;
                }

                rstabela.Rows.Clear();
                rstabela.Visible = true;
                do
                {
                    double rsWynikFunkcji = rsSumaSzeregu(rsXg, rsEps);
                    rsXg = rsXg - rsH;
                    rstabela.Rows.Add(Math.Round(rsXg + rsH, 2), Math.Round(rsWynikFunkcji, 2));
                } while (rsXd < rsXg);
            }
            else
            {
                rsCysc_srodek();
            }

        }

        private void zapiszTablicęWPlikuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RsZapiszTablice();
        }

        private void odczytajTablicęZPlikuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsOdczytTablicy();
        }

        private void button3_Click(object sender, EventArgs e)
        {

            double rsEps, rsXd, rsXg, rsH;
            errorProvider1.Clear();

            if (rsCzy_liczba(rswartosc_eps.Text, out rsEps) || rswartosc_eps.Text == "")
            {
                rsEps = 0.000001;

            }
            if (rsXdXgWarunki() && rsHWarunki())
            {
                rsCzy_liczba(rswartosc_xg.Text, out rsXg);
                rsCzy_liczba(rswartosc_xd.Text, out rsXd);
                rsCzy_liczba(rswartosc_h.Text, out rsH);

                rsCysc_srodek();
                rschart1.Visible = true;
                rschart1.BackColor = rswziernik_kolor_tla.BackColor;
                rschart1.Series["Wartość funkcji f(x)"].Points.Clear();
                rschart1.Series["Wartość funkcji f(x)"].Color = rswziernik_kolor_linii.BackColor;
                rschart1.Legends.FindByName("Legend1").Docking = Docking.Bottom;
                ChartArea chartArea = new ChartArea();


                rsZmianaOpisuWykresu();

                rsZmianaLiniiWykresu();

                do
                {
                    double rsWynikFunkcji = rsSumaSzeregu(rsXg, rsEps);

                    rschart1.Series["Wartość funkcji f(x)"].Points.AddXY(rsXg, rsWynikFunkcji);
                    rsXg = rsXg - rsH;
                } while (rsXd < rsXg);
            }
            else
            {
                rsCysc_srodek();
            }
        }

        private void pogrubionaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsTabelaBold();
        }

        private void kursywaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsTabelaItalic();
        }

        private void pogrubionaIKursywaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsTabelaBoth();
        }

        private void krójPismaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fontDialog1.ShowColor = false;
            fontDialog1.ShowApply = false;
            fontDialog1.ShowEffects = false;
            fontDialog1.ShowHelp = false;

            if (fontDialog1.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                rstabela.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);

                if (fontDialog1.Font.Style == FontStyle.Bold) rsTabelaBold();
                else if (fontDialog1.Font.Style == FontStyle.Italic) rsTabelaItalic();
                else if (fontDialog1.Font.Style == FontStyle.Regular)
                {
                    pogrubionaToolStripMenuItem.Checked = false;
                    kursywaToolStripMenuItem.Checked = false;
                    pogrubionaIKursywaToolStripMenuItem.Checked = false;
                }
                else rsTabelaBoth();
            }
        }





        public void rsZmianaOpisuWykresu()
        {
            if (rsradioButton2.Checked)
            {
                rschart1.ChartAreas[0].AxisX.Title = "Wartość f(x)";
                rschart1.ChartAreas[0].AxisY.Title = "Wartość X";
            }
            else
            {
                rschart1.ChartAreas[0].AxisX.Title = "";
                rschart1.ChartAreas[0].AxisY.Title = "";
            }
        }

        public void rsTabelaBold()
        {
            if (!pogrubionaToolStripMenuItem.Checked)
            {
                pogrubionaToolStripMenuItem.Checked = true;
                kursywaToolStripMenuItem.Checked = false;
                pogrubionaIKursywaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Bold);
            }
            else
            {
                pogrubionaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Regular);
            }
        }

        public void rsTabelaBoth()
        {
            if (!pogrubionaIKursywaToolStripMenuItem.Checked)
            {
                pogrubionaIKursywaToolStripMenuItem.Checked = true;
                kursywaToolStripMenuItem.Checked = false;
                pogrubionaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Bold | FontStyle.Italic);
            }
            else
            {
                pogrubionaIKursywaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Regular);
            }
        }

        public void rsTabelaItalic()
        {
            if (!kursywaToolStripMenuItem.Checked)
            {
                kursywaToolStripMenuItem.Checked = true;
                pogrubionaToolStripMenuItem.Checked = false;
                pogrubionaIKursywaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Italic);
            }
            else
            {
                kursywaToolStripMenuItem.Checked = false;
                rstabela.Font = new Font(rstabela.Font, FontStyle.Regular);
            }
        }

        public void rsZmianaLiniiWykresu()
        {
            switch (rsstyl_lini.Text)
            {
                case "Ciągła":
                    rschart1.Series["Wartość funkcji f(x)"].BorderDashStyle = ChartDashStyle.Solid;
                    ciągłaToolStripMenuItem.Checked = true;
                    kropkowaToolStripMenuItem.Checked = false;
                    kreskowaToolStripMenuItem.Checked = false;
                    kreskowokropkowaToolStripMenuItem.Checked = false;
                    break;
                case "Kropkowa":
                    rschart1.Series["Wartość funkcji f(x)"].BorderDashStyle = ChartDashStyle.Dot;
                    ciągłaToolStripMenuItem.Checked = false;
                    kropkowaToolStripMenuItem.Checked = true;
                    kreskowaToolStripMenuItem.Checked = false;
                    kreskowokropkowaToolStripMenuItem.Checked = false;
                    break;
                case "Kreskowa":
                    rschart1.Series["Wartość funkcji f(x)"].BorderDashStyle = ChartDashStyle.Dash;
                    ciągłaToolStripMenuItem.Checked = false;
                    kropkowaToolStripMenuItem.Checked = false;
                    kreskowaToolStripMenuItem.Checked = true;
                    kreskowokropkowaToolStripMenuItem.Checked = false;
                    break;
                case "Kreskowo-kropkowana":
                    rschart1.Series["Wartość funkcji f(x)"].BorderDashStyle = ChartDashStyle.DashDot;
                    ciągłaToolStripMenuItem.Checked = false;
                    kropkowaToolStripMenuItem.Checked = false;
                    kreskowaToolStripMenuItem.Checked = false;
                    kreskowokropkowaToolStripMenuItem.Checked = true;
                    break;
            }
        }

        public void rsCysc_srodek()
        {
            rstabela.Visible = false;
            rslabel16.Visible = false;
            rschart1.Visible = false;
        }

        public void rsKolor_tla()
        {
            if (colorDialog_linia.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                rswziernik_kolor_tla.BackColor = colorDialog_linia.Color;
                rschart1.BackColor = rswziernik_kolor_tla.BackColor;
            }
        }


        public void rsKolor_linii()
        {
            if (colorDialog_linia.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                rswziernik_kolor_linii.BackColor = colorDialog_linia.Color;
                rschart1.Series["Wartość funkcji f(x)"].Color = rswziernik_kolor_linii.BackColor;
            }
        }
        public void rsGrubosc_linii(int rsgrubosc)
        {
            rssuwak_grubosc_linii.Value = rsgrubosc;
            rsgrubosc_linii_liczba.Value = rsgrubosc;
        }

        public void rsStyl_linii(string rsstyl)
        {
            rsstyl_lini.SelectedItem = rsstyl;
        }

      

        public bool rsCzy_liczba(string rsPrzyjmowana, out double rsZwracana)
        {
            if(double.TryParse(rsPrzyjmowana, out rsZwracana))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public double rsSumaSzeregu(double rsX, double rsEps)
        {
            double rsw1 = (rsX + 1) / 1;
            double rsSumaSzeregu = rsw1;
            int rsSilnia = 2;

            double rsw2;

            for (int rsn = 2; rsn > 0; rsn++)
            {
                rsw2 =Math.Pow((rsX+1), rsn)/rsSilnia;

                rsSumaSzeregu += rsw2;
                if (rsw2 - rsw1 <= rsEps) break;
                rsw1 = rsw2;
                rsSilnia = rsSilnia * rsn;
            }
            return rsSumaSzeregu;
        }
        /////////////////////////CAŁKI/////////////////////////////////

        private void rsbutton5_Click(object sender, EventArgs e)
        {
            double rsXd, rsXg, rsEps;
            errorProvider1.Clear();

            //Warunki dolnej granicy
            if (rscalki_dolna_granica.Text == "")
            {
                errorProvider1.SetError(rscalki_dolna_granica, "Wartość nie może być pusta!");
            }
            else if (!rsCzy_liczba(rscalki_dolna_granica.Text, out rsXd))
            {
                errorProvider1.SetError(rscalki_dolna_granica, "Wartość musi być liczbą!");
            }
            //warunki górnej granicy
            else if (rscalki_gorna_granica.Text == "")
            {
                errorProvider1.SetError(rscalki_gorna_granica, "Wartość nie może być pusta!");
            }
            else if (!rsCzy_liczba(rscalki_gorna_granica.Text, out rsXg))
            {
                errorProvider1.SetError(rscalki_gorna_granica, "Wartość musi być liczbą!");
            }
            //warunki eps
            else if (rscalki_dokladnosc.Text == "")
            {
                errorProvider1.SetError(rscalki_dokladnosc, "Wartość nie może być pusta!");
            }
            else if (!rsCzy_liczba(rscalki_dokladnosc.Text, out rsEps))
            {
                errorProvider1.SetError(rscalki_dokladnosc, "Wartość musi być liczbą!");
            }
            else if (rsEps <= 0.0F || rsEps >= 0.05F)
            {
                errorProvider1.SetError(rscalki_dokladnosc, "Podana wartość nie spełnia założeń!");
            }
            //Warunek granic
            else if (rsXd > rsXg)
            {
                errorProvider1.SetError(rscalki_dolna_granica, "Podana wartość nie spełnia założeń!");
                errorProvider1.SetError(rscalki_gorna_granica, "Podana wartość nie spełnia założeń!");
            }
            //Metoda całkowania
            else if ((string)rsmetoda_calkowania.SelectedItem != "Prostokątów" && (string)rsmetoda_calkowania.SelectedItem != "Trapezów")
            {
                errorProvider1.SetError(rsmetoda_calkowania, "Wybierz metodę całkowania!");
            }


            ////////////////
            ////Oblicznie///
            ////////////////
            else
            {
                switch((string)rsmetoda_calkowania.SelectedItem)
                {
                    case "Prostokątów":
                        rsMetodaProstokatow();
                        break;

                    case "Trapezów":
                        rsMetodaTrapezow();
                        break;
                }
            }

        }

        private void rsMetodaProstokatow()
        {
            double rsa, rsb, rseps;
            rsCzy_liczba(rscalki_dolna_granica.Text, out rsa);
            rsCzy_liczba(rscalki_gorna_granica.Text, out rsb);
            rsCzy_liczba(rscalki_dokladnosc.Text, out rseps);

            double rsH, rsCi, rsCi_1, rsSumaFx;
            ushort rsLicznikWytrazowSzeregu;
            double rsX;
            int rsLicznikPrzedzialow = 1;

            rsCi = (rsb - rsa) * rsSumaSzeregu((rsa+rsb)/2.0F, rseps);

            do
            {
                rsCi_1 = rsCi;
                rsLicznikPrzedzialow++;
                rsH = (rsb - rsa) / rsLicznikPrzedzialow;
                rsX = rsa + rsH / 2.0F;
                rsSumaFx = 0.0F;
                for (ushort rsi = 0; rsi < rsLicznikPrzedzialow; rsi++)
                    rsSumaFx += rsSumaSzeregu(rsX + rsi * rsH, rseps);
                rsCi = rsH * rsSumaFx;
            } while (Math.Abs(rsCi - rsCi_1) > rseps);


            rscalki_wynik.Text = "" + rsCi;
        }

        private void rsMetodaTrapezow()
        {
            double rsd, rsg, rseps;
            rsCzy_liczba(rscalki_dolna_granica.Text, out rsd);
            rsCzy_liczba(rscalki_gorna_granica.Text, out rsg);
            rsCzy_liczba(rscalki_dokladnosc.Text, out rseps);


            double rsH, rsCi, rsCi_1, rsSumaFx;

            rsH = rsg - rsd;

            double rsSumaFaFb = rsSumaSzeregu(rsd, rseps) + rsSumaSzeregu(rsg, rseps);
            rsCi = rsH * rsSumaFaFb;
            int rsLicznikIteracji = 1;
            do
            {
                rsCi_1 = rsCi;
                rsLicznikIteracji++;
                rsH = (rsg - rsd) / rsLicznikIteracji;
                rsSumaFx = 0.0F;
                for(int rsj=1;rsj<rsLicznikIteracji;rsj++)
                    rsSumaFx += rsSumaSzeregu(rsd + rsj * rsH, rseps);
                rsCi = rsH * (rsSumaFaFb + rsSumaFx);
            } while (Math.Abs(rsCi-rsCi_1)>rseps);

            rscalki_wynik.Text = "" + rsCi;
        }
        public void RsZapiszTablice()
        {
            if (rstabela.Visible == true)
            {
                Microsoft.Office.Interop.Excel._Application rsapp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook rsworkbook = rsapp.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet rsworksheet = null;
                rsworksheet = rsworkbook.Sheets[1];
                rsworksheet = rsworkbook.ActiveSheet;
                rsworksheet.Name = "Sheet1";


                rsworksheet.Cells[1, 1] = "Wartość zmiennej niezależnej X";
                rsworksheet.Cells[1, 2] = "Wartość funkcji F(X)";


                for (int rsi = 0; rsi < rstabela.Rows.Count; rsi++)
                {
                    for (int rsj = 0; rsj < rstabela.Columns.Count; rsj++)
                    {
                        if (rstabela.Rows[rsi].Cells[rsj].Value != null)
                        {
                            rsworksheet.Cells[rsi + 2, rsj + 1] = rstabela.Rows[rsi].Cells[rsj].Value.ToString();
                        }
                    }

                }



                var saveFileDialoge = new SaveFileDialog();
                saveFileDialoge.FileName = "output";
                saveFileDialoge.DefaultExt = ".xlsx";

                if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    rsworkbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                rsapp.Quit();
            }
            else
            {
                MessageBox.Show("Tablica musi zostać pierw stworzona", "Błąd zapisu tablicy", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void rsOdczytTablicy()
        {
            rstabela.Rows.Clear();
            rstabela.Visible = true;
            Microsoft.Office.Interop.Excel.Application rsapp;
            Microsoft.Office.Interop.Excel.Workbook rsworkbook;
            Microsoft.Office.Interop.Excel.Worksheet rsworksheet;
            Microsoft.Office.Interop.Excel.Range rsrange;

            int rsRow;
            string rsstrfileName;

            openFileDialog1.Filter = "Excel Office | *.xls; *xlsx";
            openFileDialog1.ShowDialog();
            rsstrfileName = openFileDialog1.FileName;

            if (rsstrfileName != string.Empty)
            {
                rsapp = new Microsoft.Office.Interop.Excel.Application();
                rsworkbook = rsapp.Workbooks.Open(rsstrfileName);
                rsworksheet = rsworkbook.Worksheets["Sheet1"];
                rsrange = rsworksheet.UsedRange;

                for (rsRow = 2; rsRow <= rsrange.Rows.Count; rsRow++)
                {
                    rstabela.Rows.Add(rsrange.Cells[rsRow, 1].Text, rsrange.Cells[rsRow, 2].Text);
                }
                rsworkbook.Close();
                rsapp.Quit();
            }
        }

        //Sprawdzian 

        //Zapisywanie
        private void zapiszWierszeDanychKontrolkiDataGridViewWPlikuZewnętrzymToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RsZapiszTablice();
        }
        //Odczyt
        private void odczytajDaneZPlikupoprzednioWNimZapisaneIWpiszJeDoKontrolkiDataGridViewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rsOdczytTablicy();
        }
        //Zmiana Czcionki Labelow
        private void zmieńCzcionkęWszystkichKontrolekFormularzaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fontDialog1.ShowColor = true;
            fontDialog1.ShowApply = true;
            fontDialog1.ShowEffects = true;
            fontDialog1.ShowHelp = true;

            if (fontDialog1.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                label1.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label1.ForeColor = fontDialog1.Color;

                label2.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label2.ForeColor = fontDialog1.Color;

                label3.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label3.ForeColor = fontDialog1.Color;

                label4.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label4.ForeColor = fontDialog1.Color;

                label5.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label5.ForeColor = fontDialog1.Color;

                label6.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label6.ForeColor = fontDialog1.Color;

                rslabel7.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel7.ForeColor = fontDialog1.Color;

                rslabel8.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel8.ForeColor = fontDialog1.Color;

                rslabel9.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel9.ForeColor = fontDialog1.Color;

                rslabel10.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel10.ForeColor = fontDialog1.Color;

                rslabel11.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel11.ForeColor = fontDialog1.Color;

                rslabel12.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                rslabel12.ForeColor = fontDialog1.Color;

                label13.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label13.ForeColor = fontDialog1.Color;

                label14.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label14.ForeColor = fontDialog1.Color;

                label15.Font = new Font(fontDialog1.Font.Name, fontDialog1.Font.Size, fontDialog1.Font.Style);
                label15.ForeColor = fontDialog1.Color;
            }
        }
        //Zmiana koloru przycisków
        private void zmieńKolorCzcionkiWszystkichKontrolekTypuButtonToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog_linia.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                rsbutton1.ForeColor = colorDialog_linia.Color;
                rsbutton2.ForeColor = colorDialog_linia.Color;
                rsbutton3.ForeColor = colorDialog_linia.Color;
                rsbutton4.ForeColor = colorDialog_linia.Color;
                rsbutton5.ForeColor = colorDialog_linia.Color;
                rsbutton_koloru_tla.ForeColor = colorDialog_linia.Color;
                rsbutton_koloru_linii.ForeColor = colorDialog_linia.Color;
            }
        }

    }
}
