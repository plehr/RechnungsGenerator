//Diese Klasse wurde von P Lehr (https://github.com/plehr) im Rahmen des Projektes "stnr" erstellt.
//Alle Eigentumsrechte bleiben bei P Lehr (https://github.com/plehr). 

// Information: Diese Klasse funktioniert soweit.
// Im Moment können nicht mehr als 20 Posten gedruckt werden.
// Das könnte durch eine weitere Methode weitereSeite(string seitenZahl);  behoben werden, sodass unbegrenzte Posten verfügbar sind.

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RechnungsGenerator
{
    class PDF
    {
        //Wichtige Variablen definieren
        private bool debug = false;
        private String Firmenname;
        private String Firmenslogan;
        private String Firmenstr;
        private String Firmenhnr;
        private int Firmenplz;
        private String Firmenort;
        private String Anrede;
        private String Vorname;
        private String Nachname;
        private String Straße;
        private int Hausnummer;
        private int Postleitzahl;
        private String Ort;
        private PrintDialog dialog;
        private int Rechnungsnr;
        private DateTime Rechnungsdatum;
        private DataTable tabelle = new DataTable();
        private String Mailaddr;
        private String Sachbearbeiter;
        private String Telefonnr;
        private String Faxnr;
        private String IBAN;
        private String BIC;
        private Int64 Steuernr;
        private int tabellentiefe;
        private int anzahlDatensätze;
        private double Zwischensumme;
        private double Versandkosten;
        private double mwst19;
        private double mwst7;
        private double rabatt;


        public void generieren(int RNR, PrintDialog dialog)
        {
            //Diese Methode wird von der GUI aus angesteuert und übergibt das Ergebnis des Printdialogs und die Rechnungsnummer.
            Rechnungsnr = RNR; // Die Variable Rechnungsnr wird mit der mitgegebenen Rechnungsnummer gefüllt.
            this.dialog = dialog; //Der lokale Printdialog wird mit den mitgelieferten Ergebnissen des Druckdialogs gefüllt.
            DBIO();//Hier wird die Methode DBIO aufgerufen, welche die Variablen mit den passenden Datenbankeinträgen füllt.
            ausdruck();//Dieser Methodenaufruf startet den Druckvorgang der 1.Seite
            if (debug)
            { // Für Debugzwecke wird am Ende jeder Verarbeitung eine Meldung ausgegeben.
                MessageBox.Show("Ihr Druck wurde erfolgreich auf dem Drucker: " + dialog.PrinterSettings.PrinterName.ToString() + " gedruckt!");
            }
         }
       
        private void DBIO()
        {
            //Datenbankzugriff um Kundendaten zu holen!
            //Aus Krankheitsgründen wird der Datenbankzugriff erst am Wochenende fertiggestellt.
            //Wir haben uns folgende Interaktion mit der zukünftigen Datenbankklasse überlegt: (Beispiel)
           //  

            //MOCK-Daten füllen
            //Bis die Interaktion mit der Datenbank funktioniert, wird hier der Datenbankzugriff simuliert, also Beispielwerte gesetzt.
            try { // Das "try" hat den Hintergund, das im falle einer fehlerhaften Datenbankinteraktion eine Fehlermeldung erscheint.
                Firmenname = "Beispielfirma";
                Firmenslogan = "hat sogar einen eigenen Slogan";
                Firmenstr = "Musterstr.";
                Firmenhnr = "42";
                Firmenplz = 12345;
                Firmenort = "Mustershausen";
                Anrede = "Herr";
                Vorname = "Maximilian";
                Nachname = "Mustermann";
                Straße = "Musterstraße";
                Versandkosten = 10.30;
                Hausnummer = 123;
                Postleitzahl = 32132;
                Ort = "Musterdorf";
                Rechnungsnr = 1234567890;
                Rechnungsdatum = new DateTime(2014, 11 ,21);
                Mailaddr = "erika@musterfrau.tld";
                Sachbearbeiter = "Erika Musterfrau";
                Telefonnr = "+49 123 3215 - 10";
                Faxnr = "+49 123 3215 - 20";
                IBAN = "DE42 4242 0420 0000 0424 29";
                Steuernr = 04242424200;
                rabatt = 10;
                BIC = "BERLADE42XXX";
                        
                //Tabelle füllen
                tabelle.Columns.Add("Menge");
                tabelle.Columns.Add("Artikelnummer");
                tabelle.Columns.Add("Beschreibung");
                tabelle.Columns.Add("Einzelpreis Netto");
                tabelle.Columns.Add("MwSt");

                // Die Namen der Produkte müssen natürlich nicht stimmen
                tabelle.Rows.Add(2, 123456, "Produkt1", "85,00", 19);
                tabelle.Rows.Add(2, 123456, "Produkt2", "85,00", 7);
                tabelle.Rows.Add(2, 654321, "Produkt3", "41,00", 7);
                tabelle.Rows.Add(3, 654123, "Produkt4", "12,00", 7);
                tabelle.Rows.Add(4, 123654, "Produkt4", "52,00", 7);
                tabelle.Rows.Add(5, 415263, "Produkt5", "74,12", 7);
                tabelle.Rows.Add(6, 142536, "Produkt6", "5,36", 19);
                tabelle.Rows.Add(7, 123456, "Produkt7", "85,45", 19);
                tabelle.Rows.Add(8, 654321, "Produkt8", "41,85", 7);
                tabelle.Rows.Add(9, 654123, "Produkt9", "12,65", 19);
                tabelle.Rows.Add(10, 123654, "Produkt10", "52,58", 7);
                tabelle.Rows.Add(11, 415263, "Produkt11", "74,87", 19);
                tabelle.Rows.Add(12, 142536, "Produkt12", "5,41", 19);
                tabelle.Rows.Add(13, 123456, "Produkt13", "85,66", 7);
                tabelle.Rows.Add(14, 654321, "Produkt14", "41,55", 7);
                tabelle.Rows.Add(15, 654123, "Produkt16", "12,44", 19);
                tabelle.Rows.Add(16, 123654, "Produkt17", "52,33", 19);
                tabelle.Rows.Add(17, 415263, "Produkt18", "74,25", 19);
                tabelle.Rows.Add(18, 142536, "Produkt19", "5,14", 7);
                tabelle.Rows.Add(19, 654123, "Produkt20", "12,14", 19);
                /*     --> Hier haben wir die maximale Anzahl für die erste Seite an Einträgen ausgelastet
                 *       tabelle.Rows.Add(20, 123654, "Produkt21", "52,14", 7);
                 *       tabelle.Rows.Add(21, 415263, "Produkt22", "74,78", 19);
                 *       tabelle.Rows.Add(22, 142536, "Produkt23", "5,50", 7);
                 *       tabelle.Rows.Add(23, 123456, "Produkt24", "85,10", 19);
                 *       tabelle.Rows.Add(24, 654321, "Produkt25", "41,20", 7);
                 *       tabelle.Rows.Add(25, 654123, "Produkt26", "12,60", 19);
                 *       tabelle.Rows.Add(26, 123654, "Produkt", "52,85", 19);
                 *       tabelle.Rows.Add(27, 415263, "Produkt", "74,48", 7);
                 *       tabelle.Rows.Add(28, 142536, "Produkt", "5,85", 7);
                 *       tabelle.Rows.Add(29, 654123, "Produkt", "12,60", 19);
                 *       tabelle.Rows.Add(30, 123654, "Produkt", "52,85", 19);
                 *       tabelle.Rows.Add(31, 415263, "Produkt", "74,48", 7);
                 *       tabelle.Rows.Add(32, 142536, "Produkt", "5,85", 7);
                 */
                anzahlDatensätze = tabelle.Rows.Count; //Diese Zuweisung zählt die Datensätze und schreibt die Anzahl in eine Variable.
            }
                catch (Exception Fehler) {
                    MessageBox.Show("Kommunikation mit der Datenbank nicht möglich/fehlerhaft " + Fehler.Message);
            }
                        
        }

        private void ausdruck()
        {
            PrintDocument PrintDoc = new PrintDocument();

            // Hier wird der Druckdialog angebunden.
            // Der passende Drucker wird von dem Druckdialog ausgelesen und dementsprechend benutzt.
            PrintDoc.PrinterSettings.PrinterName = dialog.PrinterSettings.PrinterName;
            PrintDoc.PrinterSettings.Copies = dialog.PrinterSettings.Copies;
            PrintDoc.PrintPage += new PrintPageEventHandler(ersteSeite); //Aufruf der ersten Seite
            PrintDoc.Print();
        }

        void ersteSeite(object sender, PrintPageEventArgs e)
        {
            //!! Maximal 19 Einträge können gedruckt werden

            //Logo einfügen
            Image newImage = RechnungsGenerator.Properties.Resources.samplecustomer;
            Bitmap logo = new Bitmap(newImage);
            logo.SetResolution(400, 400);
            
            // Header
            e.Graphics.DrawString(Firmenname, new Font("Courier", 25), new SolidBrush(Color.Black), new Point(453, 25));
            e.Graphics.DrawString(Firmenslogan, new Font("Courier", 12), new SolidBrush(Color.Black), new Point(490, 65));
            e.Graphics.DrawImageUnscaled(logo, new Point(592, 105));
            e.Graphics.DrawString(Firmenname + "  -  " + Firmenstr + " " + Firmenhnr + "  -  " + Firmenplz +" " + Firmenort, new Font("Courier", 7, FontStyle.Underline), new SolidBrush(Color.DarkGreen), new Point(87, 170)); //30 - 225
           
            // Empfängeradresse
            e.Graphics.DrawString(Anrede, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, 195));
            e.Graphics.DrawString(Vorname + " " + Nachname, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, 210)); //+15 
            e.Graphics.DrawString(Straße + " " + Hausnummer, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, 225));//+15
            e.Graphics.DrawString(Postleitzahl + " " + Ort, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, 240));//+15
           
            //Absenderadresse
            e.Graphics.DrawString(Firmenstr + " " + Firmenhnr, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(590, 195));
            e.Graphics.DrawString(Firmenplz + " " + Firmenort, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(590, 210));
            e.Graphics.DrawString("Es schreibt Ihnen:", new Font("Courier", 10, FontStyle.Bold), new SolidBrush(Color.Black), new Point(470, 225));
            e.Graphics.DrawString(Sachbearbeiter, new Font("Courier", 10, FontStyle.Bold), new SolidBrush(Color.Black), new Point(590, 225));
            e.Graphics.DrawString("Tel:  " + Telefonnr, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(590, 240));
            e.Graphics.DrawString("Fax: " + Faxnr, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(590, 255));
            e.Graphics.DrawString(Mailaddr, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(590, 270));
           
            // Faltlinen - an ein standartisiertes Schema angepasst
            e.Graphics.DrawString("__", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(16, 383));
            e.Graphics.DrawString("____", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(16, 550));
            e.Graphics.DrawString("__", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(16, 779));

            //DEBUGWERTE
            //Diese Debugwerte dienen nur der Entwicklung und werden in der finalen Version verschwinden.
            if (debug)
            {

                // Set up string.
                string measureString = "_";
                Font stringFont = new Font("Courier", 10);

                // Measure string.
                SizeF stringSize1 = new SizeF();
                stringSize1 = e.Graphics.MeasureString(measureString, stringFont);

                // Draw string to screen.
                e.Graphics.DrawString(measureString, stringFont, Brushes.Black, new PointF(100 - stringSize1.Width, 100 - stringSize1.Height));
                e.Graphics.DrawString("W", new Font("Courier", 10), new SolidBrush(Color.Red), new PointF(100, 100)); // Test rechtsbündigkeit




                //Hilfslinien um Designanpassungen besser vor zu nehmen, da die Randlinien an ein standartisiertes Schema angepasst sind.
                e.Graphics.DrawString("X", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(822, 16)); // MAX vertikal ( X )
                e.Graphics.DrawString("X", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(16, 1156)); // MAX horizontal ( Y )
                e.Graphics.DrawLine(new Pen(Color.Black, 3), new Point(87, 1), new Point(87, 1500)); //links - senkrecht
                e.Graphics.DrawLine(new Pen(Color.Black, 3), new Point(735, 1), new Point(735, 1500));//rechts - senkrecht
                e.Graphics.DrawLine(new Pen(Color.Black, 3), new Point(1, 1131), new Point(822, 1131)); //unten - waagerecht
                e.Graphics.DrawLine(new Pen(Color.Black, 3), new Point(1, 25), new Point(822, 25)); //oben - waagerecht
                String debugtext = "T|E|S|T";
                debugtext = debugtext.Replace("|", Environment.NewLine);
            //  e.Graphics.DrawString(debugtext, new Font("Courier", 180, FontStyle.Bold), new SolidBrush(Color.Black), new Point(200, 0));
                e.Graphics.DrawString("TEST", new Font("Courier", 20, FontStyle.Bold), new SolidBrush(Color.Black), new Point(200, 260));

            }

           // Body der Rechnung
            e.Graphics.DrawString("Rechnungsnummer: " + Rechnungsnr, new Font("Courier", 10, FontStyle.Bold), new SolidBrush(Color.Black), new Point(87, 330));
            e.Graphics.DrawString("Rechnungsdatum: " + Rechnungsdatum.ToString("dd.MM.yyyy"), new Font("Courier", 10, FontStyle.Bold), new SolidBrush(Color.Black), new Point(542, 330));
           
             //Anschreiben
            String Anschreiben;
            Anschreiben = "Sehr geehrte(r) " + Anrede + " " + Nachname + ",||anbei erhalten Sie die Rechnung nummer " + Rechnungsnr + ".|Nachfolgend eine Detailauftellung Ihrer Bestellung:";
            Anschreiben = Anschreiben.Replace("|", Environment.NewLine);
            e.Graphics.DrawString(Anschreiben, new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, 400));

            //Start der Tabelle
            //Tabellenhöhe
            int Zeichenhöhe;
            int Zeichenhöhe_orig;
            Zeichenhöhe = 500;
            Zeichenhöhe_orig = Zeichenhöhe;

            //Obere Leiste der Tabelle
            e.Graphics.DrawString("Menge", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(87, Zeichenhöhe - 20));
            e.Graphics.DrawString("Artikelnummer", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(145, Zeichenhöhe - 20));
            e.Graphics.DrawString("Name", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(265, Zeichenhöhe - 20));
            e.Graphics.DrawString("Einzelpreis Netto", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(395, Zeichenhöhe - 20));
            e.Graphics.DrawString("Mwst", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(523, Zeichenhöhe - 20));
            e.Graphics.DrawString("Gesamtpreis Netto", new Font("Courier", 10), new SolidBrush(Color.Black), new Point(585, Zeichenhöhe - 20));
            e.Graphics.DrawLine(new Pen(Color.Black, 2), new Point(87, Zeichenhöhe), new Point(735, Zeichenhöhe)); // Trennung der Tabelle waagerecht
            //Folgende Reihenfolge wird in der Tabelle benutzt (vom Kunde vorgegeben): Menge, Artikelnummer, Name, Einzelpreis Netto, Mwst, Gesamtpreis netto

            //Abspaltungen der Tabelle waagerecht
            //Die Tabelle ist dynamisch nach anzahl der Bestellungen generiert
            tabellentiefe = Convert.ToInt32(Zeichenhöhe + (15.5f * anzahlDatensätze));
            
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(140, Zeichenhöhe), new Point(140, tabellentiefe));//nach Menge
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(260, Zeichenhöhe), new Point(260, tabellentiefe));//nach Artikelnummer
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(390, Zeichenhöhe), new Point(390, tabellentiefe));//nach Name
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(510, Zeichenhöhe), new Point(510, tabellentiefe));//nach Einzelpris Netto
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(570, Zeichenhöhe), new Point(570, tabellentiefe));//nach MWS
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(87, tabellentiefe), new Point(735, tabellentiefe));//Endlinie

            //Tabelle füllen
            foreach (DataRow row in tabelle.Rows)
            {
               Object item = row.ItemArray[0]; //Menge einfügen
               e.Graphics.DrawString(item.ToString(), new Font("Cournier", 10), new SolidBrush(Color.Black), new Point(92, Zeichenhöhe));
               Zeichenhöhe = Zeichenhöhe + 15;
            }


            Zeichenhöhe = Zeichenhöhe_orig; //Dieser Eintrag setzt den Cursor zum ausfüllen der Tabelle zurück
            foreach (DataRow row in tabelle.Rows)
            {
                Object item = row.ItemArray[1]; //Spalte Artikelnummer einfügen
                e.Graphics.DrawString(item.ToString(), new Font("Cournier", 10), new SolidBrush(Color.Black), new Point(145, Zeichenhöhe));
                Zeichenhöhe = Zeichenhöhe + 15;
            }


            Zeichenhöhe = Zeichenhöhe_orig; //Dieser Eintrag setzt den Cursor zum ausfüllen der Tabelle zurück
            foreach (DataRow row in tabelle.Rows)
            {
                Object item = row.ItemArray[2]; //Spalte Name einfügen
                e.Graphics.DrawString(item.ToString(), new Font("Cournier", 10), new SolidBrush(Color.Black), new Point(265, Zeichenhöhe));
                Zeichenhöhe = Zeichenhöhe + 15;
            }


            Zeichenhöhe = Zeichenhöhe_orig; //Dieser Eintrag setzt den Cursor zum ausfüllen der Tabelle zurück
            foreach (DataRow row in tabelle.Rows)
            {
                Object item = row.ItemArray[3]; //Spalte Einzelpreis Netto einfügen - rechtsbündig             
                string Eintrag = item.ToString() + " €"; //String zusammenfassen
                SizeF stringSize = new SizeF(); // Größe des Strings als Variable definieren
                stringSize = e.Graphics.MeasureString(Eintrag, new Font("Courier", 10)); //Größe des Strings in Variable schreiben
                e.Graphics.DrawString(Eintrag, new Font("Courier", 10), Brushes.Black, new PointF(506 - stringSize.Width, Zeichenhöhe)); //String schreiben minus die abgemessene Breite um eine künstliche Rechtsbündigkeit zu erzeugen
                Zeichenhöhe = Zeichenhöhe + 15; //Nächster Eintrag muss tiefer sein als dieser Eintrag
            }


            Zeichenhöhe = Zeichenhöhe_orig; //Dieser Eintrag setzt den Cursor zum ausfüllen der Tabelle zurück
            foreach (DataRow row in tabelle.Rows)
            {
                Object item = row.ItemArray[4]; //Spalte MwST einfügen
                string Eintrag = item.ToString() + " %";
                SizeF stringSize = new SizeF();
                stringSize = e.Graphics.MeasureString(Eintrag, new Font("Courier", 10));
                e.Graphics.DrawString(Eintrag, new Font("Courier", 10), Brushes.Black, new PointF(562 - stringSize.Width, Zeichenhöhe));
                Zeichenhöhe = Zeichenhöhe + 15;
                
            }


            Zeichenhöhe = Zeichenhöhe_orig; //Dieser Eintrag setzt den Cursor zum ausfüllen der Tabelle zurück
            foreach (DataRow row in tabelle.Rows)
            {
                //Hier wird die Produktsumme ermittelt
                    double menge = Convert.ToDouble(row.ItemArray[0]);
                    double preis = Convert.ToDouble(row.ItemArray[3]);
                    double produktsumme = menge * preis;
                
                string Eintrag = produktsumme.ToString("n2") + " €"; // Hier wird eine spezielle Geldformatierung ausgewählt. (n2)
                SizeF stringSize = new SizeF();
                stringSize = e.Graphics.MeasureString(Eintrag, new Font("Courier", 10));
                e.Graphics.DrawString(Eintrag, new Font("Courier", 10), Brushes.Black, new PointF(706 - stringSize.Width, Zeichenhöhe));
                Zeichenhöhe = Zeichenhöhe + 15;
                Zwischensumme = Zwischensumme + produktsumme; //Wir berechnen in dieser Schleife die Summe aller Artikel

                // - Berechnung MWST - Wir müssen in dieser Schleife die MWST berechnen, da nur hier die Produktsumme in Verbindung mit den verschiedenen MWST-Sätzen steht.
                Object mwstcheck = row.ItemArray[4]; //Auslesen der MWST des derzeitigen Produkts
                if (mwstcheck.ToString() == "19") //Abfrage, welcher Mehrwertsteuersatz genutzt wird. Kundenvorgabe sind 2 verschiedene Sätze.
                {
                    mwst19 = mwst19 + (produktsumme * 0.19d); //Hier wird die Mehrwertsteuer 19% des derzeitigen Produkts zu den anderen MWST-Sätzen addiert, wenn das Produkt 19% enthält.
                }
                if (mwstcheck.ToString() == "7")
                {

                    mwst7 = mwst7 + (produktsumme * 0.07d);//Hier wird die Mehrwertsteuer 7% des derzeitigen Produkts zu den anderen MWST-Sätzen addiert, wenn das Produkt 7% enthält.
                }

             }

            //Rechnungszusammenfassung

            // - Zwischensumme (unterKathegorie von Rechnungszusammenfassung)
            string Zwischensumme_string = Zwischensumme.ToString("n2") + " €"; //Hier wird die Zwischensumme als String formatiert
            SizeF Zwischensumme_setting = new SizeF(); //Eine größenangabe in der Variable
            Zwischensumme_setting = e.Graphics.MeasureString(Zwischensumme_string, new Font("Courier", 10));//Testen wie groß der String sein wird
            e.Graphics.DrawString(Zwischensumme_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - Zwischensumme_setting.Width, tabellentiefe +3));//An entsprechnder Tabelle unterhalb der Tabelle schreiben
            e.Graphics.DrawLine(new Pen(Color.Gray, 1), new Point(575, tabellentiefe + 20), new Point(735, tabellentiefe + 20));//Zwischensumme Linie
            e.Graphics.DrawString("Zwischensumme: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 3)); //Beschreibung der Zahlen

            // - Versandkosten
            string Versandkosten_string = Versandkosten.ToString("n2") + " €";
            SizeF Versandkosten_setting = new SizeF();
            Versandkosten_setting = e.Graphics.MeasureString(Versandkosten_string, new Font("Courier", 10));
            e.Graphics.DrawString(Versandkosten_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - Versandkosten_setting.Width, tabellentiefe + 20));
            e.Graphics.DrawLine(new Pen(Color.Gray, 1), new Point(575, tabellentiefe + 35), new Point(735, tabellentiefe + 35));//Versandkosten
            e.Graphics.DrawString("Versandkosten: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 20));

            // - MwSt19
            string mwst19_string = mwst19.ToString("n2") + " €";
            SizeF mwst19_setting = new SizeF();
            mwst19_setting = e.Graphics.MeasureString(mwst19_string, new Font("Courier", 10));
            e.Graphics.DrawString(mwst19_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - mwst19_setting.Width, tabellentiefe + 35));
            e.Graphics.DrawString("MwSt. 19%: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 35));

            // - MwSt7
            string mwst7_string = mwst7.ToString("n2") + " €";
            SizeF mwst7_setting = new SizeF();
            mwst7_setting = e.Graphics.MeasureString(mwst7_string, new Font("Courier", 10));
            e.Graphics.DrawString(mwst7_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - mwst7_setting.Width, tabellentiefe + 50));
            e.Graphics.DrawString("MwSt. 7%: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 50));

            // - MwSt gesamt ermitteln
            double mwstges = mwst7 + mwst19;
            string mwstges_string = mwstges.ToString("n2") + " €";
            SizeF mwstges_setting = new SizeF();
            mwstges_setting = e.Graphics.MeasureString(mwstges_string, new Font("Courier", 10));
            e.Graphics.DrawString(mwstges_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - mwstges_setting.Width, tabellentiefe + 65));
            e.Graphics.DrawString("MwSt. gesamt: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 65));
            e.Graphics.DrawLine(new Pen(Color.Gray, 1), new Point(575, tabellentiefe + 80), new Point(735, tabellentiefe + 80));//Linie über Summe


            // - Summe ermitteln
            double summe = Zwischensumme + Versandkosten + mwstges;
            string summe_string = summe.ToString("n2") + " €";
            SizeF summe_setting = new SizeF();
            summe_setting = e.Graphics.MeasureString(summe_string, new Font("Courier", 10));
            e.Graphics.DrawString(summe_string, new Font("Courier", 10, FontStyle.Bold), Brushes.Black, new PointF(706 - summe_setting.Width, tabellentiefe + 80));
            e.Graphics.DrawString("Summe: ", new Font("Courier", 10, FontStyle.Bold), Brushes.Black, new Point(265, tabellentiefe + 80));

            // - Rabatt ausrechnen & schreiben
           string rabatt_string = rabatt.ToString("n2") + " €";
            SizeF rabatt_setting = new SizeF();
            rabatt_setting = e.Graphics.MeasureString(rabatt_string, new Font("Courier", 10));
            e.Graphics.DrawString(rabatt_string, new Font("Courier", 10), Brushes.Black, new PointF(706 - rabatt_setting.Width, tabellentiefe + 95));
            e.Graphics.DrawString("Rabatt: ", new Font("Courier", 10), Brushes.Black, new Point(265, tabellentiefe + 95));
            e.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(575, tabellentiefe + 110), new Point(735, tabellentiefe + 110));//Linie über gesamt

            // - Gesamt
            double gesamt = summe - rabatt; //Gesamtbetrag ausrechnen
            string gesamt_string = gesamt.ToString("n2") + " €"; 
            SizeF gesamt_setting = new SizeF();
            gesamt_setting = e.Graphics.MeasureString(gesamt_string, new Font("Courier", 10));
            e.Graphics.DrawString(gesamt_string, new Font("Courier", 10, FontStyle.Bold), Brushes.Black, new PointF(706 - gesamt_setting.Width, tabellentiefe + 110));
            e.Graphics.DrawString("Gesamt: ", new Font("Courier", 10,FontStyle.Bold), Brushes.Black, new Point(265, tabellentiefe + 110));

             //Überweisungsaufforderung unter der Tabelle
            String überweisungsaufforderung = "Bitte überweisen Sie den Gesamtbetrag von " + gesamt_string +" innerhalb 7 Tagen |unter Angabe der Rechnungsnummer auf mein Konto. ||Vielen Dank!||" +Sachbearbeiter;
            überweisungsaufforderung = überweisungsaufforderung.Replace("|", Environment.NewLine);
            e.Graphics.DrawString(überweisungsaufforderung, new Font("Courier", 10), Brushes.Black, new Point(87, tabellentiefe + 190)); //Dynamisch gehalten, das der Ort dieses Textes je nach Anzahl der Produkte passend geschrieben wird.
           
            //Fußzeile
            String Fußzeile = "IBAN: " + IBAN +"       BIC: "+  BIC + "       Steuernummer: " + Steuernr;
            Fußzeile = Fußzeile.Replace("|", Environment.NewLine);
            e.Graphics.DrawLine(new Pen(Color.Gray, 1), new Point(87, 1100), new Point(735, 1100)); // Trennung der Fußzeile
            e.Graphics.DrawString(Fußzeile, new Font("Courier", 10), new SolidBrush(Color.Gray), new Point(87, 1110));
            }


    
    }
}
