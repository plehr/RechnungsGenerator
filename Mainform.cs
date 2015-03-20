using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RechnungsGenerator
{
    public partial class Mainform : Form
    {
        public Mainform()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Hier wird der Aufruf aus der GUI-Simuliert mit allen übergebenen Variablen/Parameter
            PrintDialog printDialog1 = new PrintDialog();
            DialogResult result = printDialog1.ShowDialog();
            if (result == DialogResult.OK)
            { //Prüfen, ob auf Abbrechen im Druckdialog gedrückt wurde
                int KDNR = 123456789;
                PDF pdfengine = new PDF();
                pdfengine.generieren(KDNR, printDialog1);
            }
        }
    }
}
