using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace easytext
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e) //Undo
        {

        }

        private void toolStripButton2_Click(object sender, EventArgs e) //Redo
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e) //Rechercher
        {

        }

        private void toolStripButton4_Click(object sender, EventArgs e) //Centrer
        {

        }

        private void toolStripButton5_Click(object sender, EventArgs e) //Bold
        {

        }

        private void toolStripButton6_Click(object sender, EventArgs e) //Italic
        {

        }

        private void toolStripButton7_Click(object sender, EventArgs e) //Underline
        {

        }

        private void toolStripButton8_Click(object sender, EventArgs e) //Imprimer le doc Word
        {
            printDialog1.ShowDialog();
        }

        private void toolStripButton9_Click(object sender, EventArgs e) //Ouverture d'un doc Word
        {
            openFileDialog1.FileName = "document.docx";
            /*DialogResult DR = */openFileDialog1.ShowDialog();     //bloquante = modale
            Form1.ActiveForm.Text = "EasyText - document.docx"; //Changement du titre de de la fenêtre
        }

        private void toolStripButton10_Click(object sender, EventArgs e) //Enregister le doc Word
        {
            String contenu = richTextBox1.Text;

            string chemin="";

            saveFileDialog1.FileName = richTextBox1.Text + ".docx";
            DialogResult DR = saveFileDialog1.ShowDialog(); //bloquante  = modale


            if (DR.Equals(DialogResult.OK)) //bouton ENREGISTRER
            {
                chemin = saveFileDialog1.FileName;

                Microsoft.Office.Interop.Word.Application woApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document monDoc = woApp.WordBasic;
                //monDoc. = contenu;


                /*
                    Excel.Application xlApp = new Excel.Application(); //lancement d'excel
                    Excel.Workbook monclasseur = xlApp.Workbooks.Open(@"C:/Users/Adrien/Dev/C#/GestionnaireNotes/TPBulletinGroupeB2/res/modelebulletin.xlsx");//ds 'mes documents'
                    Excel.Worksheet mafeuille = monclasseur.Sheets[1];//selection de la premiere feuille = premier onglet
                    //remplissage du bulletin XLSX
                    mafeuille.Cells[1, 2] = textBox1.Text; //nom ds textBox1
                    monclasseur.SaveAs(chemin);
                    monclasseur.Close(true); 
                  */
            }

        }//Fin de la sauvegarde docx

        private void toolStripButton11_Click(object sender, EventArgs e) //Aide
        {
            //Charge form2
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //Police
        {
            //Font
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) //Taille
        {
            //Size
        }
    }
}
