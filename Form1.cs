﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace easytext
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            toolStripButton1.Enabled = false;
            toolStripButton2.Enabled = false;
            if (richTextBox1.CanUndo)       toolStripButton1.Enabled = true;
            if (richTextBox1.CanRedo)       toolStripButton2.Enabled = true;
            comboBox1.Items.Add("Arial");
            comboBox1.Items.Add("Calibri");
            comboBox1.Items.Add("Comic Sans MS");
            comboBox1.Items.Add("Garamond");
            comboBox1.Items.Add("Times New Roman");

            comboBox2.Items.Add("Petit");
            comboBox2.Items.Add("Normal");
            comboBox2.Items.Add("Grand");
            comboBox2.Items.Add("Très grand");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            /*this.richTextBox1.Selection.Select(this.richTextBox1.Selection.Start.GetPositionAtOffset(-1), this.richTextBox1.Document.ContentEnd);
            this.richTextBox1.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.Red));
            this.richTextBox1.TextChanged -= new TextChangedEventHandler(richTextBox1_TextChanged);
            this.richTextBox1.Selection.Select(this.richTextBox1.Document.ContentEnd, this.richTextBox1.Document.ContentEnd);*/
        }

        private void toolStripButton1_Click(object sender, EventArgs e) //Undo
        {

        }

        private void toolStripButton2_Click(object sender, EventArgs e) //Redo
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e) //Rechercher
        {
            //Cahrge form2
        }

        int sel = 0;
        private void toolStripButton4_Click(object sender, EventArgs e) //Centrer
        {
            richTextBox1.SelectAll();

            if ((sel%2) ==0) richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
            else richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
            sel++;

            richTextBox1.DeselectAll();
        }

        private void toolStripButton5_Click(object sender, EventArgs e) //Bold
        {
            richTextBox1.SelectAll();
            //richTextBox1.SelectionColor = Color.Blue;
            richTextBox1.DeselectAll();
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
            openFileDialog1.Filter = "Fichiers texte (.txt)|*.txt|Fichiers Word (.docx)|*.docx|Tous les Fichiers (*.*)|*.*";
            
            if (openFileDialog1.ShowDialog()==DialogResult.OK) {
                richTextBox1.Text = File.ReadAllText(openFileDialog1.FileName);
                Form1.ActiveForm.Text = openFileDialog1.FileName; //Changement du titre de de la fenêtre
            }
        }

        private void toolStripButton10_Click(object sender, EventArgs e) //Enregister le doc Word
        {
            saveFileDialog1.Filter = "Fichier texte (*.txt)|*.txt";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                  try
                  {
                      StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
                      sw.Write(richTextBox1.Text);
                      sw.Close();
                      MessageBox.Show("Fichier bien sauvegardé");
                  }
                  catch (Exception argh)
                  {
                      MessageBox.Show(argh.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  }
            }
            /*string chemin="";

            saveFileDialog1.FileName = richTextBox1.Text + ".docx";
            DialogResult DR = saveFileDialog1.ShowDialog(); //bloquante  = modale


            if (DR.Equals(DialogResult.OK)) //bouton ENREGISTRER
            {
                chemin = saveFileDialog1.FileName;

                Microsoft.Office.Interop.Word.Application woApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document monDoc = woApp.WordBasic;
                monDoc = richTextBox1.Text;
                mondoc.SaveAs(chemin);
                mondoc.Close(true);
            }*/

        }//Fin de la sauvegarde docx

        private void toolStripButton11_Click(object sender, EventArgs e) //Aide
        {
            //Charge form3
        }

        //button ou combobox couleur
        /*this.richTextBox1.SelectAll();
            this.richTextBox1.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.Blue));
            this.richTextBox1.TextChanged += new TextChangedEventHandler(richTextBox1_TextChanged);*/

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //Police
        {
            richTextBox1.SelectAll();

            String font = comboBox1.SelectedItem.ToString();

            richTextBox1.SelectionFont = new Font(font, 12);

            richTextBox1.DeselectAll();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) //Taille
        {
            richTextBox1.SelectAll();

            String font = comboBox1.SelectedItem.ToString();
            String size = comboBox2.SelectedItem.ToString();

            if (size == "Petit")     richTextBox1.SelectionFont = new Font(font, 8);
            if (size == "Normal")     richTextBox1.SelectionFont = new Font(font, 12);
            if (size == "Grand")     richTextBox1.SelectionFont = new Font(font, 18);
            if (size == "Très grand")     richTextBox1.SelectionFont = new Font(font, 22);
            
            richTextBox1.DeselectAll();
        }
    }
}
