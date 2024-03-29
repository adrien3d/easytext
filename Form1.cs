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
        public static String font = "Calibri";
        public static float size = 12;

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

            richTextBox1.Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Cras egestas ac urna vitae tristique. Phasellus faucibus mi nisi, id sollicitudin elit sodales vel. Praesent suscipit nunc dignissim dui rhoncus, finibus mollis ex pharetra. Maecenas in urna ut magna volutpat elementum id eget ante. Proin eget sagittis nibh. Nulla facilisi. Suspendisse nec quam non risus vestibulum tempor a elementum nunc. Nulla ornare justo sed auctor maximus. Duis commodo pulvinar dolor, nec varius odio tincidunt non. Aliquam hendrerit tincidunt mauris, id pulvinar purus ultricies at. Morbi sodales sapien rutrum pretium vehicula. Sed a enim lacus. Integer tincidunt metus porttitor tristique porta. Donec non facilisis ex, vitae tempor mauris. ";
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
            if (richTextBox1.CanUndo) toolStripButton1.Enabled = true;
            else toolStripButton1.Enabled = false;

            if (richTextBox1.CanRedo) toolStripButton2.Enabled = true;
            else toolStripButton2.Enabled = false;

        }

        private void toolStripButton1_Click(object sender, EventArgs e) //Undo
        {
            richTextBox1.Undo();
        }

        private void toolStripButton2_Click(object sender, EventArgs e) //Redo
        {
            richTextBox1.Redo();
        }

        private void toolStripButton3_Click(object sender, EventArgs e) //Rechercher
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void toolStripButton4_Click(object sender, EventArgs e) //Centrer
        {
            
            if (richTextBox1.SelectedText == "") richTextBox1.SelectAll();
            
            if (richTextBox1.SelectionAlignment == HorizontalAlignment.Left) richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
            else richTextBox1.SelectionAlignment = HorizontalAlignment.Left;

            richTextBox1.DeselectAll();
        }

        private void toolStripButton5_Click(object sender, EventArgs e) //Bold
        {
            if (richTextBox1.SelectionFont != null)
            {
                System.Drawing.Font currentFont = richTextBox1.SelectionFont;
                System.Drawing.FontStyle newFontStyle;

                if (richTextBox1.SelectionFont.Bold == true)
                {
                    newFontStyle = ~FontStyle.Bold;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle & currentFont.Style);
                }
                else
                {
                    newFontStyle = FontStyle.Bold;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle | currentFont.Style);
                }

            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e) //Italic
        {
            if (richTextBox1.SelectionFont != null)
            {
                System.Drawing.Font currentFont = richTextBox1.SelectionFont;
                System.Drawing.FontStyle newFontStyle;

                if (richTextBox1.SelectionFont.Italic == true)
                {
                    newFontStyle = ~FontStyle.Italic;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle & currentFont.Style);
                }
                else
                {
                    newFontStyle = FontStyle.Italic;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle | currentFont.Style);
                }

            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e) //Underline
        {
            if (richTextBox1.SelectionFont != null)
            {
                System.Drawing.Font currentFont = richTextBox1.SelectionFont;
                System.Drawing.FontStyle newFontStyle;

                if (richTextBox1.SelectionFont.Underline == true)
                {
                    newFontStyle = ~FontStyle.Underline;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle & currentFont.Style);
                }
                else
                {
                    newFontStyle = FontStyle.Underline;
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle | currentFont.Style);
                }

                
            }
        }

        private void toolStripButton8_Click(object sender, EventArgs e) //Imprimer le doc Word
        {
            printDialog1.ShowDialog();
        }

        private void toolStripButton9_Click(object sender, EventArgs e) //Ouverture d'un doc Word
        {
            openFileDialog1.Filter = "Tous les Fichiers (*.*)|*.*|Fichiers Texte (.txt)|*.txt|Fichiers Word (.docx)|*.docx|Fichiers RTF (.rtf)|*.rtf";
            
            if (openFileDialog1.ShowDialog()==DialogResult.OK) {
                richTextBox1.Text = File.ReadAllText(openFileDialog1.FileName);
                Form1.ActiveForm.Text = openFileDialog1.FileName; //Changement du titre de de la fenêtre
            }
        }

        private void toolStripButton10_Click(object sender, EventArgs e) //Enregister le doc Word
        {
            saveFileDialog1.Filter = "Tous les Fichiers (*.*)|*.*|Fichiers Texte (.txt)|*.txt|Fichiers Word (.docx)|*.docx|Fichiers RTF (.rtf)|*.rtf";
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Word.Application msWord = new Microsoft.Office.Interop.Word.Application();
                msWord.Visible = false; // mettez cette variable à true si vous souhaitez visualiser les opérations.
                object missing = System.Reflection.Missing.Value;

                string chemin = saveFileDialog1.FileName;
                Microsoft.Office.Interop.Word.Document monDoc = msWord.WordBasic;
               /* monDoc = richTextBox1;
                monDoc.SaveAs(chemin);
                monDoc.Close(true);*/

                  try
                  {
                      RichTextBoxStreamType format = RichTextBoxStreamType.PlainText;
                      if (saveFileDialog1.FilterIndex == 1)
                      format = RichTextBoxStreamType.RichText;
                      richTextBox1.SaveFile(chemin, format);
                      MessageBox.Show("Fichier bien sauvegardé");
                  }
                  catch (Exception argh)
                  {
                      MessageBox.Show(argh.Message, "Erreur à la sauvegarde", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  }
            }

            /*monDoc = msWord.Documents.Add();
            object fileName = @"Mon nouveau document.docx";
            monDoc.SaveAs(chemin);
            monDoc.Close(true);*/

            /*StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
            sw.Write(richTextBox1.Text);
            sw.Close();
            MessageBox.Show("Fichier bien sauvegardé");*/

            //richTextBox1.SaveFile(@"C:\Users\pascal\Google Drive\IG2I\L2\IHM - Interface Homme Machine\test.rtf", RichTextBoxStreamType.RichText);


        }//Fin de la sauvegarde docx

        private void toolStripButton11_Click(object sender, EventArgs e) //Aide
        {
            //Charge form3
        }

        //button ou combobox couleur
        /*this.richTextBox1.SelectAll();
            this.richTextBox1.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(Colors.Blue));
            this.richTextBox1.TextChanged += new TextChangedEventHandler(richTextBox1_TextChanged);
          thiq.richTextBox1.DeselectAll()*/
        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) //Police
        {
            if (richTextBox1.SelectedText == "") richTextBox1.SelectAll();

            font = comboBox1.SelectedItem.ToString();

            richTextBox1.SelectionFont = new Font(font, size);

            richTextBox1.DeselectAll();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) //Taille
        {
            if (richTextBox1.SelectedText == "") richTextBox1.SelectAll();

            String taillechoisie = comboBox2.SelectedItem.ToString();

            if (taillechoisie == "Petit") size = 8;
            if (taillechoisie == "Normal") size = 12;
            if (taillechoisie == "Grand") size = 18;
            if (taillechoisie == "Très grand") size = 22;
            
            richTextBox1.SelectionFont = new Font(font, size);
            
            richTextBox1.DeselectAll();



            /*if (richTextBox1.SelectedText == "") richTextBox1.SelectAll();

            if (richTextBox1.SelectionAlignment == HorizontalAlignment.Left) richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
            else richTextBox1.SelectionAlignment = HorizontalAlignment.Left;

            richTextBox1.DeselectAll();*/
        }
    }
}
