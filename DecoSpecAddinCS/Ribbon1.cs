using System;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace DecoSpecAddinCS
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Run("showuserform");
            }
            catch (Exception)
            {
                MessageBox.Show("This is not a Decoration Specification!");
            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Run("ClearAll");
            }
            catch (Exception)
            {
                MessageBox.Show("This is not a Decoration Specification!");
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Documents.Add(@"P:\decoration_specification\Template\DecoSpecTemplate.dotm");
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.Run("SaveDeco");

            }
            catch (Exception)
            {
                MessageBox.Show("This is not a Decoration Specification!");
            }

        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {
            var box = new AboutBox1();
            box.Show();
        }
    }
}
