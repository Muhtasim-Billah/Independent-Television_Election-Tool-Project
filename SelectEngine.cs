using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ElectionResult
{
    public partial class FormSelectEngine : Form
    {
        private FormElectionResult fel;
        private bool isOkayButtonPressed = false;
        private string currentEngine;
        private bool isBugEngine;
        public FormSelectEngine(FormElectionResult fel, String currentEngine, bool isBugEngine)
        {

            InitializeComponent();
            this.fel = fel;
            this.currentEngine = currentEngine;
            this.isBugEngine = isBugEngine;
        }

        /// <summary>
        /// Ok button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            this.isOkayButtonPressed = true;
            this.Close();
        }

        /// <summary>
        /// cancel button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// form is closing, if ok button is pressed please send new engine, else send present engine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormSelectEngine_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.isOkayButtonPressed)
            {
                if (!this.isBugEngine)
                {
                    this.fel.receiveChildCommand(0, true, comboBox1.Text);
                }
                else
                    this.fel.receiveChildCommand(1, true, comboBox1.Text);
            }
            else
                this.fel.receiveChildCommand(0, true, this.currentEngine);
        }
    }
}
