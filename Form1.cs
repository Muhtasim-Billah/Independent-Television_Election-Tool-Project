/// <summary>
/// Author: Muhtasim Billah
/// WARNING: This software sends messages to rendering engines. You must be careful while running it.
/// </summary>
//Note: For this version some changes made here.
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;//for interacting with excel file
using System.Threading;


namespace ElectionResult
{
    
    public partial class FormElectionResult : Form
    {
        private string currentFFEngine;//current full frame engine
        private string currentBUGEngine;//current BUG engine
        private string backScene = "SCENE......./Comilla/BG";//background scene
        private string mainTitle = "Election Result";//title to be shown on title bar
        private int currentFFScene = 0;
        private int currentBUGScene = 0;
        private vizCommunication vc;//library used for connection. here removed now
        int t1, t2, t3, t4, t5, t6, n1, n2, p1, p2, p3, p4, p5, p6;
        private string scene1, scene2, scene3, scene4, scene5, scene6, scene7, scene8, scene9, scene10, scene11, scene12, scene13, scene14, scene15, scene16, scene17;
        private Thread executorThread;
        private Executor executor;
        delegate void SetUpdateCallback(int commandType);

        /// <summary>
        /// constructor class
        /// </summary>
        public FormElectionResult()
        {
            InitializeComponent();
            this.initialize();
        }
        /// <summary>
        /// make some initialization, called from constructor
        /// </summary>
        private void initialize()
        {
            this.currentFFEngine = "NONE";
            this.currentBUGEngine = "NONE";
            this.Text = this.mainTitle + " Engine: " + this.currentFFEngine;
            this.vc = new vizCommunication();
            vc.hostname = this.currentFFEngine;

            /// SCENE NAMES
            scene1 = "SCENE NAME";
            scene2 = "SCENE Name";
            scene3 = "SCENE NAme";
            scene4 = "SCENE Name";
            scene5 = "SCENE Name";
            scene6 = "SCENE Name";
            scene7 = "SCENE Name";
            scene8 = "SCENE Name";
            scene9 = "SCENE Name";
            scene10 = "SCENE Name";
            scene11 = "SCENE Name";//"";
            scene12 = "SCENE Name";
            scene13 = "SCENE Name";
            scene14 = "SCENE Name";
            scene15 = "SCENE Name";
            scene16 = "SCENE Name";
            scene17 = "SCENE Name";
            ///

            this.showText(2);//shows welcome message
            ///Start thread
            executor = new Executor(this);
            executorThread = new Thread(executor.executorMethod);
            executorThread.Start();//starts the thread for sending auto update if necessary

        }

        /// <summary>
        /// is called from other forms (like "Engine select form") for sending message back to father form about which scene is selected
        /// </summary>
        /// <param name="commandType">0 = from FFengineSetting panel, paramBool = setEnable, paramString1 = engineSetting,  1 = same but BUGEngine</param>
        /// <param name="paramBool"></param>
        public void receiveChildCommand(int commandType, bool paramBool, string paramString1)
        {
            switch(commandType)
            {
                case 0://update full frame engine
                    this.setThisFormEnableOrDisable(paramBool);
                    this.currentFFEngine = paramString1;
                    vc.hostname = this.currentFFEngine;
                    this.Text = this.mainTitle + " Full Frame Engine: " + this.currentFFEngine + ", BUG Engine: " + this.currentBUGEngine;
                    break;
                case 1://update bug engine
                    this.setThisFormEnableOrDisable(paramBool);
                    this.currentBUGEngine = paramString1;
                    this.Text = this.mainTitle + " Full Frame Engine: " + this.currentFFEngine + ", BUG Engine: " + this.currentBUGEngine;
                    break;
            }
        }

        private void sendVizCommand(String s)
        {
            //MessageBox.Show(s);
            this.vc.singleCommandSendNoBufferUseDefaults(s);
        }

        /// <summary>
        /// enables or disables main form. needed when a child form is created or closed.
        /// </summary>
        /// <param name="isEnable"></param>
        private void setThisFormEnableOrDisable(bool isEnable)
        {
            this.Enabled = isEnable;
        }

        /// <summary>
        /// Disables current engine and opens pop up to select new engine form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectEngine_Click(object sender, EventArgs e)
        {
            this.setThisFormEnableOrDisable(false);
            FormSelectEngine fse = new FormSelectEngine(this, this.currentFFEngine,false);
            fse.Show();
        }
        /// <summary>
        /// Disables current engine and opens pop up to select new engine form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelectBUGEngine_Click(object sender, EventArgs e)
        {
            this.setThisFormEnableOrDisable(false);
            FormSelectEngine fse = new FormSelectEngine(this, this.currentFFEngine, true);
            fse.Show();
        }

        /// <summary>
        /// Cleans bug or FF engine depending on the parameter
        /// </summary>
        /// <param name="isBugEngine">if true bug engine is cleared, otherwise ff engine is cleared</param>
        private void cleanUpEngine(bool isBugEngine)
        {

            if (isBugEngine)
            {
                vc.hostname = this.currentBUGEngine;
            }
            this.sendVizCommand("0 RENDERER*LAYER SET_OBJECT");
            //More Commands
            this.sendVizCommand("0 SCENE CLEANUP");
            //More Commands
            this.sendVizCommand("0 MATERIAL CLEANUP");
            this.sendVizCommand("0 MAPS CLEANUP");
            //**uncheck all rad button, disable play button, stop button
                
            if (isBugEngine)
            {
                vc.hostname = this.currentFFEngine;
            }
        }


        /// <summary>
        /// cleans up full frame engine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCleanup_Click(object sender, EventArgs e)
        {
            this.cleanUpEngine(false);
        }
        /// <summary>
        /// changes current scene
        /// </summary>
        /// <param name="sceneName">the scene name to be changed into</param>
        /// <param name="isBUGScene">says wheter it is FF or BUG scene</param>
        private void changeScene(string sceneName, bool isBUGScene)
        {
            if (this.currentFFEngine == "NONE" || this.currentBUGEngine == "NONE")
            {
                MessageBox.Show("You must choose both Full Frame engine and BUG engine first.", "Engine not selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.btnPlay.Enabled = false;
            //check if current scene should be fired in BUG Engine
            if (isBUGScene)
                    vc.hostname = this.currentBUGEngine;
            
            this.sendVizCommand("0 RENDERER SET_OBJECT " + sceneName);         
            this.sendVizCommand("0 RENDERER SET_OBJECT " + sceneName + " LOAD");
            vc.hostname = this.currentFFEngine;

            //this.sendVizCommand("0 RENDERER*STAGE SHOW 0.0");
            this.sendUpdateCommand();

            if (isBUGScene)
                vc.hostname = this.currentBUGEngine;
            this.sendVizCommand("-1 RENDERER*STAGE START");//start initially
            vc.hostname = this.currentFFEngine;
            this.btnPlay.Enabled = true;
            this.btnStop.Enabled = true;
        }
        /// <summary>
        /// plays the scene
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPlay_Click_1(object sender, EventArgs e)
        {
            this.sendVizCommand("-1 RENDERER*STAGE START");
        }

        /// <summary>
        /// stops the scene
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStop_Click(object sender, EventArgs e)
        {
            this.sendVizCommand("-1 RENDERER*STAGE STOP");
        }
        
        
        ///////////////////////////////////////////////////RADIO BUTTON ACTION LISTENER FOR SCENE CHANGES START/////////////////////////////
        
        private void radBtn1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn1.Checked)
            {
                this.currentFFScene = 1;
                this.changeScene(this.scene1,false);
            }
        }

        private void radBtn2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn2.Checked)
            {
                this.currentFFScene = 0;
                this.changeScene(this.scene2,false);
            }
        }
        private void radBtn3_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn3.Checked)
            {
                this.currentFFScene = 0;
                this.changeScene(this.scene3,false);
            }
        }
        private void radBtn4_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn4.Checked)
            {
                this.currentFFScene = 4;
                this.changeScene(this.scene4,false);
            }
        }

        private void radBtn5_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn5.Checked)
            {
                this.currentFFScene = 5;
                this.changeScene(this.scene5,false);
            }
        }

        private void radBtn6_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn6.Checked)
            {
                this.currentFFScene = 6;
                this.changeScene(this.scene6,false);
            }
        }

        private void radBtn7_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn7.Checked)
            {
                this.currentFFScene = 7;
                this.changeScene(this.scene7,false);
            }
        }

        private void radBtn8_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn8.Checked)
            {
                this.currentFFScene = 8;
                this.changeScene(this.scene8,false);
            }
        }
        private void radBtn9_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn9.Checked)
            {
                this.currentFFScene = 9;
                this.changeScene(this.scene9,false);
            }
        }
        private void radBtn10_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn10.Checked)
            {
                this.currentBUGScene = 10;
                this.changeScene(this.scene10,true);
            }
        }
        private void radBtn11_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn11.Checked)
            {
                this.currentBUGScene = 11;
                this.changeScene(this.scene11, true);
            }
        }
        /// <summary>
        /// 3 candidates vote
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radBtn12_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn12.Checked)
            {
                this.currentFFScene = 12;
                this.changeScene(this.scene12,false);
            }
        }

        /// <summary>
        /// 3 Candi pie chart
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radBtn13_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn13.Checked)
            {
                this.currentFFScene = 13;
                this.changeScene(this.scene13, false);
            }
        }
        /// <summary>
        /// 2 candi %
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radBtn14_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn14.Checked)
            {
                this.currentFFScene = 14;
                this.changeScene(this.scene14,false);
            }
        }
        /// <summary>
        /// 2 candi vote
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radBtn15_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn15.Checked)
            {
                this.currentFFScene = 15;
                this.changeScene(this.scene15,false);
            }
        }
        /// <summary>
        /// 2 candi pie 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radBtn16_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn16.Checked)
            {
                this.currentFFScene = 16;
                this.changeScene(this.scene16,false);
            }
        }
        private void radBtn17_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radBtn17.Checked)
            {
                this.currentBUGScene = 17;
                this.changeScene(this.scene17, true);
            }
        }
        ///////////////////////////////////////////////////RADIO BUTTON ACTION LISTENER FOR SCENE CHANGES END/////////////////////////////



        /// <summary>
        /// update from excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            this.updateFromExcel();
        }

        /// <summary>
        /// sends data pool updates to viz Full frame engine. different scene needed differend parameters, they are selected by switch structure
        /// </summary>
        private void sendUpdateCommand()
        {
            string command;
            switch(currentFFScene)
            {
                //Some changes made here for this version.
                case 0://normal update                
                    command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";t3=" + this.t3.ToString() + ";t4=" + this.t4.ToString() + ";t5=" + this.t5.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";p4=" + this.p4.ToString() + ";p5=" + this.p5.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
            
                case 9://give pie values for all
                    this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d1=" + p1.ToString() + "#" + p2.ToString() + "#" + p3.ToString() + "#" + p4.ToString() + "#" + p5.ToString());
                    command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";t3=" + this.t3.ToString() + ";t4=" + this.t4.ToString() + ";t5=" + this.t5.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";p4=" + this.p4.ToString() + ";p5=" + this.p5.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 2:
                    break;
                case 4://afzol single
                    command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";p1=" + this.p1.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 5://mithu single
                    command = "0 RENDERER*FUNCTION*Data SET t5=" + this.t5.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";p5=" + this.p5.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 6: 
                    //sakku single
                    command = "0 RENDERER*FUNCTION*Data SET t2=" + this.t2.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";p2=" + this.p2.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 7:
                    //Eyar Ahmed Selim Single
                    command = "0 RENDERER*FUNCTION*Data SET t3=" + this.t3.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";p3=" + this.p3.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 8:
                    //Nur ur Rahman Mahmud
                    command = "0 RENDERER*FUNCTION*Data SET t4=" + this.t4.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";p4=" + this.p4.ToString() + ";";
                    this.sendVizCommand(command);
                    break;

                case 1://3 candidates%
                    command = "0 RENDERER*FUNCTION*Data SET p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 12:
                    //3 candidate vote
                    command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";t3=" + this.t3.ToString() +  ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";";                   
                    this.sendVizCommand(command);
                    break;
                case 13:
                    //3 candidate pie
                    this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d2=" + p1.ToString() + "#" + p2.ToString() + "#" + p3.ToString() + "#" + (p4+p5).ToString()) ;
                    command = "0 RENDERER*FUNCTION*DataPool*Data SET n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString()+";";
                    this.sendVizCommand(command);
                    break;

                case 14://2 candidates%
                    command = "0 RENDERER*FUNCTION*Data SET p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
                case 15:
                    //2 candidate vote
                    command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() +";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString()  + ";";
                    this.sendVizCommand(command);
                    break;
                case 16:
                    //2 candidate pie
                    this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d3=" + p1.ToString() + "#" + p2.ToString() + "#" +  (p3 + p4 + p5).ToString());
                    command = "0 RENDERER*FUNCTIONl*Data SET n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";";
                    this.sendVizCommand(command);
                    break;
            }
            this.sendUpdateCommandToBUGEngine();//ff command sending part is over. now send command to bug engine
        }
        /// <summary>
        /// sends update to bug engine only
        /// </summary>
        private void sendUpdateCommandToBUGEngine()
        {
            string command;
            if(chkBxEnableBUG.Checked)
            {
                vc.hostname = this.currentBUGEngine;
                switch (this.currentBUGScene)
                {
                    case 10: //bug all
                        command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";t3=" + this.t3.ToString() + ";t4=" + this.t4.ToString() + ";t5=" + this.t5.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";p4=" + this.p4.ToString() + ";p5=" + this.p5.ToString() + ";";
                        this.sendVizCommand(command);
                        this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d1=" + p1.ToString() + "#" + p2.ToString() + "#" + p3.ToString() + "#" + p4.ToString() + "#" + p5.ToString());
                        break;
                    case 11:
                        //bug 3 candidate
                        command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";t3=" + this.t3.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";p3=" + this.p3.ToString() + ";";
                        this.sendVizCommand(command);
                        //this.sendVizCommand("0 RENDERER*FUNCTION*DataPool*Data SET d2=" + p1.ToString() + "#" + p2.ToString() + "#" + p3.ToString() + "#" + (100-p1-p2-p3).ToString());
                        this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d2=" + p1.ToString() + "#" + p2.ToString() + "#" + p3.ToString()) ;//**now others will not be present
                        break;
                    case 17:
                        //bug 2 candidate
                        command = "0 RENDERER*FUNCTION*Data SET t1=" + this.t1.ToString() + ";t2=" + this.t2.ToString() + ";n1=" + this.n1.ToString() + ";n2=" + this.n2.ToString() + ";p1=" + this.p1.ToString() + ";p2=" + this.p2.ToString() + ";";
                        this.sendVizCommand(command);
                        this.sendVizCommand("0 RENDERER*FUNCTION*Data SET d3=" + p1.ToString() + "#" + p2.ToString() + "#" + (100-p1-p2).ToString());
                        break;

                }
                vc.hostname = this.currentFFEngine;//change hostname to default i.e. full frame engine
            }
        }
        /// <summary>
        /// sets scene background in ff engine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetBackground_Click(object sender, EventArgs e)
        {
            this.sendVizCommand("0 RENDERER*LAYER SET_OBJECT " + backScene);
            this.sendVizCommand("0 RENDERER*STAGE*bg START");
        }

        /// <summary>
        /// releases the objects that were created by excel operation
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// starts bug engine selection process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonBUGFF_Click(object sender, EventArgs e)
        {
            this.setThisFormEnableOrDisable(false);
            FormSelectEngine fse = new FormSelectEngine(this, this.currentFFEngine, true);
            fse.Show();
        }
        /// <summary>
        /// cleans up bug engine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBUGEngineCleanup_Click(object sender, EventArgs e)
        {
            this.cleanUpEngine(true);
        }

        /// <summary>
        /// updates the result from Excel file, location is hardcoded here
        /// </summary>
        /// <returns></returns>
        private bool updateFromExcel()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;//=new Excel.Workbook();
            Excel.Worksheet xlWorkSheet;// = new Excel.Worksheet();
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open("Q:\\Election Data Update.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //MessageBox.Show(xlWorkSheet.get_Range("A2", "A2").Value2.ToString());

                /////////////////////////
                t1 = int.Parse(xlWorkSheet.get_Range("A2", "A2").Value2.ToString());
                t2 = int.Parse(xlWorkSheet.get_Range("B2", "B2").Value2.ToString());
                t3 = int.Parse(xlWorkSheet.get_Range("C2", "C2").Value2.ToString());
                t4 = int.Parse(xlWorkSheet.get_Range("D2", "D2").Value2.ToString());
                t5 = int.Parse(xlWorkSheet.get_Range("E2", "E2").Value2.ToString());
                t6 = int.Parse(xlWorkSheet.get_Range("F2", "F2").Value2.ToString());
                n1 = int.Parse(xlWorkSheet.get_Range("G2", "G2").Value2.ToString());
                n2 = int.Parse(xlWorkSheet.get_Range("H2", "H2").Value2.ToString());
                int total = t1 + t2 + t3 + t4 + t5 + t6;
                p1 = (int)Math.Round(((double)t1 / total * 100));
                p2 = (int)Math.Round(((double)t2 / total * 100));
                p3 = (int)Math.Round(((double)t3 / total * 100));
                p4 = (int)Math.Round(((double)t4 / total * 100));
                p5 = (int)Math.Round(((double)t5 / total * 100));
                p6 = (int)Math.Round(((double)t6 / total * 100));
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
                // MessageBox.Show("Successfully updated data from file", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
            this.sendUpdateCommand();
            return true;
        }
        /// <summary>
        /// this method safely sends command within running threads
        /// </summary>
        /// <param name="commandType">0 = update excel, 1 = clear recent status</param>
        public void threadSafeCommand(int commandType)
        {
            if (this.rtbOutput.InvokeRequired)
            {
                SetUpdateCallback d = new SetUpdateCallback(threadSafeCommand);
                this.Invoke(d, new object[] { commandType });
            }
            else
            {
                switch (commandType)
                {
                    case 0:
                        if (this.updateFromExcel())
                        {
                            this.showText(0);
                            //if BUG engine is 
                        }
                        else
                        {
                            this.showText(1);
                        }

                        break;
                    case 1:
                        this.showText(3);
                        break;
                }                
            }
        }
        /// <summary>
        /// when the software is closing, stop the other thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormElectionResult_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (executorThread != null)
                if (!(executorThread.ThreadState == ThreadState.Stopped) || !(executorThread.ThreadState == ThreadState.Unstarted))
                    executorThread.Abort();
        }


        /// <summary>
        /// shows status texts in console
        /// </summary>
        /// <param name="type">0=success update, 1=failedUpdate, 2=welcome, 3=clearText</param>
        private void showText(int type)
        {
            int len;
            string str;
            switch (type)
            {
                case 0://successful update
                            len = rtbOutput.TextLength;
                            str = System.DateTime.Now.ToString("ddMMMyyyy_HHmm") + ": ";
                            this.rtbOutput.AppendText(str);
                            this.rtbOutput.Select(len, str.Length);
                            this.rtbOutput.SelectionColor = Color.Green;
                            this.rtbOutput.DeselectAll();
                            len = rtbOutput.TextLength;
                            str = " Successfully updated." + Environment.NewLine;
                            this.rtbOutput.AppendText(str);
                            this.rtbOutput.Select(len, str.Length);
                            this.rtbOutput.SelectionColor = Color.Black;
                            this.rtbOutput.DeselectAll();
                    break;
                case 1://failed to update
                            len = rtbOutput.TextLength;
                            str = System.DateTime.Now.ToString("ddMMMyyyy_HHmm") + ": FATAL ERROR: Could not fetch data from Excel file." + Environment.NewLine;
                            this.rtbOutput.AppendText(str);
                            this.rtbOutput.Select(len, str.Length);
                            this.rtbOutput.SelectionColor = Color.Red;
                            this.rtbOutput.DeselectAll();
                    break;
                case 2:
                            len = rtbOutput.TextLength;
                            str = "--ELECTION RESULT GFX CONTROLLER, VERSION 1.0, DEVELOPED BY ARIF KHAN & ITV GFX TEAM, SELECT ENGINES FIRST--" + Environment.NewLine;
                            this.rtbOutput.AppendText(str);
                            this.rtbOutput.Select(len, str.Length);
                            this.rtbOutput.SelectionFont = new Font(System.Drawing.FontFamily.GenericMonospace, 10);
                            this.rtbOutput.SelectionColor = Color.Blue;

                    len = rtbOutput.TextLength;
                            str = "--WARNING: THIS SOFTWARE MUST BE USED BY AUTHORIZED & TRAINED PERSONS ONLY. MIS-OPERATION MAY AFFECT ON-AIR--" + Environment.NewLine + Environment.NewLine;
                            this.rtbOutput.AppendText(str);
                            this.rtbOutput.Select(len, str.Length);
                            this.rtbOutput.SelectionFont = new Font(System.Drawing.FontFamily.GenericSansSerif, 9,System.Drawing.FontStyle.Bold);
                            this.rtbOutput.SelectionColor = Color.Red;
                            this.rtbOutput.DeselectAll();
                    break;
                case 3:
                    this.rtbOutput.Clear();
                    break;
            }
            rtbOutput.SelectionStart = rtbOutput.TextLength;
            rtbOutput.ScrollToCaret();
        }

        /// <summary>
        /// If unchecked, cleanup BUG engine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkBxEnableBUG_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkBxEnableBUG.Checked && this.currentBUGEngine != "NONE")
            {
                this.btnBUGEngineCleanup_Click(null, null);
                //**following line is newly added
                this.radBtn11.Checked = false;
            }
            else if (chkBxEnableBUG.Checked && this.currentBUGEngine == "NONE")
            {
                MessageBox.Show("Please select BUG Engine first", "Engine not selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                chkBxEnableBUG.Checked = false;
                return;
            }
            //**newly added
            else if (chkBxEnableBUG.Checked)
            {
                this.radBtn11.Checked = true;
            }
        }
    }
}
