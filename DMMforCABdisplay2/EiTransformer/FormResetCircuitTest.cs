using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ivi.Visa.Interop;
using NationalInstruments.DAQmx;


namespace EiTransformer
{
    public partial class FormResetCircuitTest : Form
    { 
        #region Initial Form9
        
        public FormResetCircuitTest()
        {
            InitializeComponent();
           txtTime1.Text = "3000";
            txtTime2.Text = "5000";
        }

        #endregion

        #region Wait function

        public void Wait(int ms)
        {
            DateTime start = DateTime.Now;
            while ((DateTime.Now - start).TotalMilliseconds < ms) ;
            Application.DoEvents();
            
        }

        #endregion

        #region staticValiable

        public static string usbDMM = "USB0::0x2A8D::0x1601::MY59020517";//::0::INSTR"; for AMW DMM34460A  "USB0::0x2A8D::0x1601::MY59011307" Ei-transformer;
       
        public static string measStep1 = "1.Check Pull Down";
        public static double uperSpec1;//= //1.8;//Vac
        public static double lowerSpec1;// = //-0.5;//Vac

        public static string measStep2 = "2.Check Pull Up";
        public static double uperSpec2;//= //3.5;//Vac
        public static double lowerSpec2;//= //3.0;//Vac
        public static string ResultV1;
        public static string ResultV2;
        public int time1;
        public int time2;
        
        #endregion

        #region Press Start

        private void buttonStart_Click(object sender, EventArgs e)// Include the SW1,SW2 () to start testing when SW1,SW2 press as the same time(If SW1 SW2 = 0 then start).

        {
            if(usbDMM == "USB0::0x2A8D::0x1601::MY59020517")// Include the SW1,SW2 () to start testing when SW1,SW2 press as the same time.
            {
                
                try
                    {
                    
                    time1 = Convert.ToInt32(txtTime1.Text);
                    time2 = Convert.ToInt32(txtTime2.Text);
                    txtmeaStep1.Text = measStep1;
                        txtmeaStep2.Text = measStep2;
                    MessageBox.Show("Please select RFI Reset ==> OFF ");
                    Wait(1000);

                    txtShowFinal.Text = "TEST";
                        txtShowFinal.BackColor = Color.Yellow;

                        /**************************************************************     E.Set the Connection of DMM and PC    ***********************************************************/

                        FormResetCircuitTest DmmClass = new FormResetCircuitTest();        //Create an instance of this class so we can call functions from Main
                        Ivi.Visa.Interop.ResourceManager rm = new Ivi.Visa.Interop.ResourceManager();       //Open up a new resource manager
                        Ivi.Visa.Interop.FormattedIO488 myDmm = new Ivi.Visa.Interop.FormattedIO488();      //Open a new Formatted IO 488 session 
                        try
                        {
                            //Exception//  
                            //string DutAddr = "GPIB0::22";                  //String for GPIB
                            //string DutAddr = ""TCPIP0::169.254.4.61";                 //Example string for LAN
                            //string DutAddr = "USB0::0x0957::0x1A07::MY53000101";           //Example string for USB

                            string DutAddr = usbDMM;          //String for USB

                            myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, "");         //Open up a handle to the DMM with a 2 second timeout
                            myDmm.IO.Timeout = 3000;                                        //You can also set your timeout by doing this command, sets to 3 seconds

                            /**************************************************************     End Set the Connection of DMM and PC    ***********************************************************/

                            /**************************************************************     F.Check communication of DMM and PC           ************************************************************************/

                            //First start off with a reset state
                            myDmm.IO.Clear();                       //Send a device clear first to stop any measurements in process
                            myDmm.WriteString("*RST", true);            //Reset the device
                            Wait(100);
                            myDmm.WriteString("*IDN?", true);           //Get the IDN string  
                            string IDN = myDmm.ReadString();
                            //Console.WriteLine(IDN);               //report the DMM's identity

                            /**************************************************************    End Check communication of DMM and PC         *************************************************************************/

                            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            ///
                            ///                                                             Test Sequence of this Program
                            ///                                                             Test Sequence of this Program
                            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            
                            /**************************************************************     Step1.Check input AC voltage >>Test1    ***********************************************************/
                            //Configure for ACV 750 range, to read a 100 Hz signal
                            //myDmm.WriteString("CONF:VOLT:AC 100, 0.001", true);
                            myDmm.WriteString("CONF:VOLT:DC AUTO", true);
                        //myDmm.WriteString("CONF:VOLT:AC", true);
                        //myDmm.WriteString("CONF:VOLT:AUTO ONCE", true);
                        //myDmm.WriteString("SAMP:COUN 2", true);

                        // myDmm.WriteString("VOLT:AC:BAND 20", true);     //Choose the band that is lower than your input frequency, Bands are 3 Hz|20 Hz|200 Hz

                        
                        Wait(time1);
                            myDmm.WriteString("READ?", true);
                            Wait(1000);
                            double ACVResultStep1 = Convert.ToDouble(myDmm.ReadString());
                            ACVResultStep1 = Math.Round(ACVResultStep1, 4);
                            resultValue1.Text = ACVResultStep1.ToString();
                            Wait(1000);
                        DmmClass.CheckDMMError(myDmm); //Check if the DMM has any errors    
                            double resultStep1 = ACVResultStep1;


                         double H1 = Convert.ToDouble(specH1.Text);
                         double L1 = Convert.ToDouble(specL1.Text);


                        uperSpec1 = H1;
                        lowerSpec1 = L1;


                        if (resultStep1 >= lowerSpec1 && resultStep1 <= uperSpec1)
                            {
                            Wait(100);
                            resultValue1.BackColor = Color.Chartreuse;
                            ResultV1 = "OK";
                            Wait(100);
                        }
                            else
                            {
                               
                            resultValue1.BackColor = Color.Red;
                            Wait(100);
                            ResultV1 = "NG";
                            Wait(100);

                        }

                        /**

                        /**************************************************************     End Check input AC voltage        ***********************************************************/
                    
                        Wait(time2);
                            /**************************************************************     Step2 Check output AC voltage at secondary side Pin8-12 >>Test2       ***********************************************************/
                            //Configure for ACV 100V range, to read a 100 Hz signal
                            myDmm.WriteString("CONF:VOLT:DC AUTO", true);
                        Wait(1000);
                        // myDmm.WriteString("VOLT:AC:BAND 20", true);//Choose the band that is lower than your input frequency, Bands are 3 Hz|20 Hz|200 Hz
                        myDmm.WriteString("READ?", true);
                           // Wait(1000);

                           
                          
                            double ACVResultStep2 = Convert.ToDouble(myDmm.ReadString());
                            ACVResultStep2 = Math.Round(ACVResultStep2, 4);
                            resultValue2.Text = ACVResultStep2.ToString();
                        Wait(1000);
                        DmmClass.CheckDMMError(myDmm); //Check if the DMM has any errors    
                            double resultStep2 = ACVResultStep2;


                        double H2 = Convert.ToDouble(specH2.Text);
                        double L2 = Convert.ToDouble(specL2.Text);


                        uperSpec2 = H2;
                        lowerSpec2 = L2;



                        if (resultStep2 >= lowerSpec2 && resultStep2 <= uperSpec2)
                        {
                            Wait(100);
                            resultValue2.BackColor = Color.Chartreuse;
                            ResultV2 = "OK";
                            Wait(100);
                        }
                            else
                            {
                            Wait(100);
                            resultValue2.BackColor = Color.Red;
                            ResultV2 = "NG";
                            Wait(100);
                        }

                        /**************************************************************       End Check output AC voltage at secondary side >>Test2          ***********************************************************/


                        /***************************************************************************    Other function  ***********************************************************
                         *      //Configure for DCI 10A range, 100uA resolution
                      myDmm.WriteString("CONF:CURR:DC 10, 0.0001", true);
                      myDmm.WriteString("READ?", true);
                      string DCIResult = myDmm.ReadString();
                      Console.WriteLine("DCI Reading = " + DCIResult); //report the DCI reading
                      DmmClass.CheckDMMError(myDmm); //Check if the DMM has any errors          
                         * 
                         * 
                         * //Configure for DCV 100V range, 100uV resolution
                    myDmm.WriteString("CONF:VOLT:DC 100, 0.0001", true);
                    myDmm.WriteString("READ?", true);
                    string DCVResult = myDmm.ReadString();
                    Console.WriteLine("DCV Reading = " + DCVResult); //report the DCV reading
                    DmmClass.CheckDMMError(myDmm); //Check if the DMM has any errors
                         * 
                         * 
                         * //Configure for OHM 2 wire 100 Ohm range, 100uOhm resolution
                    myDmm.WriteString("CONF:RES 100, 0.0001", true);
                    myDmm.WriteString("READ?", true);
                    string Res2WResult = myDmm.ReadString();
                    Console.WriteLine("2 Wire Resistance Reading = " + Res2WResult); //report the 2W resistance reading
                    DmmClass.CheckDMMError(myDmm); //Check if the DMM has any errors
                        *
                         ***********************************************************************    Other function  ****************************************************************/
                        // Wait(100);
                        /**************************************************************     6.Show test result    *******************************************************************/
                      
                        if (ResultV1=="OK"&&ResultV2 =="OK")
                        {
                            txtShowFinal.Text = "PASS";
                            txtShowFinal.Show();
                               txtShowFinal.BackColor = Color.LimeGreen;
                          
                               
                            }
                            else
                            {
                           
                            txtShowFinal.Text = "FAIL";
                        
                            txtShowFinal.Show();
                            txtShowFinal.BackColor = Color.Red;
                                
                                Wait(100);
                               
                            }
                        /**************************************************************      End Show test result     *******************************************************************/

                           MessageBox.Show("Finished");
                        specH1.Text = Convert.ToString(H1);
                        specL1.Text = Convert.ToString(L1);

                        specH2.Text = Convert.ToString(H2);
                        specL2.Text = Convert.ToString(L2);

                        /**************************************************************     8.Check Error all function        *******************************************************************/
                    }
                        catch (Exception e1)
                        {
                            MessageBox.Show("Error occured" + e1.Message);
                            //ClearParameter();
                            Dispose();
                        }
                        finally
                        {
                            //Close out your resources
                            try
                            {
                                Wait(200);
                                myDmm.IO.Close();
                            }
                            catch
                            {
                            }
                            try
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(myDmm);
                            }
                            catch
                            {
                            }
                            try
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rm);
                            }
                            catch { }

                            //MessageBox.Show("ERROR! Please OFF Power");
                            Wait(100);
                            ClearParameter();
                          //  writerPort2_Dev2.WriteSingleSampleMultiLine(true, OFFsolenoid);//ON solenoid
                            // Set the Cover's Solenoid = OFF to open the Cover. 
                            // Set the TrnasformerLock 's Solenoid = OFF to remove the Product.
                        }
                    }
                    catch
                    {
                        MessageBox.Show("System Error, Please SHUTDOWN MACHINE! NOW! ");
                        Dispose();
                        ClearParameter();
                    }

                    //string xx1 = Convert.ToString(xx); // for loop
                    //textBox1.Text = xx1 ; //for loop

               // }// For for loop test.
            }
            else
            {
                MessageBox.Show("DMM address Changed");
               
            }

            /**************************************************************     End Check Error all function       *******************************************************************/
        }

        #endregion

        #region CheckError

        public void CheckDMMError(FormattedIO488 myDmm)
        {
            myDmm.WriteString("SYST:ERR?", true);
            string errStr = myDmm.ReadString();

            if (errStr.Contains("No error"))//If no error, then return
                return;
            //If there is an error, read out all of the errors and return them in an exception
            else
            {
                string errStr2 = "";
                do
                {
                    myDmm.WriteString("SYST:ERR?", true);
                    errStr2 = myDmm.ReadString();
                    if (!errStr2.Contains("No error")) errStr = errStr + "\n" + errStr2;

                } while (!errStr2.Contains("No error"));

                throw new Exception("Exception: Encountered system error(s)\n" + errStr);
            }

        }

        #endregion

        #region Clear all parameter

        public void ClearParameter()
        {
            resultValue1.Text = "";
            resultValue1.BackColor = default(Color);
            resultValue2.Text = "";
            resultValue2.BackColor = default(Color);
            txtShowFinal.Text = "";
            txtShowFinal .BackColor = default(Color);
            ResultV1 = "";
            ResultV2 = "";
        }
        #endregion

        #region Press stop
        private void buttonStop_Click(object sender, EventArgs e)
        {
            //Dispose();
            //this.Close();// close only Form not close program.
            //Dispose();
            // System.Windows.Forms.Application.ExitThread();// close all program.
            //System.Windows.Forms.Application.Exit();// close all program.
            // Dispose();
            Environment.Exit(Environment.ExitCode);
        }

        #endregion

        #region loadForm9
        private void Form9_Load(object sender, EventArgs e)
        {

        }

        #endregion
    }
}
