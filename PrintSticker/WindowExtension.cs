using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Common;
using DAL;
using LSExtensionWindowLib;
using One1.Controls;
using LSSERVICEPROVIDERLib;


namespace PrintSticker
{

    [ComVisible(true)]
    [ProgId("PrintSticker.PrintStickercls")]
    public partial class PrintSticker : UserControl, IExtensionWindow
    {
        #region Ctor
        public PrintSticker()
        {
            InitializeComponent();
            this.BackColor = Color.FromName("Control");

        }
        #endregion


        #region private members


        private IExtensionWindowSite2 _ntlsSite;
        private INautilusDBConnection _ntlsCon;
        private INautilusProcessXML _processXml;
        private IDataLayer dal;
        private int _port = 9100;
        private NautilusUser _user;
        private const string Type = "3";
        #endregion


        #region Implementation of IExtensionWindow

        public bool CloseQuery()
        {
            return true;
        }

        public void Internationalise() { }

        public void SetSite(object site)
        {
            _ntlsSite = (IExtensionWindowSite2)site;
            _ntlsSite.SetWindowInternalName("PrintSticker");
            _ntlsSite.SetWindowRegistryName("PrintSticker");
            _ntlsSite.SetWindowTitle("Print Sticker");
            try
            {
                _ntlsSite.SetWindowIcon("Printer.ico");

            }
            catch (Exception e)
            {

                //row;
            }


        }


        public void PreDisplay()
        {
            Utils.CreateConstring(_ntlsCon);
            dal = new DataLayer();
            dal.Connect();
            timerFocus.Start();
        }

        public WindowButtonsType GetButtons()
        {
            return LSExtensionWindowLib.WindowButtonsType.windowButtonsNone;
        }

        public bool SaveData()
        {
            return false;
        }

        public void SetServiceProvider(object serviceProvider)
        {
            var sp = serviceProvider as NautilusServiceProvider;
            _processXml = Utils.GetXmlProcessor(sp);
            _ntlsCon = Utils.GetNtlsCon(sp);
            _user = Utils.GetNautilusUser(sp);
        }

        public void SetParameters(string parameters)
        {

        }

        public void Setup()
        {

        }

        public WindowRefreshType DataChange()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public WindowRefreshType ViewRefresh()
        {
            return LSExtensionWindowLib.WindowRefreshType.windowRefreshNone;
        }

        public void refresh() { }

        public void SaveSettings(int hKey) { }

        public void RestoreSettings(int hKey) { }
        public void Close()
        {

        }


        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var val = textBoxBrCode.Text;
            Result result;
            string sampleName;
            var testcode = "";

            if (val.Contains("r") || val.Contains("R") || val.Contains("ר"))
            {
                string rId = val.Remove(0, 1);
                if (!Validate(rId)) return;
                result = dal.GetResultById(long.Parse(rId));
                if (result == null)
                {
                    CustomMessageBox.Show("לא נמצאה התוצאה המבוקשת.");
                    Init();
                    return;
                }
                sampleName = result.Test.Aliquot.Sample.Name;
                testcode = GetTestCode(result.Test.Aliquot, testcode);
            }
            else
            {
                if (!Validate(val)) return;
                var sample = dal.GetSampleByKey(Convert.ToInt32(val));
                if (sample == null)
                {
                    CustomMessageBox.Show("לא נמצאה הדגימה המבוקשת.");
                    Init();
                    return;
                }
                sampleName = sample.Name;
                testcode = GetTestCode(sample.Aliqouts.FirstOrDefault(), testcode);
            }
            var work = _user.GetWorkstationId();
            Workstation ws = dal.getWorkStaitionById((long)work);
            ReportStation reportStation = dal.getReportStationByWorksAndType(ws.NAME, Type);
            string goodIp = "";
            if (reportStation != null)
            {
                if (reportStation.Destination != null)
                {
                    //מביא את הIP של המדפסת להדפסה הזאת
                    goodIp = reportStation.Destination.ManualIP;
                }
            }
            try
            {
                for (int i = 0; i < numericUpDown1.Value; i++)
                {
                    Print(sampleName, sampleName, testcode, "", goodIp);
                }
            }
            catch (Exception ex)
            {
                One1.Controls.CustomMessageBox.Show("שגיאה בהדפסה , אנא פנה לתמיכה.", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Logger.WriteLogFile(ex);
            }
            Init();
        }

        private bool Validate(string rId)
        {
            long x;
            if (long.TryParse(rId, out x))
            {
                return true;
            }
            else
            {
                One1.Controls.CustomMessageBox.Show("ערך לא חוקי", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void Init()
        {
            textBoxBrCode.Text = "";
            numericUpDown1.Value = 1;
            textBoxBrCode.Focus();
        }
        private string GetTestCode(Aliquot aliq, string testcode)
        {
            if (aliq.Parent.Count == 0 && aliq.U_CHARGE == "T")
            {
                testcode = aliq.ShortName;
            }
            if (aliq.Parent.Count != 0)
            {
                GetTestCode(aliq.Parent.FirstOrDefault(), testcode);
            }
            return testcode;
        }

        private void Print(string name, string ID, string testcode, string mihol, string ip)
        {
            string ipAddress = ip;


            // ZPL Command(s)
            string ntxt = name;
            string tctxt = testcode;
            string mtxt = mihol;
            string itxt = ID;


            string ZPLString =
                 "^XA" +
                 "^LH0,0" +
                 "^FO20,10" +
                 "^A@N20,20" +
                string.Format("^FD{0}^FS", ntxt) + //שם
                 "^FO10,60" +
                 "^A@N20,20" +

                 string.Format("^FD{0}^FS", mtxt) +
                "^FO100,60" +
                 "^A@N20,20" +

                 string.Format("^FD{0}^FS", tctxt) +
                "^FO260,0" + "^BQN,4,3" +
                //string.Format("^FD   {0}^FS", itxt) + //ברקוד
                    string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
                "^XZ";

            // Open connection
            var client = new System.Net.Sockets.TcpClient();
            client.Connect(ipAddress, _port);
            // Write ZPL String to connection
            var writer = new StreamWriter(client.GetStream());
            writer.Write(ZPLString.Trim());
            writer.Flush();
            // Close Connection
            writer.Close();
            client.Close();

        }



        private void PrintSticker_Resize(object sender, EventArgs e)
        {
            panel1.Location = new Point(Width / 2 - panel1.Width / 2, panel1.Location.Y);

        }

        private void timerFocus_Tick(object sender, EventArgs e)
        {
            textBoxBrCode.Focus();
            timerFocus.Stop();


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Resize(object sender, EventArgs e)
        {
            lblHeader.Location = new Point(Width / 2 - panel1.Width / 2, panel1.Location.Y);

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBoxBrCode_KeyDown(object sender, KeyEventArgs e)
        {
           if (e.KeyValue == 13)
                numericUpDown1.Focus();
        }

        private void numericUpDown1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
               button1.Focus();

        }


    }
}

