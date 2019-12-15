using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using PropertyAttributes = System.Data.PropertyAttributes;
using DLA_to_Excel.Classes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using Color = System.Drawing.Color;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;


namespace DLA_to_Excel
{
    public partial class Form1 : Form
    {
        public string FileLocation { get; set; }
        public string FolderDestination { get; set; }
        public string SheetName { get; set; }
        public int SheetCounter { get; set; }
        public string CompleteFilePath { get; set; }
        public int TotalRows { get; set; }
        public ProgressBar pb { get; set; }
        public OpenFileDialog opfd = new OpenFileDialog();
        public FolderBrowserDialog fbd = new FolderBrowserDialog();
        public Workbook Workbook { get; set; }
        public Worksheet Worksheet { get; set; }
        public string SlSheetName { get; set; }
        public Application XlApp { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public DateTime TimeStart { get; set; }
        public DateTime TimeStop { get; set; }
        public TimeSpan TimeSpan { get; set; }
        public string EndTime { get; set; }
        public bool FastForward { get; set; }

        #region Initialize Form

        public Form1()
        {
            InitializeComponent();
            var ver = System.Windows.Forms.Application.ProductVersion;
            this.Text = "DLA to Excel    Version: " + ver;
            pb = progressBar1;
            ProgressBarActive();
            txtFolderDestination.Text = @"D:\Tony\" + DateTime.Now.ToString("yyyy") + @"\" +
                                        DateTime.Now.ToString("MMM") + @"\Excel Conversions";
            Row = 1;
            Column = 1;
            opfd.Filter = "Text|*.txt|All|*.*"; //.txt files only
        }

        #endregion

        #region Get File Location

        private void BtnFileLocation_Click(object sender, EventArgs e)
        {
            var result = opfd.ShowDialog();
            if (result == DialogResult.OK)
            {
                FileLocation = opfd.FileName;
                txtFilePath.Text = opfd.FileName;
            }
        }

        #endregion

        #region Get Folder Destination

        private void BtnFolderDestination_Click(object sender, EventArgs e)
        {
            var result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                FolderDestination = fbd.SelectedPath;
                txtFolderDestination.Text = fbd.SelectedPath;
            }
        }

        #endregion

        #region Start Conversion

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (!CheckForEmptyFields(sender, e)) return;
            this.BackColor = Color.DarkGreen;
            Cursor = Cursors.WaitCursor;
            var checkFields = CheckFields(sender, e);
            Row = 1;
            Column = 1;
            TimeStart = DateTime.Now;
            lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
            lblProgress.Text = "Progress: ";
            ProgressBarActive();

            //Get sheet name
            SheetName = cmbTableName.Text;
            lblCurrentDoc.Text = SheetName;

            //Get Folder Destination
            FolderDestination = "";
            FolderDestination = txtFolderDestination.Text;

            //Combine to make complete path with date
            CompleteFilePath = "";
            CompleteFilePath = FolderDestination + @"\" + SheetName + "-" + DateTime.Now.ToString("yyyy") + "-" +
                               DateTime.Now.ToString("MMM") + ".xlsx";

            //Check if file exists prior to transferring data
            if (File.Exists(CompleteFilePath))
            {
                MessageBox.Show("File Already Exists.", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            switch (cmbTableName.Text)
            {
                case "ENAC":
                    ENAC();
                    break;
                case "AMMO":
                    AMMO();
                    break;
                case "CAGECDS":
                    CAGECDS();
                    break;
                case "CAGE-DateEstAndChgd":
                    CAGEDateEstAndChgd();
                    break;
                case "CHARDAT":
                    CHARDAT();
                    break;
                case "COLXREF":
                    COLXREF();
                    break;
                case "FCAGE":
                    FCAGE();
                    break;
                case "FCAN-SEGK":
                    FCANSEGK();
                    break;
                case "FLISFOIA":
                    FLISFOIA();
                    break;
                default:
                    break;
            }
        }
        #endregion

        #region FLISFOIA
        private void FLISFOIA()
        {
            //Get # Books needed that is less than max excel row (250,000)
            var maxRows = 250000;
            var numBooks = TotalRows / maxRows;
            //Number of excel books
            var bookCounter = 7;
            var totalCounter = 0;
            var aRow = 1;
            var bRow = 1;
            var cRow = 1;
            var eRow = 1;
            var gRow = 1;
            var hRow = 1;
            var wRow = 1;
            Column = 1;
            var segA = "SEGMENT_A"; var segB = "SEGMENT_B"; var segC = "SEGMENT_C"; var segE = "SEGMENT_E"; var segG = "SEGMENT_G"; var segH = "SEGMENT_H"; var segW = "SEGMENT_W";

            pb.Maximum = numBooks;
            pb.Step = 1;
            pb.Value = 1;
            try
            {
                while (bookCounter <= numBooks)
                {
                    pb.PerformStep();
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;
                    var ffModel = new FLISFOI();
                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        //NEED THIS OR FILE WILL OVERRIDE PREVIOUS FILE
                        UpdateSheetNames(bookCounter);
                        using (SLDocument sl = new SLDocument())
                        {
                            //SEG A
                            while (bookCounter <= numBooks && aRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, segA); //RECORD TYPE 01
                                if (aRow == 1)
                                {
                                    sl.SetCellValue(aRow, Column, "FSG_3994");
                                    sl.SetCellValue(aRow, ++Column, "FSC_WI_FSG_3996");
                                    sl.SetCellValue(aRow, ++Column, "NCB_CD_4130");
                                    sl.SetCellValue(aRow, ++Column, "I_I_NBR_4131");
                                    sl.SetCellValue(aRow, ++Column, "FIIG_4065");
                                    sl.SetCellValue(aRow, ++Column, "SHRT_NM_2301");
                                    sl.SetCellValue(aRow, ++Column, "NAIN_5020");
                                    sl.SetCellValue(aRow, ++Column, "CRITL_CD_FIIG_3843");
                                    sl.SetCellValue(aRow, ++Column, "TYP_II_4820");
                                    sl.SetCellValue(aRow, ++Column, "RPDMRC_4765");
                                    sl.SetCellValue(aRow, ++Column, "DEMIL_CD_0167");
                                    sl.SetCellValue(aRow, ++Column, "DT_NIIN_ASGMT_2180");
                                    sl.SetCellValue(aRow, ++Column, "HMIC_0865");
                                    sl.SetCellValue(aRow, ++Column, "ESD_EMI_CD_2043");
                                    sl.SetCellValue(aRow, ++Column, "PMIC_0802");
                                    sl.SetCellValue(aRow, ++Column, "ADPEC_0801");
                                    aRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "01")
                                    {
                                        try { ffModel.FSG_3994 = nextLine.Substring(2, 2); } catch { ffModel.FSG_3994 = ""; }
                                        try { ffModel.FSC_WI_FSG_3996 = nextLine.Substring(3, 2); } catch { ffModel.FSC_WI_FSG_3996 = ""; }
                                        try { ffModel.NCB_CD_4130 = nextLine.Substring(5, 2); } catch { ffModel.NCB_CD_4130 = ""; }
                                        try { ffModel.I_I_NBR_4131 = nextLine.Substring(7, 7); } catch { ffModel.I_I_NBR_4131 = ""; }
                                        try { ffModel.FIIG_4065 = nextLine.Substring(14, 6); } catch { ffModel.FIIG_4065 = ""; }
                                        try { ffModel.INC_4080 = nextLine.Substring(20, 5); } catch { ffModel.INC_4080 = ""; }
                                        try { ffModel.SHRT_NM_2301 = nextLine.Substring(25, 19); } catch { ffModel.SHRT_NM_2301 = ""; }
                                        try { ffModel.NAIN_5020 = nextLine.Substring(44, 19); } catch { ffModel.NAIN_5020 = ""; }
                                        try { ffModel.CRITL_CD_FIIG_3843 = nextLine.Substring(63, 1); } catch { ffModel.CRITL_CD_FIIG_3843 = ""; }
                                        try { ffModel.TYP_II_4820 = nextLine.Substring(64, 1); } catch { ffModel.TYP_II_4820 = ""; }
                                        try { ffModel.RPDMRC_4765 = nextLine.Substring(65, 1); } catch { ffModel.RPDMRC_4765 = ""; }
                                        try { ffModel.DEMIL_CD_0167 = nextLine.Substring(66, 1); } catch { ffModel.DEMIL_CD_0167 = ""; }
                                        try { ffModel.DT_NIIN_ASGMT_2180 = nextLine.Substring(67, 7); } catch { ffModel.DT_NIIN_ASGMT_2180 = ""; }
                                        try { ffModel.HMIC_0865 = nextLine.Substring(74, 1); } catch { ffModel.HMIC_0865 = ""; }
                                        try { ffModel.ESD_EMI_CD_2043 = nextLine.Substring(75, 1); } catch { ffModel.ESD_EMI_CD_2043 = ""; }
                                        try { ffModel.PMIC_0802 = nextLine.Substring(76, 1); } catch { ffModel.PMIC_0802 = ""; }
                                        try { ffModel.ADPEC_0801 = nextLine.Substring(77, 1); } catch { ffModel.ADPEC_0801 = ""; }
                                        sl.SetCellValue(aRow, Column, ffModel.FSG_3994);
                                        sl.SetCellValue(aRow, ++Column, ffModel.FSC_WI_FSG_3996);
                                        sl.SetCellValue(aRow, ++Column, ffModel.NCB_CD_4130);
                                        sl.SetCellValue(aRow, ++Column, ffModel.I_I_NBR_4131);
                                        sl.SetCellValue(aRow, ++Column, ffModel.FIIG_4065);
                                        sl.SetCellValue(aRow, ++Column, ffModel.INC_4080);
                                        sl.SetCellValue(aRow, ++Column, ffModel.SHRT_NM_2301);
                                        sl.SetCellValue(aRow, ++Column, ffModel.NAIN_5020);
                                        sl.SetCellValue(aRow, ++Column, ffModel.CRITL_CD_FIIG_3843);
                                        sl.SetCellValue(aRow, ++Column, ffModel.TYP_II_4820);
                                        sl.SetCellValue(aRow, ++Column, ffModel.RPDMRC_4765);
                                        sl.SetCellValue(aRow, ++Column, ffModel.DEMIL_CD_0167);
                                        sl.SetCellValue(aRow, ++Column, ffModel.DT_NIIN_ASGMT_2180);
                                        sl.SetCellValue(aRow, ++Column, ffModel.HMIC_0865);
                                        sl.SetCellValue(aRow, ++Column, ffModel.ESD_EMI_CD_2043);
                                        sl.SetCellValue(aRow, ++Column, ffModel.PMIC_0802);
                                        sl.SetCellValue(aRow, ++Column, ffModel.ADPEC_0801);
                                        aRow++;
                                        Column = 1;
                                    }

                                }
                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;
                            
                            //SEG B
                            while (bookCounter <= numBooks && bRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.AddWorksheet(segB);  //RECORD TYPE 02                              
                                if (bRow == 1)
                                {
                                    //SET SEG_B COLUMN NAMES
                                    sl.SetCellValue(bRow, Column, "MOE_RULE_NBR_8290");
                                    sl.SetCellValue(bRow, ++Column, "AMC_2871");
                                    sl.SetCellValue(bRow, ++Column, "AMSC_2876");
                                    sl.SetCellValue(bRow, ++Column, "NIMSC_0076");
                                    sl.SetCellValue(bRow, ++Column, "EFF_DT_2128");
                                    sl.SetCellValue(bRow, ++Column, "IMC_2744");
                                    sl.SetCellValue(bRow, ++Column, "IMC_ACTY_2748");
                                    sl.SetCellValue(bRow, ++Column, "DSOR_0903");
                                    sl.SetCellValue(bRow, ++Column, "SUPPLM_COLLBR_2533");
                                    sl.SetCellValue(bRow, ++Column, "SUPPLM_RCVR_2534");
                                    sl.SetCellValue(bRow, ++Column, "AAC_2507");
                                    sl.SetCellValue(bRow, ++Column, "FMR_MOE_RULE_8280");
                                    bRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "02")
                                    {
                                        var numMoeRules = int.Parse(nextLine.Substring(2, 2));
                                        var index = 4;
                                        StringBuilder text = new StringBuilder("");
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.MOE_RULE_NBR_8290 = nextLine.Substring(index, 4));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow,Column,text.ToString());
                                        text.Clear();
                                        Column++;
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.AMC_2871 = nextLine.Substring(index + 4, 1));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        text.Clear();
                                        Column++;
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.AMSC_2876 = nextLine.Substring(index + 5, 1));
                                                if (i + 1 != numMoeRules) text.Append("|"); 
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.NIMSC_0076 = nextLine.Substring(index + 6, 1));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.EFF_DT_2128 = nextLine.Substring(index + 7, 5));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.IMC_2744 = nextLine.Substring(index + 12, 1));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.IMC_ACTY_2748 = nextLine.Substring(index + 13, 2));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.DSOR_0903 = nextLine.Substring(index + 15, 8));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.SUPPLM_COLLBR_2533 = nextLine.Substring(index + 23, 18));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.SUPPLM_RCVR_2534 = nextLine.Substring(index + 41, 18));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.AAC_2507 = nextLine.Substring(index + 59, 1));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        Column++;
                                        text.Clear();
                                        for (int i = 0; i < numMoeRules; i++)
                                        {
                                            index = 0;
                                            try
                                            {
                                                text.Append(ffModel.FMR_MOE_RULE_8280 = nextLine.Substring(index + 60, 4));
                                                if (i + 1 != numMoeRules) text.Append("|");
                                                index += 64;
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(bRow, Column, text.ToString());
                                        bRow++;
                                        Column = 1;
                                    }

                                }

                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //SEG C
                            while (bookCounter <= numBooks && cRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.AddWorksheet(segC);  //RECORD TYPE 03
                                if (cRow == 1)
                                {
                                    //SET SEG_C COLUMN NAMES
                                    sl.SelectWorksheet(segC);
                                    sl.SetCellValue(cRow, Column, "RNFC_2920");
                                    sl.SetCellValue(cRow, ++Column, "RNCC_2910");
                                    sl.SetCellValue(cRow, ++Column, "RNVC_4780");
                                    sl.SetCellValue(cRow, ++Column, "DAC_2640");
                                    sl.SetCellValue(cRow, ++Column, "RNAAC_2900");
                                    sl.SetCellValue(cRow, ++Column, "RNSC_2923");
                                    sl.SetCellValue(cRow, ++Column, "RNJC_2750");
                                    sl.SetCellValue(cRow, ++Column, "CAGE_CD_9250");
                                    sl.SetCellValue(cRow, ++Column, "REF_NBR_3570");
                                    sl.SetCellValue(cRow, ++Column, "SADC_4672");
                                    sl.SetCellValue(cRow, ++Column, "HCC_2579");
                                    sl.SetCellValue(cRow, ++Column, "MSDS_ID_9076");
                                    cRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "03")
                                    {
                                        var index = 2;
                                        var text = new StringBuilder("");
                                        var processCageData = true;
                                        while (processCageData)
                                        {
                                            try { text.Append(ffModel.RNFC_2920 = nextLine.Substring(index, 1)); }
                                            catch
                                            {
                                                text.Append("");
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, Column, text.ToString());
                                        text.Clear();

                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RNCC_2910 = nextLine.Substring(index + 1, 1));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();

                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RNVC_4780 = nextLine.Substring(index + 2, 1));
                                            }
                                            catch
                                            {
                                                text.Append("");
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.DAC_2640 = nextLine.Substring(index + 3, 1));
                                            }
                                            catch
                                            {                                             
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RNAAC_2900 = nextLine.Substring(index + 4, 2));
                                            }
                                            catch
                                            {
                                                text.Append("");
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };

                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RNSC_2923 = nextLine.Substring(index + 6, 1));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RNJC_2750 = nextLine.Substring(index + 7, 1));
                                            }
                                            catch
                                            {
                                                text.Append("");
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.CAGE_CD_9250 = nextLine.Substring(index + 8, 5));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.REF_NBR_3570 = nextLine.Substring(index + 13, 32));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.SADC_4672 = nextLine.Substring(index + 45, 2));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.HCC_2579 = nextLine.Substring(index + 47, 2));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        processCageData = true;
                                        while (processCageData)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.MSDS_ID_9076 = nextLine.Substring(index + 49, 5));
                                            }
                                            catch
                                            {
                                            }
                                            if (nextLine.Length > index + 54)
                                            {
                                                text.Append("|");
                                                index += 54;
                                            }
                                            else
                                                processCageData = false;
                                        };
                                        sl.SetCellValue(cRow, ++Column, text.ToString());
                                        text.Clear();
                                        cRow++;
                                        Column = 1;
                                    }

                                }

                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //SEG E
                            while (bookCounter <= numBooks && eRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, segA); //RECORD TYPE 01
                                if (eRow == 1)
                                {
                                    sl.AddWorksheet(segE);  //RECORD TYPE 04
                                    sl.SetCellValue(eRow, Column, "ISC_2650");
                                    sl.SetCellValue(eRow, ++Column, "ORG_STDZN_DEC_9325");
                                    sl.SetCellValue(eRow, ++Column, "DT_STDZN_DEC_2300");
                                    sl.SetCellValue(eRow, ++Column, "NIIN_STAT_CD_2670");
                                    sl.SetCellValue(eRow, ++Column, "RP_NSN_STD_RL_8977 / RPLM_NSN_STDZ_9525");
                                    sl.SetCellValue(eRow, ++Column, "ISC_2650");
                                    sl.SetCellValue(eRow, ++Column, "ORG_STDZN_DEC_9325");
                                    sl.SetCellValue(eRow, ++Column, "DT_STDZN_DEC_2300");
                                    sl.SetCellValue(eRow, ++Column, "NIIN_STAT_CD_2670");
                                    eRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "04")
                                    {
                                        //SEG E                                     
                                        sl.SelectWorksheet(segE);
                                        try
                                        {
                                            ffModel.ISC_2650 = nextLine.Substring(2, 1);
                                        }
                                        catch
                                        {
                                            ffModel.ISC_2650 = "";
                                        }

                                        try
                                        {
                                            ffModel.ORG_STDZN_DEC_9325 = nextLine.Substring(3, 2);
                                        }
                                        catch
                                        {
                                            ffModel.ORG_STDZN_DEC_9325 = "";
                                        }

                                        try
                                        {
                                            ffModel.DT_STDZN_DEC_2300 = nextLine.Substring(5, 7);
                                        }
                                        catch
                                        {
                                            ffModel.DT_STDZN_DEC_2300 = "";
                                        }

                                        try
                                        {
                                            ffModel.NIIN_STAT_CD_2670 = nextLine.Substring(12, 1);
                                        }
                                        catch
                                        {
                                            ffModel.NIIN_STAT_CD_2670 = "";
                                        }
                                        sl.SetCellValue(eRow, Column, ffModel.ISC_2650);
                                        sl.SetCellValue(eRow, ++Column, ffModel.ORG_STDZN_DEC_9325);
                                        sl.SetCellValue(eRow, ++Column, ffModel.DT_STDZN_DEC_2300);
                                        sl.SetCellValue(eRow, ++Column, ffModel.NIIN_STAT_CD_2670);
                                        var numRepCodes = int.Parse(nextLine.Substring(13, 2));
                                        var index = 15;
                                        var text = new StringBuilder("");
                                        for (int i = 0; i < numRepCodes; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.RP_NSN_STD_RL_8977_RPLM_NSN_STDZ_9525 =
                                                    nextLine.Substring(index, 13));
                                                if (i + 1 != numRepCodes)
                                                {
                                                    text.Append("|");
                                                    index += 24;
                                                }
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(eRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 15;
                                        for (int i = 0; i < numRepCodes; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.ISC_2650 = nextLine.Substring(index + 13, 1));
                                                if (i + 1 != numRepCodes)
                                                {
                                                    text.Append("|");
                                                    index += 24;
                                                }
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(eRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 15;
                                        for (int i = 0; i < numRepCodes; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.ORG_STDZN_DEC_9325 = nextLine.Substring(index + 14, 2));
                                                if (i + 1 != numRepCodes)
                                                {
                                                    text.Append("|");
                                                    index += 24;
                                                }
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(eRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 15;
                                        for (int i = 0; i < numRepCodes; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.DT_STDZN_DEC_2300 = nextLine.Substring(index + 16, 7));
                                                if (i + 1 != numRepCodes)
                                                {
                                                    text.Append("|");
                                                    index += 24;
                                                }
                                            }
                                            catch {}
                                        }

                                        sl.SetCellValue(eRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 15;
                                        for (int i = 0; i < numRepCodes; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.NIIN_STAT_CD_2670 = nextLine.Substring(index + 23, 1));
                                                if (i + 1 != numRepCodes)
                                                {
                                                    text.Append("|");
                                                    index += 24;
                                                }
                                            }
                                            catch{}
                                        }
                                        sl.SetCellValue(eRow, ++Column, text.ToString());
                                        text.Clear();
                                        eRow++;
                                        Column = 1;
                                    }

                                }

                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //SEG G
                            while (bookCounter <= numBooks && gRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.AddWorksheet(segG);  //RECORD TYPE 09
                                if (gRow == 1)
                                {
                                    sl.SetCellValue(gRow, Column, "INTGTY_CD_0864");
                                    sl.SetCellValue(gRow, ++Column, "ORIG_ACTY_CD_4210");
                                    sl.SetCellValue(gRow, ++Column, "RAIL_VARI_CD_4760");
                                    sl.SetCellValue(gRow, ++Column, "NMFC_2850");
                                    sl.SetCellValue(gRow, ++Column, "SUB_ITM_NBR_0861");
                                    sl.SetCellValue(gRow, ++Column, "UFC_CD_MODF_3040");
                                    sl.SetCellValue(gRow, ++Column, "HMC_2720");
                                    sl.SetCellValue(gRow, ++Column, "LCL_CD_2760");
                                    sl.SetCellValue(gRow, ++Column, "WRT_CMDTY_CD_9275");
                                    sl.SetCellValue(gRow, ++Column, "TYPE_CGO_CD_9260");
                                    sl.SetCellValue(gRow, ++Column, "SP_HDLG_CD_9240");
                                    sl.SetCellValue(gRow, ++Column, "AIR_DIM_CD_9220");
                                    sl.SetCellValue(gRow, ++Column, "AIR_CMTY_HDLG_9215");
                                    sl.SetCellValue(gRow, ++Column, "CLAS_RTNG_CD_2770");
                                    sl.SetCellValue(gRow, ++Column, "FRT_DESC_4020");
                                    gRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "09")
                                    {
                                        try
                                        {
                                            ffModel.INTGTY_CD_0864 = nextLine.Substring(2, 1);
                                        }
                                        catch
                                        {
                                            ffModel.INTGTY_CD_0864 = "";
                                        }

                                        try
                                        {
                                            ffModel.ORIG_ACTY_CD_4210 = nextLine.Substring(3, 2);
                                        }
                                        catch
                                        {
                                            ffModel.ORIG_ACTY_CD_4210 = "";
                                        }

                                        try
                                        {
                                            ffModel.RAIL_VARI_CD_4760 = nextLine.Substring(5, 1);
                                        }
                                        catch
                                        {
                                            ffModel.RAIL_VARI_CD_4760 = "";
                                        }
                                        try { ffModel.NMFC_2850 = nextLine.Substring(6, 6); } catch { ffModel.NMFC_2850 = ""; }
                                        try { ffModel.SUB_ITM_NBR_0861 = nextLine.Substring(12, 1); } catch { ffModel.SUB_ITM_NBR_0861 = ""; }
                                        try { ffModel.UFC_CD_MODF_3040 = nextLine.Substring(13, 5); } catch { ffModel.UFC_CD_MODF_3040 = ""; }
                                        try { ffModel.HMC_2720 = nextLine.Substring(18, 2); } catch { ffModel.HMC_2720 = ""; }
                                        try { ffModel.LCL_CD_2760 = nextLine.Substring(20, 1); } catch { ffModel.LCL_CD_2760 = ""; }
                                        try { ffModel.WRT_CMDTY_CD_9275 = nextLine.Substring(21, 3); } catch { ffModel.WRT_CMDTY_CD_9275 = ""; }
                                        try { ffModel.TYPE_CGO_CD_9260 = nextLine.Substring(24, 1); } catch { ffModel.TYPE_CGO_CD_9260 = ""; }
                                        try { ffModel.SP_HDLG_CD_9240 = nextLine.Substring(25, 1); } catch { ffModel.SP_HDLG_CD_9240 = ""; }
                                        try { ffModel.AIR_DIM_CD_9220 = nextLine.Substring(26, 1); } catch { ffModel.AIR_DIM_CD_9220 = ""; }
                                        try { ffModel.AIR_CMTY_HDLG_9215 = nextLine.Substring(27, 2); } catch { ffModel.AIR_CMTY_HDLG_9215 = ""; }
                                        try { ffModel.CLAS_RTNG_CD_2770 = nextLine.Substring(29, 1); } catch { ffModel.CLAS_RTNG_CD_2770 = ""; }
                                        try { ffModel.FRT_DESC_4020 = nextLine.Substring(30, 35); } catch { ffModel.FRT_DESC_4020 = ""; }

                                        sl.SetCellValue(gRow, Column, ffModel.INTGTY_CD_0864);
                                        sl.SetCellValue(gRow, ++Column, ffModel.ORIG_ACTY_CD_4210);
                                        sl.SetCellValue(gRow, ++Column, ffModel.RAIL_VARI_CD_4760);
                                        sl.SetCellValue(gRow, ++Column, ffModel.NMFC_2850);
                                        sl.SetCellValue(gRow, ++Column, ffModel.SUB_ITM_NBR_0861);
                                        sl.SetCellValue(gRow, ++Column, ffModel.UFC_CD_MODF_3040);
                                        sl.SetCellValue(gRow, ++Column, ffModel.HMC_2720);
                                        sl.SetCellValue(gRow, ++Column, ffModel.LCL_CD_2760);
                                        sl.SetCellValue(gRow, ++Column, ffModel.WRT_CMDTY_CD_9275);
                                        sl.SetCellValue(gRow, ++Column, ffModel.TYPE_CGO_CD_9260);
                                        sl.SetCellValue(gRow, ++Column, ffModel.SP_HDLG_CD_9240);
                                        sl.SetCellValue(gRow, ++Column, ffModel.AIR_DIM_CD_9220);
                                        sl.SetCellValue(gRow, ++Column, ffModel.AIR_CMTY_HDLG_9215);
                                        sl.SetCellValue(gRow, ++Column, ffModel.CLAS_RTNG_CD_2770);
                                        sl.SetCellValue(gRow, ++Column, ffModel.FRT_DESC_4020);
                                        gRow++;
                                        Column = 1;
                                    }
                                }
                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //SEG H
                            while (bookCounter <= numBooks && hRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.AddWorksheet(segH);  //RECORD TYPE 05
                                if (hRow == 1)
                                {
                                    //SET SEG_H COLUMN NAMES
                                    sl.SelectWorksheet(segH);
                                    sl.SetCellValue(hRow, Column, "MOE_CD_2833");
                                    sl.SetCellValue(hRow, ++Column, "SOS_CD_3690 OR SOSM_CD_2948");
                                    sl.SetCellValue(hRow, ++Column, "AAC_2507");
                                    sl.SetCellValue(hRow, ++Column, "QUP_6106");
                                    sl.SetCellValue(hRow, ++Column, "UI_3050");
                                    sl.SetCellValue(hRow, ++Column, "UNIT_PR_7075");
                                    sl.SetCellValue(hRow, ++Column, "SLC_2943");
                                    sl.SetCellValue(hRow, ++Column, "CIIC_2863");
                                    sl.SetCellValue(hRow, ++Column, "REP_CD_DLA_2934");
                                    sl.SetCellValue(hRow, ++Column, "REP_CD_CG_0709");
                                    sl.SetCellValue(hRow, ++Column, "ERRC_CD_AF_2655");
                                    sl.SetCellValue(hRow, ++Column, "RECOV_CD_MC_2891");
                                    sl.SetCellValue(hRow, ++Column, "RECOV_CD_AR_2892");
                                    sl.SetCellValue(hRow, ++Column, "MAT_CTL_NVY_2832");
                                    sl.SetCellValue(hRow, ++Column, "MAJ_MCC_AR_9256");
                                    sl.SetCellValue(hRow, ++Column, "AR_MCC_AP_SUB_2163");
                                    sl.SetCellValue(hRow, ++Column, "AR_MCC_USE_CD_2161");
                                    sl.SetCellValue(hRow, ++Column, "AR_MCC_SG_CD1_2167");
                                    sl.SetCellValue(hRow, ++Column, "ACTG_RQMT_AR_2665");
                                    sl.SetCellValue(hRow, ++Column, "FUND_CD_AF_2695");
                                    sl.SetCellValue(hRow, ++Column, "BDGT_CD_AF_3765");
                                    sl.SetCellValue(hRow, ++Column, "MMAC_AF_2836");
                                    sl.SetCellValue(hRow, ++Column, "PVC_AF_0858");
                                    sl.SetCellValue(hRow, ++Column, "STRS_ACT_MC_2959");
                                    sl.SetCellValue(hRow, ++Column, "CMBT_ESTL_MC_3311");
                                    sl.SetCellValue(hRow, ++Column, "MAT_MGMT_MC_9257");
                                    sl.SetCellValue(hRow, ++Column, "ECH_CD_MC_3150");
                                    sl.SetCellValue(hRow, ++Column, "MAT_IDEN_MC_4126");
                                    sl.SetCellValue(hRow, ++Column, "OPRTL_TST_MC_0572"); ;
                                    sl.SetCellValue(hRow, ++Column, "PHY_CTGY_MC_0573");
                                    sl.SetCellValue(hRow, ++Column, "COG_CD_NVY_2608");
                                    sl.SetCellValue(hRow, ++Column, "SMIC_NVY_2834");
                                    sl.SetCellValue(hRow, ++Column, "IRRC_NVY_0132");
                                    sl.SetCellValue(hRow, ++Column, "SP_MAT_CONT_0121");
                                    sl.SetCellValue(hRow, ++Column, "INV_ACT_CG_0708");
                                    sl.SetCellValue(hRow, ++Column, "SER_NO_CTL_CG_0763");
                                    sl.SetCellValue(hRow, ++Column, "SP_MAT_CONT_0121");
                                    sl.SetCellValue(hRow, ++Column, "EFF_DT_2128");
                                    sl.SetCellValue(hRow, ++Column, "USI_SVC_CD_0745");
                                    sl.SetCellValue(hRow, ++Column, "UI_CONV_FAC_3053");
                                    sl.SetCellValue(hRow, ++Column, "PHRS_CD_2862");
                                    sl.SetCellValue(hRow, ++Column, "PHRS_CD_PHRS_5240");
                                    sl.SetCellValue(hRow, ++Column, "QTY_PER_ASBL_0106");
                                    sl.SetCellValue(hRow, ++Column, "UNIT_MEAS_CD_0107"); ;
                                    sl.SetCellValue(hRow, ++Column, "OOU_0793");
                                    sl.SetCellValue(hRow, ++Column, "JTC_0792");
                                    hRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "05")
                                    {
                                        sl.SelectWorksheet(segH);
                                        try { ffModel.MOE_CD_2833 = nextLine.Substring(2, 2); } catch { ffModel.MOE_CD_2833 = ""; }
                                        try { ffModel.SOS_CD_3690_OR_SOSM_CD_2948 = nextLine.Substring(4, 3); } catch { ffModel.SOS_CD_3690_OR_SOSM_CD_2948 = ""; }
                                        try { ffModel.AAC_2507 = nextLine.Substring(7, 1); } catch { ffModel.AAC_2507 = ""; }
                                        try { ffModel.QUP_6106 = nextLine.Substring(8, 1); } catch { ffModel.QUP_6106 = ""; }
                                        try { ffModel.UI_3050 = nextLine.Substring(9, 2); } catch { ffModel.UI_3050 = ""; }
                                        var dollars = "";
                                        var dec = "";
                                        var cents = "";
                                        try { dollars = nextLine.Substring(11, 9); } catch { dollars = ""; }
                                        try { dec = nextLine.Substring(20, 1); } catch { dec = ""; }
                                        try { cents = nextLine.Substring(21, 2); } catch { cents = ""; }
                                        ffModel.UNIT_PR_7075 = dollars + dec + cents;
                                        try { ffModel.SLC_2943 = nextLine.Substring(23, 1); } catch { ffModel.SLC_2943 = ""; }
                                        try { ffModel.CIIC_2863 = nextLine.Substring(24, 1); } catch { ffModel.CIIC_2863 = ""; }
                                        var repCode = "";
                                        try { repCode = nextLine.Substring(25, 1); } catch { repCode = ""; }
                                        ffModel.REP_CD_DLA_2934 = repCode;
                                        ffModel.REP_CD_CG_0709 = repCode;
                                        ffModel.ERRC_CD_AF_2655 = repCode;
                                        ffModel.RECOV_CD_MC_2891 = repCode;
                                        ffModel.RECOV_CD_AR_2892 = repCode;
                                        ffModel.MAT_CTL_NVY_2832 = repCode;
                                        sl.SetCellValue(hRow, Column, ffModel.MOE_CD_2833);
                                        sl.SetCellValue(hRow, ++Column, ffModel.SOS_CD_3690_OR_SOSM_CD_2948);
                                        sl.SetCellValue(hRow, ++Column, ffModel.AAC_2507);
                                        sl.SetCellValue(hRow, ++Column, ffModel.QUP_6106);
                                        sl.SetCellValue(hRow, ++Column, ffModel.UI_3050);
                                        sl.SetCellValue(hRow, ++Column, ffModel.UNIT_PR_7075);
                                        sl.SetCellValue(hRow, ++Column, ffModel.SLC_2943);
                                        sl.SetCellValue(hRow, ++Column, ffModel.CIIC_2863);
                                        sl.SetCellValue(hRow, ++Column, ffModel.REP_CD_DLA_2934);
                                        sl.SetCellValue(hRow, ++Column, ffModel.REP_CD_CG_0709);
                                        sl.SetCellValue(hRow, ++Column, ffModel.ERRC_CD_AF_2655);
                                        sl.SetCellValue(hRow, ++Column, ffModel.RECOV_CD_MC_2891);
                                        sl.SetCellValue(hRow, ++Column, ffModel.RECOV_CD_AR_2892);
                                        sl.SetCellValue(hRow, ++Column, ffModel.MAT_CTL_NVY_2832);

                                        //MANAGEMENT CONTROL DATA BASED ON MILITARY BRANCH
                                        switch (ffModel.MOE_CD_2833)
                                        {
                                            //ARMY
                                            case "DA":
                                                Column = 15;
                                                var matCatCode = "";
                                                try { matCatCode = nextLine.Substring(26, 5); } catch { matCatCode = ""; }
                                                ffModel.MAJ_MCC_AR_9256 = matCatCode;
                                                ffModel.AR_MCC_AP_SUB_2163 = matCatCode;
                                                ffModel.AR_MCC_USE_CD_2161 = matCatCode;
                                                ffModel.AR_MCC_SG_CD1_2167 = matCatCode;
                                                try { ffModel.ACTG_RQMT_AR_2665 = nextLine.Substring(31, 1); } catch { ffModel.ACTG_RQMT_AR_2665 = ""; }
                                                sl.SetCellValue(hRow, Column, ffModel.MAJ_MCC_AR_9256);
                                                sl.SetCellValue(hRow, ++Column, ffModel.AR_MCC_AP_SUB_2163);
                                                sl.SetCellValue(hRow, ++Column, ffModel.AR_MCC_USE_CD_2161);
                                                sl.SetCellValue(hRow, ++Column, ffModel.AR_MCC_SG_CD1_2167);
                                                sl.SetCellValue(hRow, ++Column, ffModel.ACTG_RQMT_AR_2665);
                                                break;

                                            //AIR FORCE
                                            case "DF":
                                                Column = 20;
                                                try { ffModel.FUND_CD_AF_2695 = nextLine.Substring(26, 2); } catch { ffModel.FUND_CD_AF_2695 = ""; }
                                                try { ffModel.BDGT_CD_AF_3765 = nextLine.Substring(28, 1); } catch { ffModel.BDGT_CD_AF_3765 = ""; }
                                                try { ffModel.MMAC_AF_2836 = nextLine.Substring(29, 2); } catch { ffModel.MMAC_AF_2836 = ""; }
                                                try { ffModel.PVC_AF_0858 = nextLine.Substring(32, 1); } catch { ffModel.PVC_AF_0858 = ""; }
                                                sl.SetCellValue(hRow, Column, ffModel.FUND_CD_AF_2695);
                                                sl.SetCellValue(hRow, ++Column, ffModel.BDGT_CD_AF_3765);
                                                sl.SetCellValue(hRow, ++Column, ffModel.MMAC_AF_2836);
                                                sl.SetCellValue(hRow, ++Column, ffModel.PVC_AF_0858);
                                                break;

                                            //MARINES
                                            case "DM":
                                                Column = 24;
                                                try { ffModel.STRS_ACT_MC_2959 = nextLine.Substring(26, 1); } catch { ffModel.STRS_ACT_MC_2959 = ""; }
                                                try { ffModel.CMBT_ESTL_MC_3311 = nextLine.Substring(27, 1); } catch { ffModel.CMBT_ESTL_MC_3311 = ""; }
                                                try { ffModel.MAT_MGMT_MC_9257 = nextLine.Substring(28, 1); } catch { ffModel.MAT_MGMT_MC_9257 = ""; }
                                                try { ffModel.ECH_CD_MC_3150 = nextLine.Substring(29, 1); } catch { ffModel.ECH_CD_MC_3150 = ""; }
                                                try { ffModel.MAT_IDEN_MC_4126 = nextLine.Substring(30, 1); } catch { ffModel.MAT_IDEN_MC_4126 = ""; }
                                                try { ffModel.OPRTL_TST_MC_0572 = nextLine.Substring(31, 1); } catch { ffModel.OPRTL_TST_MC_0572 = ""; }
                                                try { ffModel.PHY_CTGY_MC_0573 = nextLine.Substring(32, 1); } catch { ffModel.PHY_CTGY_MC_0573 = ""; }
                                                sl.SetCellValue(hRow, Column, ffModel.STRS_ACT_MC_2959);
                                                sl.SetCellValue(hRow, ++Column, ffModel.CMBT_ESTL_MC_3311);
                                                sl.SetCellValue(hRow, ++Column, ffModel.MAT_MGMT_MC_9257);
                                                sl.SetCellValue(hRow, ++Column, ffModel.ECH_CD_MC_3150);
                                                sl.SetCellValue(hRow, ++Column, ffModel.MAT_IDEN_MC_4126);
                                                sl.SetCellValue(hRow, ++Column, ffModel.OPRTL_TST_MC_0572);
                                                sl.SetCellValue(hRow, ++Column, ffModel.PHY_CTGY_MC_0573);
                                                break;

                                            //NAVY
                                            case "DN":
                                                Column = 31;
                                                try { ffModel.COG_CD_NVY_2608 = nextLine.Substring(26, 2); } catch { ffModel.COG_CD_NVY_2608 = ""; }
                                                try { ffModel.SMIC_NVY_2834 = nextLine.Substring(28, 2); } catch { ffModel.SMIC_NVY_2834 = ""; }
                                                try { ffModel.IRRC_NVY_0132 = nextLine.Substring(30, 2); } catch { ffModel.IRRC_NVY_0132 = ""; }
                                                try { ffModel.SP_MAT_CONT_0121 = nextLine.Substring(32, 1); } catch { ffModel.SP_MAT_CONT_0121 = ""; }
                                                sl.SetCellValue(hRow, Column, ffModel.COG_CD_NVY_2608);
                                                sl.SetCellValue(hRow, ++Column, ffModel.SMIC_NVY_2834);
                                                sl.SetCellValue(hRow, ++Column, ffModel.IRRC_NVY_0132);
                                                sl.SetCellValue(hRow, ++Column, ffModel.SP_MAT_CONT_0121);
                                                break;

                                            //COAST GUARD
                                            case "GP":
                                                Column = 35;
                                                try { ffModel.INV_ACT_CG_0708 = nextLine.Substring(26, 1); } catch { ffModel.INV_ACT_CG_0708 = ""; }
                                                try { ffModel.SER_NO_CTL_CG_0763 = nextLine.Substring(28, 1); } catch { ffModel.SER_NO_CTL_CG_0763 = ""; }
                                                try { ffModel.SP_MAT_CONT_0121 = nextLine.Substring(29, 1); } catch { ffModel.SP_MAT_CONT_0121 = ""; }
                                                sl.SetCellValue(hRow, Column, ffModel.INV_ACT_CG_0708);
                                                sl.SetCellValue(hRow, ++Column, ffModel.SER_NO_CTL_CG_0763);
                                                sl.SetCellValue(hRow, ++Column, ffModel.SP_MAT_CONT_0121);
                                                break;
                                            default:
                                                Column = 38;
                                                break;
                                        }
                                        Column = 38;
                                        try { ffModel.EFF_DT_2128 = nextLine.Substring(33, 7); } catch { ffModel.EFF_DT_2128 = ""; }
                                        try { ffModel.USI_SVC_CD_0745 = nextLine.Substring(40, 1); } catch { ffModel.USI_SVC_CD_0745 = ""; }
                                        try { ffModel.UI_CONV_FAC_3053 = nextLine.Substring(41, 5); } catch { ffModel.UI_CONV_FAC_3053 = ""; }
                                        sl.SetCellValue(hRow, Column, ffModel.EFF_DT_2128);
                                        sl.SetCellValue(hRow, ++Column, ffModel.USI_SVC_CD_0745);
                                        sl.SetCellValue(hRow, ++Column, ffModel.UI_CONV_FAC_3053);

                                        var phraseCodeCounter = int.Parse(nextLine.Substring(46, 2));
                                        var text = new StringBuilder("");
                                        var index = 48;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.PHRS_CD_2862 = nextLine.Substring(index, 1));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        sl.SetCellValue(hRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 49;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.PHRS_CD_PHRS_5240 = nextLine.Substring(index, 36));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch{}
                                        }
                                        sl.SetCellValue(hRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 85;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.QTY_PER_ASBL_0106 = nextLine.Substring(index, 3));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch{}
                                        }
                                        sl.SetCellValue(hRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 88;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.UNIT_MEAS_CD_0107 = nextLine.Substring(index, 2));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch{}
                                        }
                                        sl.SetCellValue(hRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 90;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.OOU_0793 = nextLine.Substring(index, 3));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch{}
                                        }
                                        sl.SetCellValue(hRow, ++Column, text.ToString());
                                        text.Clear();
                                        index = 93;
                                        for (int i = 0; i < phraseCodeCounter; i++)
                                        {
                                            try
                                            {
                                                text.Append(ffModel.JTC_0792 = nextLine.Substring(index, 3));
                                                if (i + 1 != phraseCodeCounter)
                                                {
                                                    text.Append("|");
                                                    index += 48;
                                                }
                                            }
                                            catch{}
                                        }
                                        text.Clear();
                                        hRow++;
                                        Column = 1;
                                    }
                                }
                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //SEG W
                            while (bookCounter <= numBooks && wRow <= maxRows && !reader.EndOfStream)
                            {
                                sl.AddWorksheet(segW);  //RECORD TYPE 08
                                if (wRow == 1)
                                {
                                    sl.SetCellValue(wRow, Column, "PK_DTA_SRC_CD_5148");
                                    sl.SetCellValue(wRow, ++Column, "PICA_SICA_IND_5099");
                                    sl.SetCellValue(wRow, ++Column, "INTMD_CTN_QTY_5152");
                                    sl.SetCellValue(wRow, ++Column, "UP_WT_5153");
                                    sl.SetCellValue(wRow, ++Column, "UP_SZ_5154");
                                    sl.SetCellValue(wRow, ++Column, "UP_CU_5155");
                                    sl.SetCellValue(wRow, ++Column, "PKG_CTGY_CD_5159");
                                    sl.SetCellValue(wRow, ++Column, "ITM_TYP_STOR_5156");
                                    sl.SetCellValue(wRow, ++Column, "UNPKG_ITM_WT_5157");
                                    sl.SetCellValue(wRow, ++Column, "UNPKG_ITM_DIM_5158");
                                    sl.SetCellValue(wRow, ++Column, "MTHD_PRSRV_CD_5160");
                                    sl.SetCellValue(wRow, ++Column, "CLNG_DRY_PRC_5161");
                                    sl.SetCellValue(wRow, ++Column, "PRSRV_MAT_CD_5162");
                                    sl.SetCellValue(wRow, ++Column, "WRAP_MAT_CD_5163");
                                    sl.SetCellValue(wRow, ++Column, "CUSH_DUN_MAT_5164");
                                    sl.SetCellValue(wRow, ++Column, "THK_CUSH_DUN_5165");
                                    sl.SetCellValue(wRow, ++Column, "UNIT_CTNR_CD_5166");
                                    sl.SetCellValue(wRow, ++Column, "INTMD_CTNR_CD_5167");
                                    sl.SetCellValue(wRow, ++Column, "UNIT_CTNR_LVL_5168");
                                    sl.SetCellValue(wRow, ++Column, "SP_MKG_CD_5169");
                                    sl.SetCellValue(wRow, ++Column, "LVL_A_PKG_CD_5170");
                                    sl.SetCellValue(wRow, ++Column, "LVL_B_PKG_CD_5171");
                                    sl.SetCellValue(wRow, ++Column, "MINM_PK_RQ_CD_5172");
                                    sl.SetCellValue(wRow, ++Column, "OPTNL_PRO_IND_5173");
                                    sl.SetCellValue(wRow, ++Column, "SUPMTL_INST_5174");
                                    sl.SetCellValue(wRow, ++Column, "SPI_NBR_5175");
                                    sl.SetCellValue(wRow, ++Column, "SPI_REV_5176");
                                    sl.SetCellValue(wRow, ++Column, "SPI_DT_5177");
                                    sl.SetCellValue(wRow, ++Column, "CTNR_NSN_5178");
                                    sl.SetCellValue(wRow, ++Column, "PKG_DSGN_ACTY_5179");
                                    wRow++;
                                    Column = 1;
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * maxRows : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    var id = nextLine.Substring(0, 2);

                                    if (id == "08")
                                    {
                                        try { ffModel.PK_DTA_SRC_CD_5148 = nextLine.Substring(2, 1); } catch { ffModel.PK_DTA_SRC_CD_5148 = ""; }
                                        try { ffModel.PICA_SICA_IND_5099 = nextLine.Substring(3, 1); } catch { ffModel.PICA_SICA_IND_5099 = ""; }
                                        try { ffModel.INTMD_CTN_QTY_5152 = nextLine.Substring(4, 1); } catch { ffModel.INTMD_CTN_QTY_5152 = ""; }
                                        try { ffModel.UP_WT_5153 = nextLine.Substring(5, 5); } catch { ffModel.UP_WT_5153 = ""; }
                                        try { ffModel.UP_SZ_5154 = nextLine.Substring(10, 12); } catch { ffModel.UP_SZ_5154 = ""; }
                                        try { ffModel.UP_CU_5155 = nextLine.Substring(22, 7); } catch { ffModel.UP_CU_5155 = ""; }
                                        try { ffModel.PKG_CTGY_CD_5159 = nextLine.Substring(29, 4); } catch { ffModel.PKG_CTGY_CD_5159 = ""; }
                                        try { ffModel.ITM_TYP_STOR_5156 = nextLine.Substring(33, 1); } catch { ffModel.ITM_TYP_STOR_5156 = ""; }
                                        try { ffModel.UNPKG_ITM_WT_5157 = nextLine.Substring(34, 5); } catch { ffModel.UNPKG_ITM_WT_5157 = ""; }
                                        try { ffModel.UNPKG_ITM_DIM_5158 = nextLine.Substring(39, 12); } catch { ffModel.UNPKG_ITM_DIM_5158 = ""; }
                                        try { ffModel.MTHD_PRSRV_CD_5160 = nextLine.Substring(51, 2); } catch { ffModel.MTHD_PRSRV_CD_5160 = ""; }
                                        try { ffModel.CLNG_DRY_PRC_5161 = nextLine.Substring(53, 1); } catch { ffModel.CLNG_DRY_PRC_5161 = ""; }
                                        try { ffModel.PRSRV_MAT_CD_5162 = nextLine.Substring(54, 2); } catch { ffModel.PRSRV_MAT_CD_5162 = ""; }
                                        try { ffModel.WRAP_MAT_CD_5163 = nextLine.Substring(56, 2); } catch { ffModel.WRAP_MAT_CD_5163 = ""; }
                                        try { ffModel.CUSH_DUN_MAT_5164 = nextLine.Substring(58, 2); } catch { ffModel.CUSH_DUN_MAT_5164 = ""; }
                                        try { ffModel.THK_CUSH_DUN_5165 = nextLine.Substring(60, 1); } catch { ffModel.THK_CUSH_DUN_5165 = ""; }
                                        try { ffModel.UNIT_CTNR_CD_5166 = nextLine.Substring(62, 2); } catch { ffModel.UNIT_CTNR_CD_5166 = ""; }
                                        try { ffModel.INTMD_CTNR_CD_5167 = nextLine.Substring(64, 2); } catch { ffModel.INTMD_CTNR_CD_5167 = ""; }
                                        try { ffModel.UNIT_CTNR_LVL_5168 = nextLine.Substring(66, 1); } catch { ffModel.UNIT_CTNR_LVL_5168 = ""; }
                                        try { ffModel.SP_MKG_CD_5169 = nextLine.Substring(67, 2); } catch { ffModel.SP_MKG_CD_5169 = ""; }
                                        try { ffModel.LVL_A_PKG_CD_5170 = nextLine.Substring(69, 1); } catch { ffModel.LVL_A_PKG_CD_5170 = ""; }
                                        try { ffModel.LVL_B_PKG_CD_5171 = nextLine.Substring(70, 1); } catch { ffModel.LVL_B_PKG_CD_5171 = ""; }
                                        try { ffModel.MINM_PK_RQ_CD_5172 = nextLine.Substring(72, 1); } catch { ffModel.MINM_PK_RQ_CD_5172 = ""; }
                                        try { ffModel.OPTNL_PRO_IND_5173 = nextLine.Substring(73, 1); } catch { ffModel.OPTNL_PRO_IND_5173 = ""; }

                                        //GET PIPE DELIMITER INDEX
                                        var line = nextLine.Substring(74);
                                        var newLength = line.IndexOf('|');
                                        try { ffModel.SUPMTL_INST_5174 = nextLine.Substring(74, newLength); } catch { ffModel.SUPMTL_INST_5174 = ""; }
                                        try { ffModel.SPI_NBR_5175 = nextLine.Substring(newLength + 75, 10); } catch { ffModel.SPI_NBR_5175 = ""; }
                                        try { ffModel.SPI_REV_5176 = nextLine.Substring(newLength + 85, 1); } catch { ffModel.SPI_REV_5176 = ""; }
                                        try { ffModel.SPI_DT_5177 = nextLine.Substring(newLength + 86, 5); } catch { ffModel.SPI_DT_5177 = ""; }
                                        try { ffModel.CTNR_NSN_5178 = nextLine.Substring(newLength + 91, 13); } catch { ffModel.CTNR_NSN_5178 = ""; }
                                        try { ffModel.PKG_DSGN_ACTY_5179 = nextLine.Substring(newLength + 104, 5); } catch { ffModel.PKG_DSGN_ACTY_5179 = ""; }
                                        sl.SetCellValue(wRow, Column, ffModel.PK_DTA_SRC_CD_5148);
                                        sl.SetCellValue(wRow, ++Column, ffModel.PICA_SICA_IND_5099);
                                        sl.SetCellValue(wRow, ++Column, ffModel.INTMD_CTN_QTY_5152);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UP_WT_5153);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UP_SZ_5154);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UP_CU_5155);
                                        sl.SetCellValue(wRow, ++Column, ffModel.PKG_CTGY_CD_5159);
                                        sl.SetCellValue(wRow, ++Column, ffModel.ITM_TYP_STOR_5156);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UNPKG_ITM_WT_5157);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UNPKG_ITM_DIM_5158);
                                        sl.SetCellValue(wRow, ++Column, ffModel.MTHD_PRSRV_CD_5160);
                                        sl.SetCellValue(wRow, ++Column, ffModel.CLNG_DRY_PRC_5161);
                                        sl.SetCellValue(wRow, ++Column, ffModel.PRSRV_MAT_CD_5162);
                                        sl.SetCellValue(wRow, ++Column, ffModel.WRAP_MAT_CD_5163);
                                        sl.SetCellValue(wRow, ++Column, ffModel.CUSH_DUN_MAT_5164);
                                        sl.SetCellValue(wRow, ++Column, ffModel.THK_CUSH_DUN_5165);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UNIT_CTNR_CD_5166);
                                        sl.SetCellValue(wRow, ++Column, ffModel.INTMD_CTNR_CD_5167);
                                        sl.SetCellValue(wRow, ++Column, ffModel.UNIT_CTNR_LVL_5168);
                                        sl.SetCellValue(wRow, ++Column, ffModel.SP_MKG_CD_5169);
                                        sl.SetCellValue(wRow, ++Column, ffModel.LVL_A_PKG_CD_5170);
                                        sl.SetCellValue(wRow, ++Column, ffModel.LVL_B_PKG_CD_5171);
                                        sl.SetCellValue(wRow, ++Column, ffModel.MINM_PK_RQ_CD_5172);
                                        sl.SetCellValue(wRow, ++Column, ffModel.OPTNL_PRO_IND_5173);
                                        sl.SetCellValue(wRow, ++Column, ffModel.SUPMTL_INST_5174);
                                        sl.SetCellValue(wRow, ++Column, ffModel.SPI_NBR_5175);
                                        sl.SetCellValue(wRow, ++Column, ffModel.SPI_REV_5176);
                                        sl.SetCellValue(wRow, ++Column, ffModel.SPI_DT_5177);
                                        sl.SetCellValue(wRow, ++Column, ffModel.CTNR_NSN_5178);
                                        sl.SetCellValue(wRow, ++Column, ffModel.PKG_DSGN_ACTY_5179);
                                        wRow++;
                                        Column = 1;
                                    }
                                }
                            }
                            reader.BaseStream.Position = 0;
                            reader.DiscardBufferedData();
                            FastForward = true;

                            //STEP 3
                            try
                            {
                                Step_3(sl);
                                aRow = 1;
                                bRow = 1;
                                cRow = 1;
                                eRow = 1;
                                gRow = 1;
                                hRow = 1;
                                wRow = 1;
                                bookCounter++;
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                                System.Windows.Forms.Application.Exit();
                            }
                        }
                    }
                }
                SuccessMessage();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
        #endregion

        #region FCAN-SEGK
        private void FCANSEGK()
        {
            //Get # Books needed that is less than max excel row (1,000,000)
            var numBooks = TotalRows / 250000;
            //Number of excel books
            var bookCounter = 0;
            var totalCounter = 0;
            Row = 1;
            Column = 1;

            try
            {
                while (bookCounter <= numBooks)
                {
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;
                    var fsModel = new FCAN_SEGK();

                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        UpdateSheetNames(bookCounter);

                        using (SLDocument sl = new SLDocument())
                        {
                            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, SlSheetName);
                            while (bookCounter <= numBooks && SheetCounter < 250000 && !reader.EndOfStream &&
                                   totalCounter <= TotalRows)
                            {
                                //STEP 1
                                if (SheetCounter == 0)
                                {
                                    sl.SetCellValue(Row, Column, "RECORD_TYPE");
                                    sl.SetCellValue(Row, ++Column, "FSC");
                                    sl.SetCellValue(Row, ++Column, "NIIN");
                                    sl.SetCellValue(Row, ++Column, "REPLACEMENT_NIIN");
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * 250000 : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }
                                        FastForward = false;
                                    }
                                    var nextLine = reader.ReadLine();
                                    fsModel.RECORD_TYPE = nextLine.Substring(0, 1);
                                    fsModel.FSC = nextLine.Substring(1, 4);
                                    fsModel.NIIN = nextLine.Substring(5, 9);
                                    fsModel.REPLACEMENT_NSN = !string.IsNullOrEmpty(nextLine.Substring(14)) ? nextLine.Substring(14) : "";
                                    sl.SetCellValue(Row, Column, fsModel.RECORD_TYPE);
                                    sl.SetCellValue(Row, ++Column, fsModel.FSC);
                                    sl.SetCellValue(Row, ++Column, fsModel.NIIN);
                                    sl.SetCellValue(Row, ++Column, fsModel.REPLACEMENT_NSN);

                                }

                                //STEP 2
                                Step_2();

                                //INCREMENT TOTAL SHEET ROW COUNTER
                                totalCounter++;
                            }
                            //STEP 3
                            try
                            {
                                Step_3(sl);
                                bookCounter++;
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                                                   MessageBoxIcon.Exclamation);
                            }
                        }
                    }
                }
                SuccessMessage();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
        #endregion

        #region STEP 3

        private void Step_3(SLDocument sl)
        {
            //RESET FOR NEW SHEET
            Row = 1;
            Column = 1;

            //Reset sheet counter...should never go over 250,000 rows
            SheetCounter = 0;

            //Re-initiate Fast Forward counter
            FastForward = true;

            try
            {
                CheckIfFolderExists();
                CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                   DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                   ".xlsx";

                sl.SaveAs(CompleteFilePath);
                SendToTimerFile();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                System.Windows.Forms.Application.Exit();
            }
        }

        #endregion

        #region Step_2

        private void Step_2()
        {
            //INCREMENT SHEET ROW COUNTER
            SheetCounter++;

            //RESET COLUMN COUNT
            Column = 1;

            //INCREMENT ROW
            Row++;
        }

        #endregion

        #region FCAGE

        private void FCAGE()
        {
            //Get # Books needed that is less than max excel row (1,000,000)
            var numBooks = TotalRows / 250000;

            //Number of excel books
            var bookCounter = 0;
            var totalCounter = 0;
            Row = 1;
            Column = 1;

            var fCageList = new[]
            {
                "CAGE_CODE", "COMPANY_NAME_1", "COMPANY_NAME_2", "COMPANY_NAME_3", "COMPANY_NAME_4",
                "COMPANY_NAME_5", "DOMESTIC_STREET_ADDRESS_1",
                "DOMESTIC_STREET_ADDRESS_2", "DOMESTIC_POST_OFFICE_BOX", "DOMESTIC_CITY", "DOMESTIC_STATE",
                "DOMESTIC_ZIP_CODE",
                "DOMESTIC_COUNTRY", "DOMESTIC_PHONE_FAX_NUMBER_1", "DOMESTIC_PHONE_FAX_NUMBER_2","FOREIGN_STREET_ADDRESS_1", "FOREIGN_STREET_ADDRESS_2",
                "FOREIGN_POST_OFFICE_BOX", "FOREIGN_CITY", "FOREIGN_PROVINCE", "FOREIGN_COUNTRY",
                "FOREIGN_POSTAL_ZONE",
                "FOREIGN_PHONE_NUMBER", "FOREIGN_FAX_NUMBER", "CAO_CODE", "ADP_CODE", "STATUS_CODE",
                "ASSOCIATION_CODE", "TYPE_CODE", "AFFILIATION_CODE", "SIZE_OF_BUSINESS_CODE",
                "PRIMARY_BUSINESS_CATEGORY", "TYPE_OF_BUSINESS_CODE", "WOMAN_OWNED_BUSINESS",
                "STANDARD_INDUSTRIAL_1",
                "STANDARD_INDUSTRIAL_2", "STANDARD_INDUSTRIAL_3", "STANDARD_INDUSTRIAL_4", "REPLACEMENT_CAGE",
                "FORMER_NAME_1", "FORMER_NAME_2"
            };

            try
            {
                while (bookCounter <= numBooks)
                {
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;
                    //NewXlApp();


                    //Initialize class
                    var fcModel = new FCAGE();
                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        UpdateSheetNames(bookCounter);

                        using (SLDocument sl = new SLDocument())
                        {
                            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, SlSheetName);

                            while (bookCounter <= numBooks && SheetCounter < 250000 && !reader.EndOfStream &&
                                   totalCounter <= TotalRows)
                            {
                                if (SheetCounter == 0)
                                {
                                    sl.SetCellValue(Row, Column, "CAGE_CODE");

                                    //Worksheet.Cells[Row, Column] = "CAGE_CODE";
                                    for (int i = 1; i < fCageList.Length; i++)
                                    {
                                        sl.SetCellValue(Row, i + 1, fCageList[i]);
                                        //Worksheet.Cells[Row, i + 1] = fCageList[i];
                                    }
                                }
                                else
                                {
                                    while (FastForward)
                                    {
                                        //Used to 'jump' to row
                                        var fastForwardCounter = bookCounter > 0 ? bookCounter * 250000 : 0;

                                        //Jump to row first
                                        for (int i = 0; i < fastForwardCounter; i++)
                                        {
                                            reader.ReadLine();
                                        }

                                        FastForward = false;
                                    }

                                    var nextLine = reader.ReadLine();

                                    string[] strList = nextLine.Split('|');

                                    fcModel.CAGE_CODE = strList[0];
                                    fcModel.COMPANY_NAME_1 = strList[1];
                                    fcModel.COMPANY_NAME_2 = strList[2];
                                    fcModel.COMPANY_NAME_3 = strList[3];
                                    fcModel.COMPANY_NAME_4 = strList[4];
                                    fcModel.COMPANY_NAME_5 = strList[5];

                                    sl.SetCellValue(Row, Column, "\t" + fcModel.CAGE_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.COMPANY_NAME_1.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.COMPANY_NAME_2.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.COMPANY_NAME_3.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.COMPANY_NAME_4.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.COMPANY_NAME_5.Trim());

                                    var addressIndicator = int.Parse(strList[6].Substring(0, 1));
                                    if (addressIndicator == 1)
                                    {
                                        fcModel.DOMESTIC_STREET_ADDRESS_1 = strList[6];
                                        fcModel.DOMESTIC_STREET_ADDRESS_2 = strList[7];
                                        fcModel.DOMESTIC_POST_OFFICE_BOX = strList[8];
                                        fcModel.DOMESTIC_CITY = strList[9];
                                        fcModel.DOMESTIC_STATE = strList[10];
                                        fcModel.DOMESTIC_ZIP_CODE = strList[11];
                                        fcModel.DOMESTIC_COUNTRY = strList[12];
                                        fcModel.DOMESTIC_PHONE_FAX_NUMBER_1 = strList[13];
                                        fcModel.DOMESTIC_PHONE_FAX_NUMBER_2 = strList[14];

                                        fcModel.FOREIGN_STREET_ADDRESS_1 = "";
                                        fcModel.FOREIGN_STREET_ADDRESS_2 = "";
                                        fcModel.FOREIGN_POST_OFFICE_BOX = "";
                                        fcModel.FOREIGN_CITY = "";
                                        fcModel.FOREIGN_PROVINCE = "";
                                        fcModel.FOREIGN_COUNTRY = "";
                                        fcModel.FOREIGN_POSTAL_ZONE = "";
                                        fcModel.FOREIGN_PHONE_NUMBER = "";
                                        fcModel.FOREIGN_FAX_NUMBER = "";

                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_STREET_ADDRESS_1.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_STREET_ADDRESS_2.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_POST_OFFICE_BOX.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_CITY.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_STATE.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_ZIP_CODE.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_COUNTRY.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_PHONE_FAX_NUMBER_1.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.DOMESTIC_PHONE_FAX_NUMBER_2.Trim());

                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                    }
                                    else
                                    {
                                        fcModel.DOMESTIC_STREET_ADDRESS_1 = "";
                                        fcModel.DOMESTIC_STREET_ADDRESS_2 = "";
                                        fcModel.DOMESTIC_POST_OFFICE_BOX = "";
                                        fcModel.DOMESTIC_CITY = "";
                                        fcModel.DOMESTIC_STATE = "";
                                        fcModel.DOMESTIC_ZIP_CODE = "";
                                        fcModel.DOMESTIC_COUNTRY = "";
                                        fcModel.DOMESTIC_PHONE_FAX_NUMBER_1 = "";
                                        fcModel.DOMESTIC_PHONE_FAX_NUMBER_2 = "";

                                        fcModel.FOREIGN_STREET_ADDRESS_1 = strList[6];
                                        fcModel.FOREIGN_STREET_ADDRESS_2 = strList[7];
                                        fcModel.FOREIGN_POST_OFFICE_BOX = strList[8];
                                        fcModel.FOREIGN_CITY = strList[9];
                                        fcModel.FOREIGN_PROVINCE = strList[10];
                                        fcModel.FOREIGN_COUNTRY = strList[11];
                                        fcModel.FOREIGN_POSTAL_ZONE = strList[12];
                                        fcModel.FOREIGN_PHONE_NUMBER = strList[13];
                                        fcModel.FOREIGN_FAX_NUMBER = strList[14];

                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");
                                        sl.SetCellValue(Row, ++Column, "");

                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_STREET_ADDRESS_1.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_STREET_ADDRESS_2.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_POST_OFFICE_BOX.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_CITY.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_PROVINCE.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_COUNTRY.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_POSTAL_ZONE.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_PHONE_NUMBER.Trim());
                                        sl.SetCellValue(Row, ++Column, "\t" + fcModel.FOREIGN_FAX_NUMBER.Trim());
                                    }

                                    fcModel.CAO_CODE = strList[15];
                                    fcModel.ADP_CODE = strList[16];
                                    fcModel.STATUS_CODE = strList[17];
                                    fcModel.ASSOCIATION_CODE = strList[18];
                                    fcModel.TYPE_CODE = strList[19];
                                    fcModel.AFFILIATION_CODE = strList[20];
                                    fcModel.SIZE_OF_BUSINESS_CODE = strList[21];
                                    fcModel.PRIMARY_BUSINESS_CATEGORY = strList[22];
                                    fcModel.TYPE_OF_BUSINESS_CODE = strList[23];
                                    fcModel.WOMAN_OWNED_BUSINESS = strList[24];
                                    fcModel.STANDARD_INDUSTRIAL_1 = strList[25];
                                    fcModel.STANDARD_INDUSTRIAL_2 = strList[26];
                                    fcModel.STANDARD_INDUSTRIAL_3 = strList[27];
                                    fcModel.STANDARD_INDUSTRIAL_4 = strList[28];
                                    fcModel.REPLACEMENT_CAGE = strList[29];
                                    fcModel.FORMER_NAME_1 = strList[30];
                                    fcModel.FORMER_NAME_2 = strList[31];

                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.CAO_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.ADP_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.STATUS_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.ASSOCIATION_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.TYPE_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.AFFILIATION_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.SIZE_OF_BUSINESS_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.PRIMARY_BUSINESS_CATEGORY.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.TYPE_OF_BUSINESS_CODE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.WOMAN_OWNED_BUSINESS.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.STANDARD_INDUSTRIAL_1.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.STANDARD_INDUSTRIAL_2.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.STANDARD_INDUSTRIAL_3.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.STANDARD_INDUSTRIAL_4.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.REPLACEMENT_CAGE.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.FORMER_NAME_1.Trim());
                                    sl.SetCellValue(Row, ++Column, "\t" + fcModel.FORMER_NAME_2.Trim());
                                }

                                //INCREMENT STEP FOR PROGRESS BAR
                                pb.PerformStep();

                                //INCREMENT SHEET ROW COUNTER
                                SheetCounter++;

                                //INCREMENT TOTAL SHEET ROW COUNTER
                                totalCounter++;

                                //RESET COLUMN COUNT
                                Column = 1;

                                //INCREMENT ROW
                                Row++;
                            }

                            //RESET FOR NEW SHEET
                            //Reset Row
                            Row = 1;
                            Column = 1;

                            //Reset sheet counter...should never go over 250,000 rows
                            SheetCounter = 0;

                            //Re-initiate Fast Forward counter
                            FastForward = true;

                            try
                            {
                                CheckIfFolderExists();
                                CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                                   DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                                   ".xlsx";

                                sl.SaveAs(CompleteFilePath);
                                SendToTimerFile();
                                bookCounter++;

                                //CloseExcelApplication();
                                //MessageBox.Show("Book " + bookCounter + " finished.");
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                                System.Windows.Forms.Application.Exit();
                            }
                        }

                    }
                }

                SuccessMessage();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        #endregion

        #region Check If Folder Exists

        public void CheckIfFolderExists()
        {
            bool exists = Directory.Exists(FolderDestination);

            if (!exists)
                Directory.CreateDirectory(FolderDestination);
        }



        #endregion

        #region COLXREF

        private void COLXREF()
        {
            var totalCounter = 0;
            try
            {
                TimeStart = DateTime.Now;
                lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                NewXlApp();
                var cxModel = new COLXREF();
                UpdateSheetNames(0);

                using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                using (var reader = new StreamReader(stream))
                {
                    while (!reader.EndOfStream)
                    {
                        if (SheetCounter == 0)
                        {
                            Worksheet.Cells[Row, Column] = "NAME";
                            Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE_1";
                            Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE_2";
                            Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE_3";
                            Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE_4";
                            Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE_5";
                        }
                        else
                        {
                            var nextline = reader.ReadLine();
                            cxModel.NAME = nextline.Substring(0, 32);
                            try { cxModel.ITEM_NAME_CODE_1 = nextline.Substring(32, 5); } catch { cxModel.ITEM_NAME_CODE_1 = ""; }
                            try { cxModel.ITEM_NAME_CODE_2 = nextline.Substring(38, 5); } catch { cxModel.ITEM_NAME_CODE_2 = ""; }
                            try { cxModel.ITEM_NAME_CODE_3 = nextline.Substring(43, 5); } catch { cxModel.ITEM_NAME_CODE_3 = ""; }
                            try { cxModel.ITEM_NAME_CODE_4 = nextline.Substring(48, 5); } catch { cxModel.ITEM_NAME_CODE_4 = ""; }
                            try { cxModel.ITEM_NAME_CODE_5 = nextline.Substring(53, 5); } catch { cxModel.ITEM_NAME_CODE_5 = ""; }

                            //ADD TO WORK SHEET
                            Worksheet.Cells[Row, Column] = "\t" + cxModel.NAME.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cxModel.ITEM_NAME_CODE_1.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cxModel.ITEM_NAME_CODE_2.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cxModel.ITEM_NAME_CODE_3.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cxModel.ITEM_NAME_CODE_4.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cxModel.ITEM_NAME_CODE_5.Trim();
                        }

                        ;
                        //INCREMENT STEP FOR PROGRESS BAR
                        pb.PerformStep();

                        //INCREMENT SHEET ROW COUNTER
                        SheetCounter++;

                        //INCREMENT TOTAL SHEET ROW COUNTER
                        totalCounter++;

                        //RESET COLUMN COUNT
                        Column = 1;

                        //INCREMENT ROW
                        Row++;
                    }

                    try
                    {
                        CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                           DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                           ".xlsx";

                        CloseExcelApplication();
                        //MessageBox.Show("Book " + bookCounter + " finished.");
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                        System.Windows.Forms.Application.Exit();
                    }
                }

                SuccessMessage();
            }

            catch (Exception e)
            {
                MessageBox.Show("Error...\r\n" + e.ToString(), "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Check For Empty Fields

        private bool CheckForEmptyFields(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cmbTableName.Text))
            {
                MessageBox.Show("Choose Table", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbTableName.Focus();
                return false;
            }

            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("Need File Path", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BtnFileLocation_Click(sender, e);
                return false;
            }

            if (string.IsNullOrEmpty(txtFolderDestination.Text))
            {
                MessageBox.Show("Need Folder Path", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BtnFolderDestination_Click(sender, e);
                return false;
            }

            return true;
        }

        #endregion

        #region CHARDAT

        private void CHARDAT()
        {
            //Get # Books needed that is less than max excel row (1,000,000)
            var numBooks = TotalRows / 1000000;
            //Number of excel books
            var bookCounter = 0;
            var totalCounter = 0;

            try
            {
                while (bookCounter <= numBooks)
                {
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;
                    NewXlApp();

                    //Initialize class
                    var cdModel = new CHARDAT();

                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        UpdateSheetNames(bookCounter);

                        while (bookCounter <= numBooks && SheetCounter < 1000000 && !reader.EndOfStream &&
                               totalCounter <= TotalRows)
                        {
                            if (SheetCounter == 0)
                            {
                                Worksheet.Cells[Row, Column] = "FEDERAL_SUPPLY_CLASS";
                                Worksheet.Cells[Row, ++Column] = "NATIONAL_STOCK_NUMBER";
                                Worksheet.Cells[Row, ++Column] = "ITEM_NAME_CODE";
                                Worksheet.Cells[Row, ++Column] = "ITEM_NAME";
                                Worksheet.Cells[Row, ++Column] = "MRC|DECODED_MRC|MRC_REPLY";
                                Worksheet.Cells[Row, ++Column] = "ENAC_CODES";
                            }
                            else
                            {
                                while (FastForward)
                                {
                                    //Used to 'jump' to row
                                    //var fastForwardCounter = bookCounter > 0 ? bookCounter * 1000000 : 0;
                                    var fastForwardCounter = 1000000;

                                    //Jump to row first
                                    for (int i = 0; i < fastForwardCounter; i++)
                                    {
                                        reader.ReadLine();
                                    }

                                    FastForward = false;
                                }

                                var nextLine = reader.ReadLine();
                                var numMrcs = 0;
                                var fullMrc = "";
                                var fullEnac = "";
                                var decodedMrcLength = 0;
                                var mrcReplyLength = 0;
                                var startIndex = 0;

                                cdModel.FEDERAL_SUPPLY_CLASS = nextLine.Substring(0, 4);
                                cdModel.NATIONAL_STOCK_NUMBER = nextLine.Substring(4, 9);
                                cdModel.ITEM_NAME_CODE = nextLine.Substring(13, 5);

                                var itemCharacters = int.Parse(nextLine.Substring(18, 2));
                                cdModel.ITEM_NAME = nextLine.Substring(20, itemCharacters);
                                numMrcs = int.Parse(nextLine.Substring(itemCharacters + 20, 4));

                                startIndex = itemCharacters + 20 + 4;

                                //MULTIPLE MRC | DECODED MRCs | MRC REPLIES - PIPELINE DELIMITED
                                for (int i = 0; i < numMrcs; i++)
                                {
                                    fullMrc = nextLine.Substring(startIndex, 4) + " | ";
                                    try
                                    {
                                        decodedMrcLength = int.Parse(nextLine.Substring(startIndex + 4, 4));
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show("Error\n\n" +
                                                        "numMrcs: " + numMrcs +
                                                        "\n\rfullMrc: " + fullMrc +
                                                        "\n\rstartIndex: " + startIndex + "\n\r" +
                                                        "\n\rRow: " + Row +
                                                        "\n\rnextLine.Substring(" + startIndex + " + 4): \n\r" +
                                                        nextLine.Substring(startIndex));
                                    }

                                    var decodedMrc = nextLine.Substring(startIndex + 8, decodedMrcLength);
                                    fullMrc += decodedMrc + " | ";

                                    mrcReplyLength = int.Parse(nextLine.Substring(startIndex + 8 + decodedMrcLength,
                                        4));

                                    var fullMrcReply = nextLine.Substring(startIndex + 8 + decodedMrcLength + 4);
                                    fullMrc += fullMrcReply + " | ";

                                    startIndex = startIndex + 12 + decodedMrcLength + mrcReplyLength;
                                }

                                //MULTIPLE ENAC CODES
                                var enacCounter = 0;
                                var enacStartIndex = 2;

                                //Go to next line
                                nextLine = reader.ReadLine();

                                enacCounter = int.Parse(nextLine.Substring(0, 2));

                                for (int i = 0; i < enacCounter; i++)
                                {
                                    //IF FIRST ROUND OF ENAC CODES or LAST ENAC CODE
                                    if (i == 0)
                                    {
                                        fullEnac += fullEnac + nextLine.Substring(enacStartIndex, 2);
                                        enacStartIndex += 2;
                                        continue;
                                    }

                                    if (enacCounter - 1 == i && i > 0)
                                    {
                                        fullEnac += fullEnac + nextLine.Substring(enacStartIndex, 2);
                                    }
                                    else
                                    {
                                        fullEnac += fullEnac + nextLine.Substring(enacStartIndex, 2) + " | ";
                                    }

                                    enacStartIndex = +2;
                                }

                                //ADD TO WORK SHEET
                                Worksheet.Cells[Row, Column] = "\t" + cdModel.FEDERAL_SUPPLY_CLASS.Trim();
                                Worksheet.Cells[Row, ++Column] = "\t" + cdModel.NATIONAL_STOCK_NUMBER.Trim();
                                Worksheet.Cells[Row, ++Column] = "\t" + cdModel.ITEM_NAME_CODE.Trim();
                                Worksheet.Cells[Row, ++Column] = "\t" + cdModel.ITEM_NAME.Trim();
                                Worksheet.Cells[Row, ++Column] = "\t" + fullMrc.Trim();
                                Worksheet.Cells[Row, ++Column] = "\t" + fullEnac.Trim();
                            }

                            //INCREMENT STEP FOR PROGRESS BAR
                            pb.PerformStep();

                            //INCREMENT SHEET ROW COUNTER
                            SheetCounter++;

                            //INCREMENT TOTAL SHEET ROW COUNTER
                            totalCounter++;

                            //RESET COLUMN COUNT
                            Column = 1;

                            //INCREMENT ROW
                            Row++;
                        }

                        //RESET FOR NEW SHEET
                        //Reset Row
                        Row = 1;

                        //Reset sheet counter...should never go over 1,000,000 rows
                        SheetCounter = 0;

                        //Re-initiate Fast Forward counter
                        FastForward = true;

                        try
                        {
                            CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                               DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                               ".xlsx";
                            bookCounter++;

                            CloseExcelApplication();
                            //MessageBox.Show("Book " + bookCounter + " finished.");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                            System.Windows.Forms.Application.Exit();
                        }
                    }
                }

                SuccessMessage();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error...\r\n" + e.ToString(), "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region CAGE_Date_Est_and_Changed

        private void CAGEDateEstAndChgd()
        {
            //Get # Books needed that is less than max excel row (1,000,000)
            var numBooks = TotalRows / 1000000;

            //Number of excel books
            var bookCounter = 0;
            var totalCounter = 0;

            try
            {
                while (bookCounter <= numBooks)
                {
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;

                    NewXlApp();

                    //Initialize class
                    var cModel = new CAGE_DATE_EST_AND_CHGD();

                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        UpdateSheetNames(bookCounter);

                        lblCurrentDoc.Text = "Working On: " + SheetName;

                        while (bookCounter <= numBooks && SheetCounter < 1000000 && !reader.EndOfStream &&
                               totalCounter <= TotalRows)
                        {
                            if (SheetCounter == 0)
                            {
                                cModel.CAGE_CODE = "CAGE";
                                cModel.CAGE_RECORD_ESTABLISHED = "DT_EST";
                                cModel.CAGE_RECORD_LAST_CHANGED = "DT_CHGD";
                            }
                            else
                            {
                                while (FastForward)
                                {
                                    //Used to 'jump' to row
                                    var fastForwardCounter = bookCounter > 0 ? bookCounter * 1000000 : 0;

                                    //Jump to row first
                                    for (int i = 0; i < fastForwardCounter; i++)
                                    {
                                        reader.ReadLine();
                                    }

                                    FastForward = false;
                                }

                                var nextLine = reader.ReadLine();

                                cModel.CAGE_CODE = nextLine.Substring(0, 5);
                                cModel.CAGE_RECORD_ESTABLISHED = nextLine.Substring(6, 7);
                                cModel.CAGE_RECORD_LAST_CHANGED = nextLine.Substring(14);
                            }

                            Worksheet.Cells[Row, Column] = "\t" + cModel.CAGE_CODE.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cModel.CAGE_RECORD_ESTABLISHED.Trim();
                            Worksheet.Cells[Row, ++Column] = "\t" + cModel.CAGE_RECORD_LAST_CHANGED.Trim();

                            pb.PerformStep();
                            lblProgress.Text = "Row: " + (totalCounter - 1) + " of " + TotalRows;
                            SheetCounter++;
                            totalCounter++;
                            Row++;
                            Column = 1;
                        }

                        //Reset Row
                        Row = 1;

                        //Reset sheet counter...should never go over 1,000,000 rows
                        SheetCounter = 0;

                        //Re-initiate Fast Forward counter
                        FastForward = true;

                        try
                        {
                            CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                               DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                               ".xlsx";
                            bookCounter++;

                            SendToTimerFile();
                            Workbook.Close(true, CompleteFilePath, Missing.Value); //close and save individual book
                            XlApp.Quit();
                            Marshal.ReleaseComObject(Worksheet);
                            Marshal.ReleaseComObject(Workbook);
                            Marshal.ReleaseComObject(XlApp);

                            //MessageBox.Show("Book " + bookCounter + " finished.");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                            System.Windows.Forms.Application.Exit();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error...\n\r" + e.ToString(), "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

            }
        }

        #endregion

        #region Update Sheet Names

        private void UpdateSheetNames(int bookCounter)
        {
            SheetName = cmbTableName.Text;
            SheetName = SheetName + "_" + (bookCounter + 1);
            //Worksheet.Name = SheetName;
            SlSheetName = SheetName;
            lblCurrentDoc.Text = "Working On: " + SheetName;
        }

        #endregion

        #region CAGECDS

        public void CAGECDS()
        {
            //Get # Books needed that is less than max excel row (1,000,000)
            var numBooks = TotalRows / 1000000;

            //Number of excel books
            var bookCounter = 0;
            var totalCounter = 0;
            var fastForwardCounter = 0;
            string nextline = "";
            string type = "";
            CageObject co;

            try
            {
                while (bookCounter <= numBooks)
                {
                    TimeStart = DateTime.Now;
                    lblStart.Text = "Start: " + TimeStart.ToString("h:mm:ss tt");
                    FastForward = true;

                    co = new CageObject { App = new Application() };
                    co.xBook = co.App.Workbooks.Add();

                    //Worksheets
                    co.xCageSheet = co.xBook.Worksheets.Add();
                    co.xAddressSheet = co.xBook.Worksheets.Add();
                    co.xCaoSheet = co.xBook.Worksheets.Add();
                    co.xOtherCodes = co.xBook.Worksheets.Add();
                    co.xStrdSheet = co.xBook.Worksheets.Add();
                    co.xReplSheet = co.xBook.Worksheets.Add();
                    co.xFormerSheet = co.xBook.Worksheets.Add();
                    co.xCageSheet.Name = "CAGE_DATA";
                    co.xAddressSheet.Name = "ADDRESS";
                    co.xCaoSheet.Name = "CAO_ADP_POINT_CODES";
                    co.xOtherCodes.Name = "OTHER_CODES";
                    co.xStrdSheet.Name = "STRD_IND_CLASS";
                    co.xReplSheet.Name = "REPLACEMENT";
                    co.xFormerSheet.Name = "FORMER_DATA";

                    //Row Counters
                    co.cageRow = 1;
                    co.addressRow = 1;
                    co.caoRow = 1;
                    co.otherCodesRow = 1;
                    co.strdRow = 1;
                    co.replRow = 1;
                    co.formerRow = 1;

                    //Models
                    co.cageM = new CAGECDS.CAGE_DATA();
                    co.addressM = new CAGECDS.ADDRESS();
                    co.caoM = new CAGECDS.CAO_ADP_POINT_CODES();
                    co.otherM = new CAGECDS.OTHER_CODES();
                    co.strdM = new CAGECDS.STANDARD_INDUSTRIAL_CLASSIFICATION();
                    co.replM = new CAGECDS.REPLACEMENT();
                    co.formerM = new CAGECDS.FORMER_DATA();

                    using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
                    using (var reader = new StreamReader(stream))
                    {
                        UpdateSheetNames(bookCounter);

                        while (bookCounter <= numBooks && SheetCounter < 1000000 && !reader.EndOfStream)
                        {
                            if (SheetCounter == 0)
                            {
                                co.cageM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.cageM.ADRS_NM_LN_NO_1087 = "ADRS_NM_LN_NO_1087";
                                co.cageM.ADRS_NM_C_TXT_1086 = "ADRS_NM_C_TXT_1086";
                                co.xCageSheet.Cells[co.cageRow, Column] = "\t" + co.cageM.CAGE_CD_9250;
                                co.xCageSheet.Cells[co.cageRow, ++Column] = "\t" + co.cageM.ADRS_NM_LN_NO_1087;
                                co.xCageSheet.Cells[co.cageRow, ++Column] = "\t" + co.cageM.ADRS_NM_C_TXT_1086;
                                Column = 1;
                                co.cageRow++;

                                co.addressM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.addressM.ST_ADRS_1_1082 = "ST_ADRS_1_1082";
                                co.addressM.ST_ADRS_2_1083 = "ST_ADRS_2_1083";
                                co.addressM.POBOX_1361 = "POBOX_1361";
                                co.addressM.CITY_1084 = "CITY_1084";
                                co.addressM.ST_US_POSN_AB_0186 = "ST_US_POSN_AB_0186";
                                co.addressM.ZIP_CD_4400 = "ZIP_CD_4400";
                                co.addressM.CNTRY_1085 = "CNTRY_1085";
                                co.addressM.TEL_NBR_1356 = "TEL_NBR_1356";
                                co.xAddressSheet.Cells[co.addressRow, Column] = "\t" + co.addressM.CAGE_CD_9250;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.ST_ADRS_1_1082;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.ST_ADRS_2_1083;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.POBOX_1361;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.CITY_1084;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.ST_US_POSN_AB_0186;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.ZIP_CD_4400;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.CNTRY_1085;
                                co.xAddressSheet.Cells[co.addressRow, ++Column] = "\t" + co.addressM.TEL_NBR_1356;
                                Column = 1;
                                co.addressRow++;

                                co.caoM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.caoM.CAO_CD_8870 = "CAO_CD_8870";
                                co.caoM.ADP_PNT_CD_8835 = "ADP_PNT_CD_8835";
                                co.xCaoSheet.Cells[co.caoRow, Column] = "\t" + co.caoM.CAGE_CD_9250;
                                co.xCaoSheet.Cells[co.caoRow, ++Column] = "\t" + co.caoM.CAO_CD_8870;
                                co.xCaoSheet.Cells[co.caoRow, ++Column] = "\t" + co.caoM.ADP_PNT_CD_8835;
                                Column = 1;
                                co.caoRow++;

                                co.otherM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.otherM.CAGE_STAT_CD_2694 = "CAGE_STAT_CD_2694";
                                co.otherM.ASSOC_CD_CAGE_8855 = "ASSOC_CD_CAGE_8855";
                                co.otherM.CAGE_TYP_CD_4238 = "CAGE_TYP_CD_4238";
                                co.otherM.TYPE_CAGE_AFFL_0250 = "TYPE_CAGE_AFFL_0250";
                                co.otherM.SZ_OF_BUS_CD_1364 = "SZ_OF_BUS_CD_1364";
                                co.otherM.PR_BUS_CAT_CD_1365 = "PR_BUS_CAT_CD_1365";
                                co.otherM.TYPE_BUS_CD_1366 = "TYPE_BUS_CD_1366";
                                co.otherM.WMN_OWND_BUS_1367 = "WMN_OWND_BUS_1367";
                                co.xOtherCodes.Cells[co.otherCodesRow, Column] = "\t" + co.otherM.CAGE_CD_9250;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.CAGE_STAT_CD_2694;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.ASSOC_CD_CAGE_8855;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.CAGE_TYP_CD_4238;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.TYPE_CAGE_AFFL_0250;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.SZ_OF_BUS_CD_1364;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.PR_BUS_CAT_CD_1365;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.TYPE_BUS_CD_1366;
                                co.xOtherCodes.Cells[co.otherCodesRow, ++Column] = "\t" + co.otherM.WMN_OWND_BUS_1367;
                                Column = 1;
                                co.otherCodesRow++;

                                co.strdM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.strdM.STD_IND_CL_CD_1368 = "STD_IND_CL_CD_1368";
                                co.xStrdSheet.Cells[co.strdRow, Column] = "\t" + co.strdM.CAGE_CD_9250;
                                co.xStrdSheet.Cells[co.strdRow, ++Column] = "\t" + co.strdM.STD_IND_CL_CD_1368;
                                Column = 1;
                                co.strdRow++;

                                co.replM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.replM.RPLM_CAGE_CD_3595 = "RPLM_CAGE_CD_3595";
                                co.xReplSheet.Cells[co.replRow, Column] = "\t" + co.replM.CAGE_CD_9250;
                                co.xReplSheet.Cells[co.replRow, ++Column] = "\t" + co.replM.RPLM_CAGE_CD_3595;
                                Column = 1;
                                co.replRow++;

                                co.formerM.CAGE_CD_9250 = "CAGE_CD_9250";
                                co.formerM.ADRS_NM_LN_NO_1087 = "ADRS_NM_C_NO_1087";
                                co.formerM.ADRS_NM_C_TXT_1086 = "ADRS_NM_C_TXT_1086";
                                co.xFormerSheet.Cells[co.formerRow, Column] = "\t" + co.formerM.CAGE_CD_9250;
                                co.xFormerSheet.Cells[co.formerRow, ++Column] = "\t" + co.formerM.ADRS_NM_LN_NO_1087;
                                co.xFormerSheet.Cells[co.formerRow, ++Column] = "\t" + co.formerM.ADRS_NM_C_TXT_1086;
                                Column = 1;
                                co.formerRow++;
                            }
                            else
                            {
                                //Used to 'jump' to row
                                fastForwardCounter = bookCounter > 0 ? bookCounter * 1000000 : 0;

                                while (FastForward)
                                {
                                    //Jump to row first
                                    for (int i = 0; i < fastForwardCounter; i++)
                                    {
                                        reader.ReadLine();
                                    }

                                    FastForward = false;
                                }

                                nextline = reader.ReadLine();
                                type = nextline.Substring(0, 1);

                                switch (type)
                                {
                                    case "1":
                                        int cageMLength;
                                        cageMLength = nextline.Length;
                                        try
                                        {
                                            co.cageM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.cageM.ADRS_NM_LN_NO_1087 =
                                                cageMLength > 6 ? nextline.Substring(6, 2) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.cageM.ADRS_NM_C_TXT_1086 = cageMLength > 8 ? nextline.Substring(8) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xCageSheet.Cells[co.cageRow, Column] = "\t" + co.cageM.CAGE_CD_9250.Trim();
                                        co.xCageSheet.Cells[co.cageRow, ++Column] =
                                            "\t" + co.cageM.ADRS_NM_LN_NO_1087.Trim();
                                        co.xCageSheet.Cells[co.cageRow, ++Column] =
                                            "\t" + co.cageM.ADRS_NM_C_TXT_1086.Trim();
                                        Column = 1;
                                        co.cageRow++;
                                        break;
                                    case "2":
                                        int addressLength;
                                        addressLength = nextline.Length;
                                        try
                                        {
                                            co.addressM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.ST_ADRS_1_1082 =
                                                addressLength > 6 ? nextline.Substring(6, 36) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.ST_ADRS_2_1083 =
                                                addressLength > 42 ? nextline.Substring(42, 36) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.POBOX_1361 =
                                                addressLength > 78 ? nextline.Substring(78, 36) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.CITY_1084 =
                                                addressLength > 114 ? nextline.Substring(114, 36) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.ST_US_POSN_AB_0186 =
                                                addressLength > 150 ? nextline.Substring(150, 2) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.ZIP_CD_4400 =
                                                addressLength > 152 ? nextline.Substring(152, 10) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.CNTRY_1085 =
                                                addressLength > 161 ? nextline.Substring(161, 36) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.addressM.TEL_NBR_1356 =
                                                addressLength > 198 ? nextline.Substring(198) : "";
                                        }
                                        catch
                                        {
                                        }

                                        //Check if 1st character is number...if so, remove
                                        var isNumber = int.TryParse(co.addressM.CNTRY_1085.Substring(0, 1),
                                            out var result);
                                        co.addressM.CNTRY_1085 = isNumber
                                            ? co.addressM.CNTRY_1085.Remove(0, 1)
                                            : co.addressM.CNTRY_1085;

                                        co.xAddressSheet.Cells[co.addressRow, Column] =
                                            "\t" + co.addressM.CAGE_CD_9250.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.ST_ADRS_1_1082.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.ST_ADRS_2_1083.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.POBOX_1361.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.CITY_1084.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.ST_US_POSN_AB_0186.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.ZIP_CD_4400.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.CNTRY_1085.Trim();
                                        co.xAddressSheet.Cells[co.addressRow, ++Column] =
                                            "\t" + co.addressM.TEL_NBR_1356.Trim();
                                        Column = 1;
                                        co.addressRow++;
                                        break;
                                    case "3":
                                        int caoMLength;
                                        caoMLength = nextline.Length;
                                        try
                                        {
                                            co.caoM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.caoM.CAO_CD_8870 = caoMLength > 6 ? nextline.Substring(6, 6) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.caoM.ADP_PNT_CD_8835 = caoMLength > 12 ? nextline.Substring(12) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xCaoSheet.Cells[co.caoRow, Column] = "\t" + co.caoM.CAGE_CD_9250.Trim();
                                        co.xCaoSheet.Cells[co.caoRow, ++Column] = "\t" + co.caoM.CAO_CD_8870.Trim();
                                        co.xCaoSheet.Cells[co.caoRow, ++Column] = "\t" + co.caoM.ADP_PNT_CD_8835.Trim();
                                        Column = 1;
                                        co.caoRow++;
                                        break;
                                    case "4":
                                        int otherLength;
                                        otherLength = nextline.Length;
                                        try
                                        {
                                            co.otherM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.CAGE_STAT_CD_2694 =
                                                otherLength > 6 ? nextline.Substring(6, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.ASSOC_CD_CAGE_8855 =
                                                otherLength > 7 ? nextline.Substring(7, 5) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.CAGE_TYP_CD_4238 =
                                                otherLength > 12 ? nextline.Substring(12, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.TYPE_CAGE_AFFL_0250 =
                                                otherLength > 13 ? nextline.Substring(13, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.SZ_OF_BUS_CD_1364 =
                                                otherLength > 14 ? nextline.Substring(14, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.PR_BUS_CAT_CD_1365 =
                                                otherLength > 15 ? nextline.Substring(15, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.TYPE_BUS_CD_1366 =
                                                otherLength > 16 ? nextline.Substring(16, 1) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.otherM.WMN_OWND_BUS_1367 =
                                                otherLength > 17 ? nextline.Substring(17) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xOtherCodes.Cells[co.otherCodesRow, Column] =
                                            "\t" + co.otherM.CAGE_CD_9250.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.CAGE_STAT_CD_2694.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.ASSOC_CD_CAGE_8855.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.CAGE_TYP_CD_4238.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.TYPE_CAGE_AFFL_0250.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.SZ_OF_BUS_CD_1364.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.PR_BUS_CAT_CD_1365.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.TYPE_BUS_CD_1366.Trim();
                                        co.xOtherCodes.Cells[co.otherCodesRow, ++Column] =
                                            "\t" + co.otherM.WMN_OWND_BUS_1367.Trim();
                                        Column = 1;
                                        co.otherCodesRow++;
                                        break;
                                    case "5":
                                        int strdLength;
                                        strdLength = nextline.Length;
                                        try
                                        {
                                            co.strdM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.strdM.STD_IND_CL_CD_1368 = strdLength > 6 ? nextline.Substring(6) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xStrdSheet.Cells[co.strdRow, Column] = "\t" + co.strdM.CAGE_CD_9250.Trim();
                                        co.xStrdSheet.Cells[co.strdRow, ++Column] =
                                            "\t" + co.strdM.STD_IND_CL_CD_1368.Trim();
                                        Column = 1;
                                        co.strdRow++;
                                        break;
                                    case "6":
                                        int replLength;
                                        replLength = nextline.Length;
                                        try
                                        {
                                            co.replM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.replM.RPLM_CAGE_CD_3595 = replLength > 6 ? nextline.Substring(6) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xReplSheet.Cells[co.replRow, Column] = "\t" + co.replM.CAGE_CD_9250.Trim();
                                        co.xReplSheet.Cells[co.replRow, ++Column] =
                                            "\t" + co.replM.RPLM_CAGE_CD_3595.Trim();
                                        Column = 1;
                                        co.replRow++;
                                        break;
                                    case "7":
                                        int formerLength;
                                        formerLength = nextline.Length;
                                        try
                                        {
                                            co.formerM.CAGE_CD_9250 = nextline.Substring(1, 5);
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.formerM.ADRS_NM_LN_NO_1087 =
                                                formerLength > 6 ? nextline.Substring(6, 2) : "";
                                        }
                                        catch
                                        {
                                        }

                                        try
                                        {
                                            co.formerM.ADRS_NM_C_TXT_1086 =
                                                formerLength > 8 ? nextline.Substring(8) : "";
                                        }
                                        catch
                                        {
                                        }

                                        co.xFormerSheet.Cells[co.formerRow, Column] =
                                            "\t" + co.formerM.CAGE_CD_9250.Trim();
                                        co.xFormerSheet.Cells[co.formerRow, ++Column] =
                                            "\t" + co.formerM.ADRS_NM_LN_NO_1087.Trim();
                                        co.xFormerSheet.Cells[co.formerRow, ++Column] =
                                            "\t" + co.formerM.ADRS_NM_C_TXT_1086.Trim();
                                        break;
                                    default:
                                        break;
                                }
                            }

                            pb.PerformStep();
                            lblProgress.Text = "Row: " + (totalCounter - 1) + " of " + TotalRows;
                            SheetCounter++;
                            totalCounter++;
                        }

                        //Reset sheet counter...should never go over 1,000,000 rows
                        SheetCounter = 0;
                        //Re-initiate Fast Forward counter
                        FastForward = true;


                        try
                        {
                            CompleteFilePath = FolderDestination + @"\" + SheetName + "_" +
                                               DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MMM") +
                                               ".xlsx";
                            bookCounter++;
                            SendToTimerFile();
                            co.xBook.Close(true, CompleteFilePath, Missing.Value); //close and save individual book
                            co.App.Quit();
                            Marshal.ReleaseComObject(co.xCageSheet);
                            Marshal.ReleaseComObject(co.xAddressSheet);
                            Marshal.ReleaseComObject(co.xCaoSheet);
                            Marshal.ReleaseComObject(co.xOtherCodes);
                            Marshal.ReleaseComObject(co.xStrdSheet);
                            Marshal.ReleaseComObject(co.xReplSheet);
                            Marshal.ReleaseComObject(co.xFormerSheet);
                            Marshal.ReleaseComObject(co.xBook);
                            Marshal.ReleaseComObject(co.App);
                            //MessageBox.Show("Book " + bookCounter + " finished.");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation);
                            System.Windows.Forms.Application.Exit();
                        }
                    }
                }

                MessageBox.Show("Conversion/Save Succesful", "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Windows.Forms.Application.Exit();
            }

        }

        #endregion

        #region Append Timer File

        private void SendToTimerFile()
        {
            TimeStop = DateTime.Now;
            lblStop.Text = "Stop: " + TimeStop.ToString("h:mm:ss tt");
            TimeSpan = TimeStop - TimeStart;
            lblElapsed.Text = "Elapsed: " + TimeSpan.ToString("g");

            //Timer File
            string path = @"D:\Tony\" + DateTime.Now.ToString("yyyy") + @"\" + DateTime.Now.ToString("MMM") +
                          @"\Excel Conversions\Timer.txt";

            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine(SheetName + " - " + DateTime.Now.ToString("d") + " - " + lblElapsed.Text);
                }
            }
            else
            {
                // This text is always added, making the file longer over time
                // if it is not deleted.
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(SheetName + " - " + DateTime.Now.ToString("d") + " - " + lblElapsed.Text);
                }
            }
        }

        #endregion

        #region AMMO

        private void AMMO()
        {
            NewXlApp();
            var am = new AMMO();
            using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream))
            {
                while (!reader.EndOfStream)
                {
                    if (!reader.EndOfStream)
                    {
                        if (Row == 1)
                        {
                            am.ITEM_NAME = "ITEM NAME";
                            am.DODAC = "DODAC";
                            am.DESCRIPTION = "DESCRIPTION";
                            am.STATUS = "STATUS";
                            am.FSC = "FSC";
                            am.DODIC = "DODIC";
                        }
                        else
                        {
                            var nextLine = reader.ReadLine();
                            am.NAME_LENGTH = nextLine.Substring(0, 3);
                            var varOne = int.Parse(am.NAME_LENGTH);
                            var pipeIndex = nextLine.IndexOf("|", 0);

                            am.ITEM_NAME = nextLine.Substring(3, varOne);
                            am.DODAC = nextLine.Substring(varOne + 3, 9);

                            var index1 = varOne + 12;
                            var length1 = pipeIndex - index1;
                            am.DESCRIPTION = nextLine.Substring(index1, length1);

                            var index2 = index1 + 1 + length1;
                            am.STATUS = nextLine.Substring(index2, 1);

                            var index3 = index2 + 1;
                            am.FSC = nextLine.Substring(index3, 4);

                            var index4 = index3 + 4;
                            am.DODIC = nextLine.Substring(index4, 4);

                        }

                        Worksheet.Cells[Row, Column] = "\t" + am.ITEM_NAME.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + am.DODAC.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + am.DESCRIPTION.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + am.STATUS.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + am.FSC.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + am.DODIC.Trim();

                        pb.PerformStep();
                        Row++;
                        Column = 1;
                    }
                }
            }

            CloseExcelApplication();
            SuccessMessage();
        }

        #endregion

        #region ENAC

        private void ENAC()
        {
            //Excel Object
            NewXlApp();

            var enacModel = new ENAC();

            using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream))

            {
                while (!reader.EndOfStream)
                {
                    //var line = reader.ReadLine(); USE TO FIND DELIMITER SUCH AS PIPES

                    if (!reader.EndOfStream)
                    {

                        if (Row % 100 == 0 && Row != 0) Console.WriteLine("Converting Row " + Row + " of " + TotalRows);

                        if (Row == 1)
                        {
                            enacModel.FSC_FSG = "FSC_FSG";
                            enacModel.NIIN = "NIIN";
                            enacModel.ENAC_3025 = "ENAC_3025";
                            enacModel.NAME = "NAME";
                            enacModel.DT_NIIN_ASGMT_2180 = "DT_NIIN_ASGMT_2180";
                            enacModel.EFF_DT_2128 = "EFF_DT_2128";
                            enacModel.INC_4080 = "INC_4080";
                            enacModel.SOS_CD_3690_or_SOSM_CD_2948 = "SOS_CD_3690_or_SOSM_CD_2948";
                        }
                        else
                        {
                            var nextLine = reader.ReadLine();
                            enacModel.FSC_FSG = nextLine.Substring(0, 4);
                            enacModel.NIIN = nextLine.Substring(4, 9);
                            enacModel.ENAC_3025 = nextLine.Substring(13, 2);
                            enacModel.NAME = nextLine.Substring(15, 32);
                            enacModel.DT_NIIN_ASGMT_2180 = nextLine.Substring(47, 7);
                            enacModel.EFF_DT_2128 = nextLine.Substring(54, 7);
                            enacModel.INC_4080 = nextLine.Substring(61, 5);
                            enacModel.SOS_CD_3690_or_SOSM_CD_2948 = nextLine.Substring(66, 3);
                        }

                        Worksheet.Cells[Row, Column] = "\t" + enacModel.FSC_FSG.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.NIIN.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.ENAC_3025.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.NAME.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.DT_NIIN_ASGMT_2180.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.EFF_DT_2128.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.INC_4080.Trim();
                        Worksheet.Cells[Row, ++Column] = "\t" + enacModel.SOS_CD_3690_or_SOSM_CD_2948.Trim();

                        pb.PerformStep();
                        Row++;
                        Column = 1;
                    }
                }
            }

            CloseExcelApplication();
            SuccessMessage();
        }

        #endregion

        #region Success Message

        private void SuccessMessage()
        {
            Cursor = Cursors.Default;
            this.BackColor = Color.Navy;
            MessageBox.Show("Conversion/Save Succesful", "DLA to Excel", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        #endregion

        #region Close Excel Application

        private void CloseExcelApplication()
        {
            try
            {
                SendToTimerFile();
                Workbook.Close(true, CompleteFilePath, Missing.Value); //close and save individual book
                XlApp.Quit();
                Marshal.ReleaseComObject(Worksheet);
                Marshal.ReleaseComObject(Workbook);
                Marshal.ReleaseComObject(XlApp);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error.\n\r" + e, "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.Windows.Forms.Application.Exit();
            }
        }

        #endregion

        #region New Excel Application

        private void NewXlApp()
        {
            XlApp = new Application();
            Workbook = XlApp.Workbooks.Add();
            Worksheet = Workbook.Worksheets.Add();
            Worksheet.Name = SheetName;
        }

        #endregion

        #region Activate Progress Bar

        private void ProgressBarActive()
        {
            var counter = 0;
            if (!string.IsNullOrEmpty(opfd.FileName))
            {
                counter += File.ReadLines(opfd.FileName).Count();
            }

            TotalRows = counter;
            pb.Maximum = TotalRows;
            pb.Minimum = 1;
            pb.Value = 1;
            pb.Step = 1;
        }

        #endregion

        #region Check For Empty Fields

        private bool CheckFields(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cmbTableName.Text))
            {
                MessageBox.Show("Choose Table", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbTableName.Focus();
                return false;
            }

            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                MessageBox.Show("Choose File Path", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BtnFileLocation_Click(sender, e);
                return false;
            }

            if (string.IsNullOrEmpty(txtFolderDestination.Text))
            {
                MessageBox.Show("Choose Folder Destination", "DLA to Excel", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                BtnFolderDestination_Click(sender, e);
                return false;
            }

            return true;
        }

        #endregion

        #region Exit Application

        private void BtnExit_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Exit Application?", "DLA to Excel", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (result == DialogResult.Yes) System.Windows.Forms.Application.Exit();
        }

        #endregion

        #region Row Counter Button

        private void BtnRowCount_Click(object sender, EventArgs e)
        {
            var counter = 0.0;
            if (!CheckForEmptyFields(sender, e)) return;
            counter += File.ReadLines(opfd.FileName).Count();
            var excelBooks = Math.Ceiling(counter / 250000);
            var text = "Rows: " + counter.ToString("N0") + " | Books: " + excelBooks.ToString("N0");
            lblRowCount.Text = text;
        }

        #endregion

        #region Sample Button

        private void BtnSample_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(opfd.FileName))
            {
                MessageBox.Show("Need File Path", "DLA to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BtnFileLocation_Click(sender, e);
                return;
            }

            var sample = "";
            using (var stream = new FileStream(FileLocation, FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream))
            {
                for (int i = 0; i < 20; i++)
                {
                    sample += reader.ReadLine() + "\r\n";
                }
            }

            MessageBox.Show(sample, "First 20 Rows", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        #endregion
    }
}
