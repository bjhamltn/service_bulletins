using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Xml;
using System.Security.AccessControl;
using HtmlAgilityPack;
using System.Web;
using MyXPathReader;
using System.Net;
using System.IO.Compression;
using System.Xml.Xsl;


namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {

        public string airline_operator = "FDX";
        bool exitOnComplet = false;
        public string activeSb = "";
        public string activeSb_dir = "";
        public string SBNum = "";
        bool isAirbus = false;
        public HttpWebRequest request = null;
        public CookieContainer CookieContainerSession = new CookieContainer();
        public XmlDocument adaptCache = new XmlDocument();
        public NameValueCollection GroupReasons = new NameValueCollection();
        public NameValueCollection CPN_MPNCollection = new NameValueCollection();
        public List<object> MaterialInfo = new List<object>();
        public List<string> ConfigurationGroups = new List<string>();
        NameValueCollection manHours = new NameValueCollection();        
        public NameValueCollection FigureTables = new NameValueCollection();
        public NameValueCollection FigureInstrunctions = new NameValueCollection();
        public NameValueCollection FigureGraphics = new NameValueCollection();
        public NameValueCollection InstrunctionsFigures = new NameValueCollection();
        public NameValueCollection instrunctionCollection = new NameValueCollection();
        public NameValueCollection instrCollection = new NameValueCollection();
        public NameValueCollection instrunctionTables = new NameValueCollection();
        NameValueCollection groupTables_Material_Media = new NameValueCollection();
        NameValueCollection groupTables_Material_KITS = new NameValueCollection();
        NameValueCollection groupTables_Material_WIRES = new NameValueCollection();
        public Hashtable doctypeEntity = new Hashtable();
        public XmlReaderSettings readerSetting = new XmlReaderSettings();
        public List<string> EmptyTags = new List<string>();
        List<string> FedExGroups = new List<string>();
        public string SB_SUBJECT = "";
        public string ATA_System = "";

        public DialogResult dialogResult_adapt_bw = new DialogResult();

        public bool pogressReportingCompete = false;

        #region MyRegion

       
        public class FileSorterDateModifies : IComparer
        {

            int IComparer.Compare(Object x, Object y)
            {
                long a = File.GetCreationTime(x.ToString()).Ticks;
                long b = File.GetCreationTime(y.ToString()).Ticks;

                if (a > b)
                {
                    return 1;
                }
                else
                {
                    return -1;
                }
            }
        }

        public class InstrunctionSorter2 : IComparer
        {
            int IComparer.Compare(Object x, Object y)
            {
                string a = x.ToString();
                string b = y.ToString();
                string[] aParmas = a.Split('.');
                string[] bParmas = b.Split('.');

                aParmas[0] = aParmas[0].PadLeft(6, '0');
                bParmas[0] = bParmas[0].PadLeft(6, '0');

                a = "";
                b = "";
                foreach (string kk in aParmas)
                {
                    a = a + kk + "_";
                }

                foreach (string kk in bParmas)
                {
                    b = b + kk + "_";
                }


                int sdsdf = (new CaseInsensitiveComparer()).Compare(a, b);

                return ((new CaseInsensitiveComparer()).Compare(a, b));
            }
        }

        public class InstrunctionSorter : IComparer
        {

            int IComparer.Compare(Object x, Object y)
            {
                string a = x.ToString();
                string b = y.ToString();

                a = a.Substring(0, a.IndexOf(".")).PadLeft(6, '0');
                b = b.Substring(0, b.IndexOf(".")).PadLeft(6, '0');

                int sdsdf = (new CaseInsensitiveComparer()).Compare(a, b);

                return ((new CaseInsensitiveComparer()).Compare(a, b));
            }

        }

        public class FigureSorter : IComparer
        {

            int IComparer.Compare(Object x, Object y)
            {
                string a = x.ToString().ToUpper().Replace("FIGURE", "").Trim();
                string b = y.ToString().ToUpper().Replace("FIGURE", "").Trim();

                a = a.ToString().PadLeft(6, '0');

                b = b.ToString().PadLeft(6, '0');

                int sdsdf = (new CaseInsensitiveComparer()).Compare(a, b);

                return ((new CaseInsensitiveComparer()).Compare(a, b));
            }

        }

        public Form1(string fileName)
        {
            exitOnComplet = true;

            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            string user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;



            openFileDialog1.FileName = fileName;

            BackgroundWorker serviceBulliten = new BackgroundWorker();
            serviceBulliten.WorkerReportsProgress = true;
            serviceBulliten.ProgressChanged += serviceBulliten_ProgressChanged;
            serviceBulliten.RunWorkerCompleted += serviceBulliten_RunWorkerCompleted;
            serviceBulliten.DoWork += serviceBullitenXML_DoWork;
            serviceBulliten.RunWorkerAsync(DialogResult.OK);

        }

        public Form1()
        {

            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            string user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            HtmlAgilityPack.HtmlNode.ElementsFlags.Add("revst", HtmlElementFlag.Empty);
            HtmlAgilityPack.HtmlNode.ElementsFlags.Add("revend", HtmlElementFlag.Empty);
            HtmlAgilityPack.HtmlNode.ElementsFlags.Add("spanspec", HtmlElementFlag.Empty);
            HtmlAgilityPack.HtmlNode.ElementsFlags.Add("colspec", HtmlElementFlag.Empty);
            readerSetting = new XmlReaderSettings();
            readerSetting.DtdProcessing = DtdProcessing.Parse;
            EmptyTags = HtmlAgilityPack.HtmlNode.ElementsFlags.Where(fd => ((HtmlElementFlag)fd.Value) == HtmlElementFlag.Empty).ToList().Select(fd=> fd.Key).ToList();
        }

        private void Sort_VariableNames_Click(object sender, EventArgs e) { }

        private List<Object> expandEffectivity(string effectivity)
        {
            string[] scope_effectivity = effectivity.Trim().Split(new string[] { "-", "–" }, StringSplitOptions.RemoveEmptyEntries);
            int lower = -1;
            int upper = -1;

            List<string> expandedEffectivity = new List<string>();
            List<Object> expandedEffectivityINTs = new List<Object>();
            if (scope_effectivity.Length == 2)
            {
                if (int.TryParse(scope_effectivity[0], out lower))
                {
                    if (int.TryParse(scope_effectivity[1], out upper))
                    {
                        int i = 0;
                        for (i = lower; i <= upper; i++)
                        {
                            expandedEffectivity.Add(i.ToString());
                            expandedEffectivityINTs.Add(i);
                        }
                    }
                }
            }
            else if (scope_effectivity.Length == 1)
            {
                if (int.TryParse(scope_effectivity[0], out lower))
                {
                    expandedEffectivityINTs.Add(lower);
                }
            }
            return expandedEffectivityINTs;
        }

        private List<Object> expandEffectivity_Array(string Effectivity)
        {
            string[] scope_effectivitySplit = Effectivity.Trim().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();

            List<string> expandedEffectivity = new List<string>();
            List<Object> expandedEffectivityINTs = new List<Object>();

            foreach (string effectivity in scope_effectivitySplit)
            {

                string[] scope_effectivity = effectivity.Trim().Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                int lower = -1;
                int upper = -1;


                if (scope_effectivity.Length == 2)
                {
                    if (int.TryParse(scope_effectivity[0], out lower))
                    {
                        if (int.TryParse(scope_effectivity[1], out upper))
                        {
                            int i = 0;
                            for (i = lower; i <= upper; i++)
                            {
                                expandedEffectivityINTs.Add(i);
                            }

                        }
                    }
                }
                else
                {
                    if (int.TryParse(effectivity.Trim(), out lower))
                    {
                        expandedEffectivityINTs.Add(lower);
                    }
                }
            }

            expandedEffectivityINTs.Sort();
            return expandedEffectivityINTs.Distinct().ToList();
        }

        private List<string> expandEffectivity_StringList(string Effectivity)
        {
            string[] scope_effectivitySplit = Effectivity.Trim().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();

            List<string> expandedEffectivity = new List<string>();
            

            foreach (string effectivity in scope_effectivitySplit)
            {

                string[] scope_effectivity = effectivity.Trim().Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                int lower = -1;
                int upper = -1;


                if (scope_effectivity.Length == 2)
                {
                    if (int.TryParse(scope_effectivity[0], out lower))
                    {
                        if (int.TryParse(scope_effectivity[1], out upper))
                        {
                            int i = 0;
                            for (i = lower; i <= upper; i++)
                            {
                                expandedEffectivity.Add(i.ToString());
                            }

                        }
                    }
                }
                else
                {
                    if (int.TryParse(effectivity.Trim(), out lower))
                    {
                        expandedEffectivity.Add(lower.ToString());
                    }
                }
            }

            expandedEffectivity.Sort();
            return expandedEffectivity.Distinct().ToList();
        }

        private string expandEffectivity_str(string Effectivity, string configCode)
        {

            string[] scope_effectivitySplit = Effectivity.Trim().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sb = new StringBuilder();

            List<string> expandedEffectivity = new List<string>();
            List<Object> expandedEffectivityINTs = new List<Object>();

            foreach (string effectivity in scope_effectivitySplit)
            {

                string[] scope_effectivity = effectivity.Trim().Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                int lower = -1;
                int upper = -1;


                if (scope_effectivity.Length == 2)
                {
                    if (int.TryParse(scope_effectivity[0], out lower))
                    {
                        if (int.TryParse(scope_effectivity[1], out upper))
                        {
                            int i = 0;
                            for (i = lower; i <= upper; i++)
                            {
                                expandedEffectivityINTs.Add(i);
                            }

                        }
                    }
                }
                else
                {
                    if (int.TryParse(effectivity.Trim(), out lower))
                    {
                        expandedEffectivityINTs.Add(lower);
                    }
                }
            }

            expandedEffectivityINTs.Sort();
            expandedEffectivityINTs = expandedEffectivityINTs.Distinct().ToList();
            foreach (int tailNbr in expandedEffectivityINTs)
            {
                sb.Append(tailNbr.ToString().Trim().PadLeft(3, '0') + configCode);
                sb.Append(", ");
            }

            return sb.ToString().Trim().Trim(new char[] { ',' });
        }

        public List<Array> connectionHits = new List<Array>();
  
        private void button1_WrkInstuction_Click(object sender, EventArgs e)
        {
            progressBar_total.Value = 0;
            progressBar_total.Maximum = 100;
            progressBar_total.Minimum = 0;
            progressBar1.MarqueeAnimationSpeed = 10;
            progressBar1.Value = 0;
            progressBar1.Maximum = 100;
            progressBar1.Minimum = 0;
            BackgroundWorker serviceBulliten = new BackgroundWorker();
            serviceBulliten.WorkerReportsProgress = true;
            serviceBulliten.ProgressChanged += serviceBulliten_ProgressChanged;
            serviceBulliten.RunWorkerCompleted += serviceBulliten_RunWorkerCompleted;
            serviceBulliten.DoWork += serviceBullitenXML_DoWork;
            serviceBulliten.RunWorkerAsync(openFileDialog1.ShowDialog());
        }

        void serviceBulliten_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (exitOnComplet)
            {
                this.Close();
            }
            progressBar1.Value = 100;            
        }

        void serviceBulliten_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
         
            if (e.UserState.ToString() == "progress")
            {
                progressBar1.Value = Math.Min(100, e.ProgressPercentage);
                return;
            }
            else if (e.UserState.ToString() == "progress_total")
            {
                progressBar_total.Value = Math.Min(100, e.ProgressPercentage);
                return;
            }
            switch (e.ProgressPercentage)
            {
                case 0:
                    if (e.UserState.GetType() == typeof(HtmlAgilityPack.HtmlDocument))
                    {
                        textBox_SBNum.Text = SBNum;
                        HtmlAgilityPack.HtmlDocument sds = ((HtmlAgilityPack.HtmlDocument)e.UserState);
                        ATA_System = sds.DocumentNode.SelectSingleNode("//sb").Attributes["chapsect"].Value.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                        HtmlNode firstREASON = sds.DocumentNode.SelectSingleNode("//plansect[@sectname='RES']/para");
                        HtmlNodeCollection groupReasons = sds.DocumentNode.SelectNodes("//plansect[@sectname='RES']/unlist/unlitem");
                        GroupReasons = new NameValueCollection();

                        if (groupReasons != null)
                        {
                            #region MyRegion
                            foreach (HtmlNode reasonNode in groupReasons)
                            {
                                HtmlAgilityPack.HtmlDocument gg = new HtmlAgilityPack.HtmlDocument();
                                gg.DocumentNode.InnerHtml = "<aa>" + reasonNode.InnerHtml + "</aa>";


                                HtmlNode GroupTitle = gg.DocumentNode.SelectSingleNode("//para");
                                if (GroupTitle != null)
                                {
                                    string groupText = GroupTitle.InnerText;

                                    groupText = Regex.Match(groupText, @"Group.+?:").Value.Replace("Group", "").Replace(":", "");

                                    List<object> groupList = expandEffectivity_Array(groupText);

                                    HtmlNodeCollection unlitems = gg.DocumentNode.SelectNodes("//unlitem");

                                    StringBuilder reasonBuilder = new StringBuilder();

                                    string fristReason = HttpUtility.HtmlDecode(firstREASON.InnerText.Trim()) + "\r\n";
                                    string secondReason = HttpUtility.HtmlDecode(GroupTitle.InnerText.Substring(GroupTitle.InnerText.IndexOf(":") + 1).Trim());

                                    reasonBuilder.AppendLine(secondReason);

                                    foreach (object grp in groupList)
                                    {

                                        if (GroupReasons[grp.ToString()] != null)
                                        {
                                            if (GroupReasons.GetValues(grp.ToString()).Contains(fristReason) == false)
                                            {
                                                GroupReasons.Add(grp.ToString(), fristReason);
                                            }
                                        }
                                        else
                                        {
                                            GroupReasons.Add(grp.ToString(), fristReason);
                                        }
                                        if (unlitems != null)
                                        {
                                            foreach (HtmlNode unlitem in unlitems)
                                            {
                                                reasonBuilder.AppendLine(HttpUtility.HtmlDecode(unlitem.InnerText.Trim()) + "\r\n");
                                            }
                                        }
                                        GroupReasons.Add(grp.ToString(), reasonBuilder.ToString());
                                    }

                                }
                            }
                            #endregion
                        }
                    }
                    
                    break;
                case 1:
                    progressBar1.Value = Math.Min(100, e.ProgressPercentage);
                    textBox_SBNum.Text = SBNum;
                    break;
    
                case 5:

                    break;
         
              
   
                case 14:
                    break;        

            }
            
        }

        private void serviceBullitenXML_DoWork(object sender, DoWorkEventArgs e)
        {

            StringBuilder text = new StringBuilder();
            if ((DialogResult)e.Argument == DialogResult.OK)
            {
                BackgroundWorker bw = (BackgroundWorker)sender;
                backgroundWorker1 = bw;
                string zipFileName_SEL = openFileDialog1.FileName;
                string[] files = Directory.GetFiles(Path.GetDirectoryName(zipFileName_SEL));
                files = files.Where(FD => Path.GetExtension(FD) == ".zip").ToArray();
                string zipFileName = zipFileName_SEL;
                bw.ReportProgress(0, "progress_total");
                string innerZip = "innerZip.zip";
                int cnt = 0;
                //foreach (string zipFileName_X in files)
                {
                    doctypeEntity = new Hashtable();
                    //zipFileName = zipFileName_X;
                    isAirbus = false;
                    string sbFilenName = Path.GetFileNameWithoutExtension(zipFileName);
                    isAirbus = Regex.Match(sbFilenName, @"SB_.+_r\d{1,}").Success;

                    Directory.CreateDirectory(Application.StartupPath + "\\temp\\sb");
                    string sbName = Path.GetFileNameWithoutExtension(zipFileName);
                    activeSb = Application.StartupPath + "\\temp\\sb\\" + sbName;
                    innerZip = Application.StartupPath + "\\temp\\sb\\" + sbName+".zip";
                    bool innerZipCheck = true;
                    string dtd = "";
                    #region Extract Files
                Extract_Files:
                    try
                    {
                        using (ZipArchive zip = ZipFile.Open(zipFileName, ZipArchiveMode.Read))
                        {
                            foreach (ZipArchiveEntry entry in zip.Entries.Where(fd => fd.Name.ToLower().EndsWith("sgm")))
                            {
                                if (entry.Name.ToLower().EndsWith("sgm"))
                                {
                                    innerZip = "innerZip.zip";
                                    entry.ExtractToFile(activeSb, true);
                                    break;
                                }
                            }

                            if (!File.Exists(activeSb) && innerZipCheck)
                            {
                                innerZipCheck = false;
                                foreach (ZipArchiveEntry entry in zip.Entries.Where(fd => fd.Name.EndsWith("zip")))
                                {
                                    entry.ExtractToFile(innerZip, true);
                                    break;
                                }
                            }
                        }
                        if (File.Exists(innerZip))
                        {
                            zipFileName = innerZip;
                            goto Extract_Files;
                        }
                    }
                    catch (Exception invalidData)
                    {
                        Console.WriteLine(invalidData.Message);
                    }

                    
                    #endregion

                    if (File.Exists(activeSb))
                    {
                        
                        #region Prepare and format SGML as valid xml
                        if(File.Exists(activeSb + ".tmp"))
                        {
                            File.Delete(activeSb + ".tmp");
                        }
                        
                        
                        StreamReader strR = new StreamReader(activeSb);
                        int lineIdx = 0;
                        string _line = "";
                        while (_line.ToLower().Contains("<sb") == false) 
                        {
                            _line = strR.ReadLine();
                            lineIdx += _line.Length;
                        }                        
                        strR.Close();
                        lineIdx -= _line.Length;
                        lineIdx++;
                        using (Stream stream = File.Open(activeSb, FileMode.Open))
                        {
                            byte[] buff  = new byte[lineIdx];                            
                            stream.Read(buff, 0, buff.Length);
                            _line = String.Join("",  buff.Select(fd=> ((char)fd).ToString() ).ToArray());
                            while (_line.ToLower().Contains("<sb") == false)
                            {
                                stream.Position = 0;
                                buff = new byte[lineIdx++];
                                stream.Read(buff, 0, buff.Length);
                                _line = String.Join("", buff.Select(fd => ((char)fd).ToString()).ToArray());
                            }
                            stream.Position = 0;
                            lineIdx -= 4;
                            buff = new byte[lineIdx];
                            stream.Read(buff, 0, buff.Length);
                            _line = String.Join("", buff.Select(fd => ((char)fd).ToString()).ToArray());

                            buff = Enumerable.Repeat((byte)' ', _line.Length).ToArray();
                            stream.Position = 0;
                            stream.Write(buff, 0, buff.Length);
                            stream.Close();
                        }
                        
                        File.Copy(activeSb, activeSb + ".xml");
                        strR = new StreamReader(activeSb + ".xml");
                        using (StreamWriter strW = new StreamWriter(activeSb, false))
                        {
                            strW.Write(Resource1.entity_def);
                            while (!strR.EndOfStream)
                            {
                                GC.Collect();

                                List<bool> boos = Enumerable.Repeat(true, 100000).ToList();
                                List<string> lines = boos.Select(fd => !strR.EndOfStream ? strR.ReadLine() : "").ToList();
                                string block = string.Join("\r\n", lines).Trim();
                                block = block.Replace("\r\n>", ">");

                                foreach (string tagname in EmptyTags)
                                {
                                    block = Regex.Replace(block, @"<(?<a>" + tagname + "[^>]*?)>", delegate(Match m)
                                    {
                                        if (m.Groups["a"].Value.Trim().EndsWith("/"))
                                        {
                                            return m.Value;
                                        }
                                        else
                                        {
                                            string sd = string.Format("<{0}/>", m.Groups["a"].Value);
                                            return sd;
                                        }

                                    }, RegexOptions.IgnoreCase);
                                }
                                if (isAirbus)
                                {
                                    block = Regex.Replace(block, @"<(?<a>\S+?)(?<b>[\s>])", delegate(Match m)
                                    {
                                        string sd = "<" + m.Groups["a"].Value.ToLower() + m.Groups["b"].Value;
                                        return sd;
                                    }, RegexOptions.IgnoreCase);

                                    block = Regex.Replace(block, @"</(?<a>\S{1,})>", delegate(Match m)
                                    {
                                        string sd = "</" + m.Groups["a"].Value.ToLower() + ">";
                                        return sd;
                                    }, RegexOptions.IgnoreCase);

                                    block = Regex.Replace(block, @"\s(?<a>\S{1,})=", delegate(Match m)
                                    {
                                        string sd = " " + m.Groups["a"].Value.ToLower() + "=";
                                        return sd;
                                    }, RegexOptions.IgnoreCase);
                                }

                                boos = null;
                                lines = null;
                                strW.WriteLine(block);
                                block = null;
                                GC.Collect();
                            }
                            strR.Close();
                            strW.Close();
                            strW.Dispose();
                            strR.Dispose();
                        }
                        

                        //HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        //htmlDoc.OptionWriteEmptyNodes = true;
                        //htmlDoc.OptionAutoCloseOnEnd = true;
                        //htmlDoc.OptionOutputAsXml = true;                        
                        //htmlDoc.Load(activeSb + ".xml");
                        //htmlDoc.Save(activeSb);
                        //htmlDoc = new HtmlAgilityPack.HtmlDocument();
                        File.Delete(activeSb + ".xml");                        
                        GC.Collect();
                        GC.Collect();
                        GC.Collect();
                        #endregion
                        
                        #region Update Sb Number
                        bw.ReportProgress(10, "progress");
                        XmlReader reader = XmlReader.Create(activeSb, readerSetting);
                        reader.ReadToDescendant("sb");
                        reader.MoveToAttribute("docnbr");
                        SBNum = reader.GetAttribute("docnbr").ToUpper();
                        
                        if (SBNum.Contains("A300") || SBNum.Contains("A310"))
                        {
                            isAirbus = true;
                        }
                        bw.ReportProgress(1, SBNum); 
                        #endregion
                        GC.Collect();
                        GC.Collect();                  
                        #region Get revision, chapsect, fleet, and SB_RevDate
                        string SB_RevDate = reader.GetAttribute("revdate").ToUpper();
                        string SB_Rev = reader.GetAttribute("tsn").ToUpper();
                        string chapsect = reader.GetAttribute("chapsect");
                        string chapnbr = reader.GetAttribute("chapnbr");
                        string seqnbr = reader.GetAttribute("seqnbr");
                        chapsect = chapsect == null ? String.Format("{0}-{1}", chapnbr, seqnbr) : chapsect.ToUpper();
                        string model = reader.GetAttribute("model").ToUpper();
                        reader.Close();
                        String fleet = model.Split(new[] { "-" }, StringSplitOptions.RemoveEmptyEntries)[0];
                        #endregion

                        activeSb_dir = Application.StartupPath + "\\" + fleet + "\\" + SBNum;
                       
                        #region Determine if file is a new revision (refresh == true)
                        bool refresh = false;
                        if (Directory.Exists(activeSb_dir) == false)
                        {
                            Directory.CreateDirectory(activeSb_dir);
                        }
                      
                        if (refresh == false)
                        {
                            if (File.Exists(activeSb_dir + "\\cgms.zip") == false)
                            {
                                refresh = true;
                            }
                            else
                            {
                                XmlDocument revDoc = new XmlDocument();
                                revDoc.Load(activeSb_dir + "\\rev_info.xml");
                                refresh = int.Parse(revDoc.DocumentElement.SelectSingleNode("rev").InnerText) <
                                    int.Parse(SB_Rev);
                            }
                        }
                        #endregion
                        
                        //if (refresh)
                        {
                            #region Parse Data
                            if (Directory.Exists(activeSb_dir) && !refresh)
                            {
                            isFileDeleted:
                                #region MyRegion
                                try
                                {
                                    while (Directory.Exists(activeSb_dir))
                                    {
                                        Directory.Delete(activeSb_dir, true);
                                        System.Threading.Thread.Sleep(100);
                                    }
                                }
                                catch (Exception err)
                                {
                                    System.Threading.Thread.Sleep(100);
                                    goto isFileDeleted;
                                }
                                #endregion
                            }
                            else
                            {
                                getSbMeta();
                                //continue;
                            }

                            #region MyRegion
                            Directory.CreateDirectory(activeSb_dir);
                            using (StreamWriter strW = new StreamWriter(activeSb_dir + "\\doctype.dtd", false))
                            {
                                strW.Write(_line);
                                strW.Close();
                                strW.Dispose();
                            }
                            MatchCollection entityMatches = Regex.Matches(_line, @"<!ENTITY.+");
                            foreach (Match m in entityMatches)
                            {
                                string key = Regex.Match(m.Value, @"ENTITY\s+(?<a>\S{1,})\s").Groups["a"].Value;
                                string value = Regex.Match(m.Value, "\"(?<a>.+)\"").Groups["a"].Value;
                                doctypeEntity.Add(key, value);

                            }


                            using (StreamWriter swRev = new StreamWriter(activeSb_dir + "\\rev_info.xml", false))
                            {
                                swRev.WriteLine("<revdata>");
                                swRev.WriteLine("<revdate>" + SB_RevDate + "</revdate>");
                                swRev.WriteLine("<rev>" + SB_Rev + "</rev>");
                                swRev.WriteLine("<chapsect>" + chapsect + "</chapsect>");
                                swRev.WriteLine("</revdata>");
                                swRev.Close();
                                swRev.Dispose();
                            }

                            Directory.CreateDirectory(activeSb_dir + "\\cgms\\");
                            File.Copy(zipFileName, activeSb_dir + "\\cgms.zip");



                            GC.Collect();
                            getRefIds();
                            HtmlAgilityPack.HtmlDocument sbXML = new HtmlAgilityPack.HtmlDocument();
                            getSbGraphics();
                            GC.Collect();
                            getSbInstrs();
                            getSbMeta();
                            getSbInstrsGen();
                            getReferences();
                            getAppendix();
                            GC.Collect();
                            getPlansects("plansect", "sect");
                            GC.Collect();
                            getPlansects("matsect", "sect");
                            GC.Collect();
                            getPlansects("tssect", "sect");
                            GC.Collect();
                            getPlansects("instsect", "inst");
                            GC.Collect();

                            MaterialInfo = new List<object>();
                            ConfigurationGroups = new List<string>();
                            manHours = new NameValueCollection();
                            FigureTables = new NameValueCollection();
                            FigureInstrunctions = new NameValueCollection();
                            FigureGraphics = new NameValueCollection();
                            InstrunctionsFigures = new NameValueCollection();
                            instrunctionCollection = new NameValueCollection();
                            instrCollection = new NameValueCollection();
                            instrunctionTables = new NameValueCollection();
                            groupTables_Material_Media = new NameValueCollection();
                            groupTables_Material_KITS = new NameValueCollection();
                            groupTables_Material_WIRES = new NameValueCollection();
                            #endregion

                            ProcessBoeingSB_XML(sbXML, bw);
                            #endregion
                            #region zip file up
                            bw.ReportProgress(100, "progress");

                            string compresedFile = activeSb_dir + "\\" + SBNum + ".zip";
                            using (ZipArchive archive = ZipFile.Open(compresedFile, ZipArchiveMode.Create))
                            {
                                string dir2zip = activeSb_dir + "\\Group-WICs";
                                string[] totalFiles = Directory.GetFiles(activeSb_dir, "*.*", SearchOption.AllDirectories);
                                int cnt_del = 0;
                                foreach (string file in totalFiles)
                                {
                                    if (Path.GetFileNameWithoutExtension(file) == "cgms")
                                    {
                                        continue;
                                    }
                                    if (Path.GetFileNameWithoutExtension(file) == SBNum)
                                    {
                                        continue;
                                    }
                                    if (Path.GetFileNameWithoutExtension(file) == "rev_info")
                                    {
                                        continue;
                                    }
                                    else if (Path.GetFileNameWithoutExtension(file) == "meta")
                                    {
                                        continue;
                                    }
                                    archive.CreateEntryFromFile(file, Path.GetFileName(file));
                                    while (File.Exists(file))
                                    {
                                        File.Delete(file);
                                        System.Threading.Thread.Sleep(10);
                                    }
                                    cnt_del++;
                                    bw.ReportProgress(100 * cnt_del / totalFiles.Count(), "progress");
                                }
                            }
                            bw.ReportProgress(100, "progress");
                            string[] dirs = Directory.GetDirectories(activeSb_dir).ToArray();
                            foreach (string dir in dirs)
                            {
                                Directory.Delete(dir);
                            }
                            #endregion
                        }
                    }
                    while (File.Exists(activeSb))
                    {
                        Thread.Sleep(100);
                        File.Delete(activeSb);
                    }

                    while (File.Exists(activeSb + ".zip"))
                    {
                        Thread.Sleep(100);
                        File.Delete(activeSb + ".zip");
                    }
                    GC.Collect();
                    cnt++;
                    bw.ReportProgress(100 * cnt / files.Count(), "progress_total");
                }
            }
        }

        public void getSbMeta()
        {

            String f = activeSb_dir + "\\";
            if (Directory.Exists(f) == false)
            {
                Directory.CreateDirectory(f);
            }

            XmlDocument meta = new XmlDocument();
            meta.LoadXml("<meta/>");
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);

            reader.ReadToDescendant("sb");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("chapnbr")).Value = reader.GetAttribute("chapnbr");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("chapsect")).Value = reader.GetAttribute("chapsect");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("model")).Value = reader.GetAttribute("model");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("oidate")).Value = reader.GetAttribute("oidate");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("revdate")).Value = reader.GetAttribute("revdate");
            meta.DocumentElement.Attributes.Append(meta.CreateAttribute("docnbr")).Value = reader.GetAttribute("docnbr");
            reader.ReadToDescendant("title");
            string title = reader.ReadInnerXml();
            meta.DocumentElement.InnerText = title;
            reader.Close();
            meta.Save(f + "meta.xml");
            
            

        }

        public void getTitle(HtmlAgilityPack.HtmlDocument sbDoc)
        {

            String f = activeSb_dir + "\\title.txt";
            if (Directory.Exists(activeSb_dir) == false)
            {
                Directory.CreateDirectory(activeSb_dir);
            }

            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");

            bool match = reader.ReadToDescendant("title");
            if (match)
            {
                data = reader.ReadInnerXml();
                using (StreamWriter sw = new StreamWriter(string.Format("{0}", f), false))
                {
                    sw.Write(data);
                    sw.Close();                    
                }
                
            }

            reader.Close();
        }

        public void getReferences()
        {
            if (Directory.Exists(activeSb_dir) == false)
            {
                Directory.CreateDirectory(activeSb_dir);
            }

            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant("plansect");
            while ("REF" != reader.GetAttribute("key").ToUpper() || !match)
            {
                match = reader.ReadToNextSibling("plansect");
                if (!match) { break; }
            }
            if (match)
            {
                string key = "references.xml";
                data = reader.ReadOuterXml();
                HtmlAgilityPack.HtmlDocument pred1 = new HtmlAgilityPack.HtmlDocument();
                pred1.LoadHtml(data);

                {
                    HtmlNodeCollection tables = pred1.DocumentNode.SelectNodes("//table");
                    if (tables != null)
                    {
                        tables.ToList().ForEach(delegate(HtmlNode nd)
                        {
                            string tbl_html = convert2HtmlTable(nd, null);
                            HtmlNode tnew = pred1.CreateElement("table");
                            tnew.InnerHtml = tbl_html;
                            nd.ParentNode.ReplaceChild(tnew, nd);
                        });

                    }
                    data = pred1.DocumentNode.OuterHtml;

                    data = Regex.Replace(data, @"list\d{1,200}", "ul");
                    data = Regex.Replace(data, @"l\d{1,100}item", "li");
                    data = Regex.Replace(data, @"para>", "p>");
                    data = Regex.Replace(data, @"title>", "h1>");
                    data = Regex.Replace(data, @"tssect>", "div>");
                    data = Regex.Replace(data, @"refext>", "p>");
                    using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", activeSb_dir, key), false))
                    {
                        sw.Write(data);
                        sw.Close();
                        sw.Dispose();
                    }
                }
            }
            reader.Close();
        }


        public void getDocType()
        {            
             XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Parse;

            XmlReader reader = XmlReader.Create(activeSb, settings);
            reader.Read();
            while (!reader.EOF)
            {
                
                Console.WriteLine(reader.Name);
                reader.Read();
            }
            reader.Close();
        }
       
        
        public void getRefIds()
        {
           
            XmlDocument figIdx = new XmlDocument();
            if (Directory.Exists(activeSb_dir) == false)
            {
                Directory.CreateDirectory(activeSb_dir);
            }
            figIdx.LoadXml("<figures/>");
          
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant("refint");
            while (reader.EOF == false)
            {
                if (reader.Name == "refint")
                {
                    string refid = reader.GetAttribute("refid");
                    string figName = reader.ReadInnerXml();
                    if (figName.ToUpper().StartsWith("FIGURE"))
                    {
                        XmlNode fig = figIdx.CreateElement("fig");
                        fig.Attributes.Append(figIdx.CreateAttribute("refid")).Value = refid;
                        fig.InnerXml = figName;
                        figIdx.DocumentElement.AppendChild(fig);
                    }
                }
                else
                {
                    reader.Read();
                }
            }
            reader.Close();
            reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            match = reader.ReadToDescendant("grphcref");
            while (reader.EOF == false)
            {
                if (reader.Name == "grphcref")
                {
                    string refid = reader.GetAttribute("refid");
                    string sheetnbr = reader.GetAttribute("sheetnbr");
                    string figName = reader.ReadInnerXml();                    
                    {
                        string name = Regex.Match(refid, @"\-.+").Value;
                        XmlNode fig = figIdx.CreateElement("fig");
                        fig.Attributes.Append(figIdx.CreateAttribute("refid")).Value = refid;
                        fig.InnerXml = name + "_" + sheetnbr;
                        if (figIdx.DocumentElement.SelectSingleNode(".//fig[text()='" + fig.InnerXml + "']") == null)
                        {
                            figIdx.DocumentElement.AppendChild(fig);
                        }
                    }
                }
                else
                {
                    reader.Read();
                }
            }
            reader.Close();
            figIdx.Save(activeSb_dir + "\\figure_index.xml");
        }
       
        public void getPlansects(string sectionname, string ext = "sect")
        {
            if (Directory.Exists(activeSb_dir) == false)
            {
                Directory.CreateDirectory(activeSb_dir);
            }

            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant(sectionname);

            while (match)
            {

                XmlReader section = reader.ReadSubtree();
                section.ReadToDescendant("title");
                data = section.ReadInnerXml();
                XmlNode d = (new XmlDocument()).CreateElement("title");

                d.InnerXml = data;
                string key = d.InnerXml.Replace("/", "-").Replace(" ", "_").ToLower() + "." + ext;

                key = Regex.Replace(key, @"^[A-Z]{1,}\._","", RegexOptions.IgnoreCase).Trim();

                using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", activeSb_dir, key), false))
                {
                    data = d.OuterXml;
                    data = Regex.Replace(data, @"list\d{1,200}", "ul");
                    data = Regex.Replace(data, @"l\d{1,100}item", "li");
                    data = Regex.Replace(data, @"unlist", "ul");
                    data = Regex.Replace(data, @"unlitem", "li");
                    data = Regex.Replace(data, @"para>", "p>");
                    data = Regex.Replace(data, @"title>", "h1>");
                    data = Regex.Replace(data, @"tssect>", "div>");
                    sw.Write("<div>");
                    sw.Write(data);

                    while (section.EOF == false)
                    {
                        if (section.NodeType == XmlNodeType.Element)
                        {
                            string name = section.Name;

                            if (name == "table")
                            {
                                data = section.ReadOuterXml();
                                data = Regex.Replace(data, "<rev.+?/>", "");
                                HtmlAgilityPack.HtmlDocument tables = new HtmlAgilityPack.HtmlDocument();
                                tables.LoadHtml(data);
                                string tbl_html = convert2HtmlTable(tables.DocumentNode, null);
                                tables.DocumentNode.InnerHtml = tbl_html;
                                data = tables.DocumentNode.OuterHtml;
                                tables = new HtmlAgilityPack.HtmlDocument();
                                GC.Collect();
                                data = Regex.Replace(data, @"list\d{1,200}", "ul");
                                data = Regex.Replace(data, @"l\d{1,100}item", "li");
                                data = Regex.Replace(data, @"unlist", "ul");
                                data = Regex.Replace(data, @"unlitem", "li");
                                data = Regex.Replace(data, @"para>", "p>");
                                data = Regex.Replace(data, @"title>", "h1>");
                                data = Regex.Replace(data, @"tssect>", "div>");
                                sw.Write(data);
                            }
                            else if (Regex.Match(name, @"list\d").Success || Regex.Match(name, @"\ditem").Success)
                            {
                                section.Read();
                            }

                            else
                            {
                                data = section.ReadOuterXml();
                                data = Regex.Replace(data, "<rev.+?>", "");


                                HtmlAgilityPack.HtmlDocument tag = new HtmlAgilityPack.HtmlDocument();
                                tag.LoadHtml(data);

                                HtmlNodeCollection tableNodes = tag.DocumentNode.SelectNodes("//table[not(@sgml_table)]");
                                if (tableNodes != null)
                                {
                                    foreach (HtmlNode table in tableNodes)
                                    {
                                        string tbl_html = convert2HtmlTable(table, null);
                                        table.ParentNode.ReplaceChild(HtmlNode.CreateNode(tbl_html), table);
                                    }
                                }
                                tableNodes = null;
                                data = tag.DocumentNode.OuterHtml;
                                tag = new HtmlAgilityPack.HtmlDocument();
                                GC.Collect();
                                data = Regex.Replace(data, @"list\d{1,200}", "ul");
                                data = Regex.Replace(data, @"l\d{1,100}item", "li");
                                data = Regex.Replace(data, @"unlist", "ul");
                                data = Regex.Replace(data, @"unlitem", "li");
                                data = Regex.Replace(data, @"para>", "p>");
                                data = Regex.Replace(data, @"title>", "h1>");
                                data = Regex.Replace(data, @"tssect>", "div>");

                                sw.Write(data);
                                data = null;

                            }
                        }
                        else
                        {

                            section.Read();
                        }
                        data = null;

                    }
                    sw.WriteLine("</div>");
                    sw.Close();
                }
                match = reader.ReadToNextSibling(sectionname);
            }
            reader.Close();
        }

        public void updateTableHisto(string binName, List<HtmlNode> tableCollection)
        {

            if (tableCollection != null)
            {
                foreach (HtmlNode tableNode in tableCollection)
                {
                    string tableID = DateTime.Now.Ticks.ToString();
                    if (tableNode.Attributes["id"] != null)
                    {
                        tableID = tableNode.Attributes["id"].Value;
                    }
                    if (instrunctionTables[binName] == null)
                    {
                        instrunctionTables.Add(binName, tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml));
                    }
                    else
                    {
                        if (instrunctionTables.GetValues(binName).Contains(tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml)) == false)
                        {
                            instrunctionTables.Add(binName, tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml));
                        }
                    }
                }
            }
        }

        public void updateTableHisto(string binName, HtmlNodeCollection tableCollection)
        {

            if (tableCollection != null)
            {
                foreach (HtmlNode tableNode in tableCollection)
                {
                    string tableID = DateTime.Now.Ticks.ToString();
                    if (tableNode.Attributes["id"] != null)
                    {
                        tableID = tableNode.Attributes["id"].Value;
                    }
                    if (instrunctionTables[binName] == null)
                    {
                        instrunctionTables.Add(binName, tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml));
                    }
                    else
                    {
                        if (instrunctionTables.GetValues(binName).Contains(tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml)) == false)
                        {
                            instrunctionTables.Add(binName, tableID + "::" + HttpUtility.HtmlDecode(tableNode.OuterHtml));
                        }
                    }
                }
            }
        }

        public void updateGraphics(HtmlNode figureNode, HtmlNode instuctionNode)
        {
            HtmlAttribute refid = figureNode.Attributes["refid"];
            HtmlNodeCollection graphicSheets = null;
            if (refid != null)
            {
                HtmlNode graphicNode = null;

                String f = activeSb_dir + "\\graphics";
                HtmlAgilityPack.HtmlDocument graphicsDoc = new HtmlAgilityPack.HtmlDocument();
                StringBuilder sgmlBuilder = new StringBuilder();
                if (Directory.Exists(f) != false)
                {
                    string graphicFile = f + "\\" + refid.Value + ".xml";
                    if (File.Exists(graphicFile))
                    {
                        HtmlAgilityPack.HtmlDocument ff = new HtmlAgilityPack.HtmlDocument();
                        ff.Load(graphicFile);
                        graphicNode = ff.DocumentNode.SelectSingleNode("//graphic[@key='" + refid.Value + "']");
                    }
                    else
                    {
                        Directory.GetFiles(f).Where(fd => Path.GetExtension(fd) == ".xml").ToList().ForEach(delegate(string fd)
                        {
                            HtmlAgilityPack.HtmlDocument ff = new HtmlAgilityPack.HtmlDocument();
                            ff.Load(fd);
                            sgmlBuilder.AppendLine(ff.DocumentNode.InnerHtml);
                        });
                        graphicsDoc.LoadHtml("<graphics>" + sgmlBuilder.ToString() + "</graphics>");
                        graphicNode = graphicsDoc.DocumentNode.SelectSingleNode("//graphic[@key='" + refid.Value + "']");
                    }

                    if (graphicNode == null)
                    {
                        graphicNode = instuctionNode.OwnerDocument.DocumentNode.SelectSingleNode("//graphic[@key='" + refid.Value + "']");
                    }
                }
                else
                {
                    graphicNode = instuctionNode.OwnerDocument.DocumentNode.SelectSingleNode("//graphic[@key='" + refid.Value + "']");
                }



                if (graphicNode == null) { return; }
                graphicSheets = graphicNode.SelectNodes(".//sheet");
                int tableIdx = 0;
                foreach (HtmlNode graphicSheet in graphicSheets)
                {
                   

                    if (FigureGraphics[figureNode.InnerText] == null)
                    {
                        FigureGraphics.Add(figureNode.InnerText, graphicSheet.Attributes["gnbr"].Value);
                    }
                    else if (FigureGraphics.GetValues(figureNode.InnerText).Contains(graphicSheet.Attributes["gnbr"].Value) == false)
                    {
                        FigureGraphics.Add(figureNode.InnerText, graphicSheet.Attributes["gnbr"].Value);
                    }

                    HtmlNodeCollection sheetFigTables = graphicSheet.SelectNodes(".//table");
                    if (sheetFigTables != null)
                    {
                        foreach (HtmlNode FigureTable in sheetFigTables)
                        {
                            string jj = "";
                            if (FigureTable.Attributes["id"] == null)
                            {
                                jj = refid.Value + (tableIdx.ToString().PadLeft(3, '0'));
                                tableIdx++;
                            }
                            else
                            {
                                jj = FigureTable.Attributes["id"].Value;
                            }
                            FigureTables.Add(figureNode.InnerText + "_" + graphicSheet.Attributes["gnbr"].Value, jj);
                        }
                    }
                }

            }
        }
 
        public void subStepsGroups(HtmlNode instuction, string groupMain, string parentNumber)
        {
            string subStep = Regex.Match(instuction.Name, @"\d{1,666}").Value;
            subStep = (int.Parse(subStep) + 1).ToString();
            HtmlNodeCollection instuctions = instuction.SelectNodes(String.Format(".//list{0}//l{0}item", subStep.ToString()));
            if (instuctions == null) { return; }
            int idx = 1;
            foreach (HtmlNode stepX in instuctions)
            {
                if (stepX.Attributes["group"] == null)
                {
                    stepX.Attributes["group"].Value = groupMain;
                }


                if (stepX.ParentNode.Attributes["group"] == null)
                {
                    stepX.ParentNode.Attributes.Add(stepX.OwnerDocument.CreateAttribute("group"));
                    stepX.ParentNode.Attributes["group"].Value = groupMain;
                }


                stepX.Attributes.Add(stepX.OwnerDocument.CreateAttribute("number"));
                stepX.Attributes["number"].Value = parentNumber + "." + idx.ToString();

                HtmlNodeCollection groupTagsALL = stepX.SelectNodes(".//para");

                if (groupTagsALL != null)
                {
                    for (int gt = 0; gt < groupTagsALL.Count; gt++)
                    {
                        HtmlNode thisStep = groupTagsALL[gt].ParentNode;
                        while (Regex.Match(thisStep.Name, @"l\d{1,1000}item").Success == false)
                        {
                            thisStep = thisStep.ParentNode;
                        }
                        if (thisStep.Attributes["group"] == null)
                        {
                            thisStep.Attributes.Add("group", "ALL");
                        }
                        if (groupTagsALL[gt].InnerText.Trim() == "")
                        {
                            groupTagsALL[gt].Remove();
                        }
                    }

                }

                HtmlNode groupNode = stepX.ParentNode.SelectSingleNode(".//title");
                if (groupNode != null)
                {
                    groupNode.ParentNode.SelectNodes(".//para").ToList().ForEach(delegate(HtmlNode nd)
                    {
                        if (nd.Attributes["title"] == null)
                        {
                            nd.Attributes.Add("title", "");
                        }
                        nd.Attributes["title"].Value = groupNode.InnerText;
                    });
                }


                string configCode = "";
                Match titleGroupsMathes = null;
                string groups = "";
                HtmlNodeCollection groupTags = stepX.SelectNodes(".//para[starts-with(.,'Group')]");
                if (groupTags != null)
                {

                    for (int gt = 0; gt < groupTags.Count; gt++)
                    {
                        groups = groupTags[gt].ParentNode.InnerText.Trim().Replace("\r", " ").Replace("\n", " ");
                        Match config = Regex.Match(groups, @"Configuration \d{1,6}");
                        configCode = "";
                        if (config.Success)
                        {
                            configCode = config.Value.Replace("Configuration ", "-C");
                        }
                        else
                        {
                            configCode = "-C0";
                        }
                        groups = Regex.Replace(groups, @"\t", " ");
                        groups = Regex.Replace(groups, @"\s{2,5000}", " ").Trim();
                        HtmlNode thisStep = groupTags[gt].ParentNode;
                        while (Regex.Match(thisStep.Name, @"l\d{1,1000}item").Success == false)
                        {
                            thisStep = thisStep.ParentNode;
                        }

                        titleGroupsMathes = Regex.Match(groups, @"Groups{0,1}\s(?<a>\d.*?):");
                        string groups_step = titleGroupsMathes.ToString();
                        groups_step = Regex.Replace(groups_step, @"Configuration.+?[;:]", "");
                        Match titleGroupsMathesRaw = Regex.Match(groups, @"Groups{0,1}\s(?<a>\d.*?):(?<b>.+)");
                        if (titleGroupsMathesRaw.Groups["b"].Value.Length > 3)
                        {
                            groupTags[gt].InnerHtml = groupTags[gt].InnerHtml.Replace(groups_step, "").Trim();
                            HtmlNode actualStep = stepX.OwnerDocument.CreateElement("para");
                            actualStep.InnerHtml = groups_step;
                            groupTags[gt].ParentNode.InsertBefore(actualStep, groupTags[gt]);
                        }
                        groups = groups_step;
                        groups = expandEffectivity_str(Regex.Replace(groups, @"[^\d\-\,]", "", RegexOptions.IgnoreCase), configCode);
                        thisStep.Attributes["group"].Value = groups.Replace(", ", ",");
                    }
                }

                idx++;
                subStep = subStep + 1;
                subStepsGroups(stepX, stepX.Attributes["group"].Value, stepX.Attributes["number"].Value);
            }
        }

        public string decodeGroupTitle(string title)
        {


            #region MyRegion
            string[] configs = title.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
            string combindedGropus = "";
            foreach (string configData in configs)
            {
                string configDataAlt = configData.Contains(":") ? configData : configData + " :";
                Match config = Regex.Match(configDataAlt, @"Configuration.+?[;:]");
                string configCode = "";
                if (config.Success)
                {
                    configCode = config.Value.Replace(":", "").Replace("Configuration ", "").Trim();
                    configCode = Regex.Replace(configCode, @"[;:]", "");
                }
                else
                {
                    configCode = "0";
                }
                List<string> configList = expandEffectivity_Array(configCode).Select(fd => "-C" + fd.ToString()).ToList();
                string grp = Regex.Replace(configDataAlt, @"Configuration.+?[;:]", " :");
                grp = grp.Contains(":") ? grp : grp + " :";
                grp = grp.Replace(":", " :");
                grp = Regex.Match(grp, @"Group.+?\d.+?:").Value.Trim();
                grp = Regex.Match(grp, @"Group.+?\d,{0,1}.+?[a-z:]", RegexOptions.IgnoreCase).Value;

                foreach (string c in configList)
                {
                    combindedGropus = combindedGropus + "," + expandEffectivity_str(Regex.Replace(grp, @"[^\d\-\,]", "", RegexOptions.IgnoreCase), c);
                }
                combindedGropus = combindedGropus.Trim(new char[] { ',', ' ' });
            }
            return combindedGropus;
            #endregion
        }

        public void getInstuctions(HtmlNodeCollection instuctions, List<string> groupdedConfigs, BackgroundWorker bw)
        {

            foreach (HtmlNode stepX in instuctions)
            {
                #region Tag Steps with Appliable Confiuration Groups
                HtmlNodeCollection groupTagsALL = stepX.SelectNodes(".//para");
                if (groupTagsALL != null)
                {
                    for (int gt = 0; gt < groupTagsALL.Count; gt++)
                    {
                        HtmlNode thisStep = groupTagsALL[gt].ParentNode;
                        while (Regex.Match(thisStep.Name, @"l\d{1,1000}item").Success == false)
                        {
                            thisStep = thisStep.ParentNode;
                        }
                        if (thisStep.InnerText.Contains("Group"))
                        {
                            MatchCollection fff = Regex.Matches(thisStep.InnerText, @"Group.+");
                        }
                        if (thisStep.Attributes["group"] == null)
                        {
                            thisStep.Attributes.Add("group", "ALL");
                        }
                        if (groupTagsALL[gt].Attributes["group"] == null)
                        {
                            groupTagsALL[gt].Attributes.Add("group", thisStep.Attributes["group"].Value);
                        }

                        if (groupTagsALL[gt].InnerText.Trim() == "")
                        {
                            groupTagsALL[gt].Remove();
                        }
                    }
                }
                string configCode = "";
                HtmlNodeCollection groupTags = stepX.ParentNode.SelectNodes(".//title");
                if (groupTags != null)
                {
                    #region MyRegion
                    foreach (HtmlNode groupNode in groupTags)
                    {
                        groupNode.ParentNode.ChildNodes.Where(fd => fd.NodeType == HtmlNodeType.Element).ToList().ForEach(delegate(HtmlNode nd)
                        {

                            if (nd.Attributes["title"] == null)
                            {
                                nd.Attributes.Add("title", groupNode.InnerText);
                            }
                            else
                            {
                                nd.Attributes["title"].Value = groupNode.InnerText;
                            }

                            if (nd.ParentNode.Attributes["title"] == null)
                            {
                                nd.ParentNode.Attributes.Add("title", groupNode.InnerText);
                            }
                            else
                            {
                                nd.ParentNode.Attributes["title"].Value = groupNode.InnerText;
                            }
                        });
                        bool isGroupChange = groupNode.InnerText.Trim().StartsWith("Group") || groupNode.InnerText.Trim().StartsWith("For Group");
                        if (isGroupChange)
                        {
                            #region MyRegion
                            string[] configs = groupNode.InnerText.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                            string combindedGropus = "";
                            foreach (string configData in configs)
                            {
                                string configDataAlt = configData.Contains(":") ? configData : configData + " :";
                                Match config = Regex.Match(configDataAlt, @"Configuration.+?[;:]");
                                configCode = "";
                                if (config.Success)
                                {
                                    configCode = config.Value.Replace(":", "").Replace("Configuration ", "").Trim();
                                    configCode = Regex.Replace(configCode, @"[;:]", "");
                                }
                                else
                                {
                                    configCode = "0";
                                }
                                List<string> configList = expandEffectivity_Array(configCode).Select(fd => "-C" + fd.ToString()).ToList();
                                string grp = Regex.Replace(configDataAlt, @"Configuration.+?[;:]", " :");
                                grp = grp.Contains(":") ? grp : grp + " :";
                                grp = grp.Replace(":", " :");
                                grp = Regex.Match(grp, @"Group.+?\d.+?:").Value.Trim();
                                grp = Regex.Match(grp, @"Group.+?\d,{0,1}.+?[a-z:]", RegexOptions.IgnoreCase).Value;

                                foreach (string c in configList)
                                {
                                    combindedGropus = combindedGropus + "," + expandEffectivity_str(Regex.Replace(grp, @"[^\d\-\,]", "", RegexOptions.IgnoreCase), c);
                                }
                                combindedGropus = combindedGropus.Trim(new char[] { ',', ' ' });
                            }
                            HtmlNode thisStep = groupNode;
                            if (groupNode.Attributes["group"] == null)
                            {
                                groupNode.Attributes.Add("group", combindedGropus);
                            }
                            else
                            {
                                groupNode.Attributes["group"].Value = combindedGropus;
                            }

                            #region MyRegion
                            if (groupNode.ParentNode.SelectNodes(".//para") != null)
                            {
                                groupNode.ParentNode.SelectNodes(".//para").ToList().ForEach(delegate(HtmlNode nd)
                                {
                                    if (nd.Attributes["group"] == null)
                                    {
                                        nd.Attributes.Add("group", combindedGropus);
                                    }
                                    else if (nd.Attributes["group"].Value == "ALL")
                                    {
                                        nd.Attributes["group"].Value = combindedGropus;
                                    }
                                });
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion
                }

                groupTags = stepX.SelectNodes(".//para[starts-with(.,'Group') or starts-with(.,'For Group')]");
                if (groupTags != null)
                {
                    #region MyRegion

                    for (int gt = 0; gt < groupTags.Count; gt++)
                    {

                        if (groupTags[gt].Attributes["tnode"] == null)
                        {
                            groupTags[gt].Attributes.Add("tnode", "true");
                        }
                        HtmlNode thisStep = groupTags[gt];


                        string[] configs = groupTags[gt].InnerText.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                        StringBuilder combindedGropus2 = new StringBuilder();
                        foreach (string configData in configs)
                        {
                            string configDataAlt = configData.Contains(":") ? configData : configData + " :";

                            Match config = Regex.Match(configDataAlt, @"Configuration.+?[;:]");

                            configCode = "";
                            if (config.Success)
                            {
                                configCode = config.Value.Replace("Configuration ", "");
                                configCode = Regex.Replace(configCode, @"[;:]", "");
                            }
                            else
                            {
                                configCode = "0";
                            }

                            List<string> configList = expandEffectivity_Array(configCode).Select(fd => "-C" + fd.ToString()).ToList();
                            string grp = Regex.Replace(configDataAlt, @"Configuration.+?[;:]", " :");
                            grp = Regex.Replace(grp, @"\t", " ");
                            grp = Regex.Replace(grp, @"\s{2,5000}", " ").Trim();
                            grp = grp.Contains(":") ? grp : grp + ":";
                            grp = grp.Replace(":", " :");
                            grp = Regex.Match(grp, @"Group.+?\d,{0,1}.+?[a-z:]", RegexOptions.IgnoreCase).Value;

                            foreach (string c in configList)
                            {
                                combindedGropus2.Append(expandEffectivity_str(Regex.Replace(grp, @"[^\d\-\,]", "", RegexOptions.IgnoreCase), c) + ",");
                            }
                        }


                        string combindedGropus = combindedGropus2.ToString().Trim(new char[] { ',', ' ' });


                        if (groupTags[gt].Attributes["group"] == null)
                        {
                            groupTags[gt].Attributes.Add("group", combindedGropus);
                        }
                        else if (groupTags[gt].Attributes["group"].Value == "")
                        {
                            groupTags[gt].Attributes["group"].Value = combindedGropus;
                        }
                        else if (groupTags[gt].Attributes["group"].Value == "ALL")
                        {
                            groupTags[gt].Attributes["group"].Value = combindedGropus;
                        }
                        else
                        {
                            groupTags[gt].Attributes["group"].Value = combindedGropus;
                        }



                        string allGroups = combindedGropus;


                        thisStep = groupTags[gt].ParentNode;

                        while (thisStep.Name != "l1item")
                        {
                            thisStep = thisStep.ParentNode;
                        }

                        if (thisStep.Attributes["group"] == null)
                        {
                            thisStep.Attributes.Add("group", combindedGropus);
                        }
                        else if (thisStep.Attributes["group"].Value == "ALL")
                        {
                            thisStep.Attributes["group"].Value = combindedGropus;
                        }
                        else
                        {
                            thisStep.Attributes["group"].Value += ("," + combindedGropus);
                        }

                        if (thisStep != null)
                        {
                            thisStep.SelectNodes(".//para").ToList().ForEach(delegate(HtmlNode nd)
                            {
                                if (nd.Attributes["tnode"] == null)
                                {
                                    thisStep = nd;
                                    if (thisStep.Attributes["group"] == null)
                                    {
                                        thisStep.Attributes.Add("group", combindedGropus);
                                    }
                                    else if (thisStep.Attributes["group"].Value == "ALL")
                                    {
                                        thisStep.Attributes["group"].Value = combindedGropus;
                                    }
                                    else
                                    {
                                        thisStep.Attributes["group"].Value += ("," + combindedGropus);
                                    }
                                }
                            });
                        }
                        thisStep = groupTags[gt].NextSibling;
                        while (thisStep != null)
                        {
                            if (thisStep.NodeType == HtmlNodeType.Element)
                            {
                                if (thisStep.Attributes["group"] == null)
                                {

                                    thisStep.Attributes.Add("group", combindedGropus);
                                }
                                else
                                {
                                    thisStep.Attributes["group"].Value = combindedGropus;
                                }
                            }

                            thisStep = thisStep.NextSibling;
                        }
                    }
                    #endregion
                }

                groupTags = stepX.SelectNodes(".//para[not(@tnode)]");
                if (groupTags != null)
                {
                    #region MyRegion
                    for (int gt = 0; gt < groupTags.Count; gt++)
                    {
                        HtmlNode thisStep = groupTags[gt];
                        while (thisStep.NextSibling != null)
                        {
                            thisStep = thisStep.NextSibling;
                            if (thisStep.Attributes["tnode"] != null)
                            {
                                break;
                            }
                            if (Regex.Match(thisStep.Name, @"list\d").Success || Regex.Match(thisStep.Name, @"\ditem").Success || Regex.Match(thisStep.Name, @"table").Success)
                            {
                                break;
                            }
                            if ("numitem unitem caution warining note".Contains(thisStep.ParentNode.Name))
                            {
                                break;
                            }
                            else if (thisStep.NodeType == HtmlNodeType.Element && thisStep.ParentNode == groupTags[gt].ParentNode)
                            {
                                bool addnode = true;
                                if (thisStep.Name == "para")
                                {
                                    addnode = (thisStep.Attributes["group"].Value == groupTags[gt].Attributes["group"].Value);

                                }
                                if (addnode)
                                {
                                    if (thisStep.Attributes["skip"] == null)
                                    {
                                        thisStep.Attributes.Add("skip", "true");
                                    }

                                    thisStep.ParentNode.RemoveChild(thisStep);
                                    thisStep.InnerHtml = " " + thisStep.InnerHtml;
                                    groupTags[gt].InnerHtml += (thisStep.Name == "para" ? thisStep.InnerHtml : thisStep.OuterHtml);
                                    thisStep = groupTags[gt];
                                }
                            }
                        }

                    }
                    #endregion
                }

                HtmlNodeCollection stepTitles = stepX.SelectNodes(".//title");
                if (stepTitles != null)
                {
                    #region MyRegion
                    foreach (HtmlNode titleNode in stepTitles)
                    {
                        if (Regex.Match(titleNode.ParentNode.Name, @"\ditem").Success || Regex.Match(titleNode.ParentNode.Name, @"list\d").Success)
                        {
                            HtmlNode firstStep = titleNode.ParentNode.SelectSingleNode(".//para");
                            firstStep.InnerHtml = String.Format("[b]{0}[/b]", titleNode.InnerHtml) + firstStep.InnerHtml;
                        }
                    }
                    #endregion
                }
                #endregion
            }

            FigureGraphics = new NameValueCollection();

            instuctions.ToList().ForEach(delegate(HtmlNode nd)
            {
                if (nd.Attributes["group"] != null)
                {
                    nd.Attributes["group"].Value = string.Join(",", nd.Attributes["group"].Value.Split(new string[] { ",", " " }, StringSplitOptions.RemoveEmptyEntries).ToList().Distinct());
                }
            });

            groupdedConfigs = groupdedConfigs.Where(fd => !fd.StartsWith("ALL")).ToList();


            using (StreamWriter InstructionsWriter = new StreamWriter(activeSb_dir + "\\configurations.xml", false))
            {
                string fsds = "<configurations><config>" + String.Join("</config><config>", groupdedConfigs.ToArray().Distinct()) + "</config></configurations>";
                InstructionsWriter.WriteLine(fsds);
                InstructionsWriter.Close();
            }
            using (StreamWriter InstructionsWriter = new StreamWriter(activeSb_dir + "\\instructions.xml", false))
            {

                InstructionsWriter.WriteLine("<instuctions>");

                string group = "ALL";
                int counter = 0;
                int instructionNumber = 0;
                int nCheck = 0;
                string workPackage = "";

                FigureInstrunctions = new NameValueCollection();
                instrunctionCollection = new NameValueCollection();
                instrunctionTables = new NameValueCollection();

                #region Create Group-Config WIC
                string wic_idx = "1";
                InstructionsWriter.WriteLine("<wic id='1'>");
                foreach (HtmlNode instructionNode in instuctions.ToList())
                {
                    counter++;
                    HtmlNode instuctionTextL1 = null;

                    int curLevel = 0;
                    HtmlNodeCollection groupedSteps = instructionNode.SelectNodes(".//*[@group]");

                    List<string> applicableGroups = instructionNode.Attributes["group"].Value.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries).ToList();

                    HtmlNode[] nodes = instructionNode.SelectNodes(".//para").Where(fd => fd.Attributes["tnode"] == null).Where(fd => fd.Attributes["skip"] == null).Where(fd => Regex.Match(fd.ParentNode.Name, @"l\d+item").Success).ToArray();

                    string instructionNumberSTR = instructionNumber.ToString();
                    string instrTxt = "";

                    foreach (HtmlNode nodePara in nodes)
                    {
                        instuctionTextL1 = nodePara;
                        HtmlNode enumNode = instuctionTextL1;

                        instructionNumberSTR = instructionNumberSTR.Replace(wic_idx + "||", "");
                        #region Calculate Steps Number
                        if (nodePara.Attributes["wic_idx"].Value != wic_idx)
                        {
                            InstructionsWriter.WriteLine("</wic>");
                            wic_idx = nodePara.Attributes["wic_idx"].Value;
                            counter = 0;
                            instructionNumber = 0;
                            nCheck = 0;
                            workPackage = "";
                            InstructionsWriter.WriteLine("<wic id='" + wic_idx + "'>");
                        }



                        while (Regex.Match(enumNode.Name, @"l(?<a>\d+)item").Success == false)
                        {
                            enumNode = enumNode.ParentNode;
                            if (enumNode != null)
                            {
                                Match itemNbr = Regex.Match(enumNode.Name, @"l(?<a>\d+)item");
                                if (itemNbr.Success)
                                {
                                    int n = int.Parse(itemNbr.Groups["a"].Value);
                                    nCheck = Regex.Matches(instructionNumberSTR, @"\.").Count + 1;
                                    if (n == 1)
                                    {
                                        instructionNumber++;
                                        instructionNumberSTR = instructionNumber.ToString();
                                    }
                                    else if (curLevel == n)
                                    {
                                        instructionNumberSTR = Regex.Replace(instructionNumberSTR, @"\d+\z", delegate(Match m)
                                        {
                                            return (int.Parse(m.Value) + 1).ToString();
                                        });
                                    }
                                    else if (curLevel < n)
                                    {
                                        instructionNumberSTR += ".1";
                                    }
                                    else if (curLevel > n)
                                    {
                                        instructionNumberSTR = instructionNumberSTR.Substring(0, instructionNumberSTR.LastIndexOf("."));
                                        instructionNumberSTR = Regex.Replace(instructionNumberSTR, @"\d+\z", delegate(Match m)
                                        {
                                            return (int.Parse(m.Value) + 1).ToString();
                                        });
                                    }
                                    curLevel = n;
                                    nCheck = Regex.Matches(instructionNumberSTR, @"\.").Count + 1;
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                        }
                        #endregion
                        instuctionTextL1.Attributes.Add("step", instructionNumberSTR);
                        instructionNumberSTR = wic_idx + "||" + instructionNumberSTR;

                        if (instuctionTextL1.ParentNode == instructionNode)
                        {
                            #region Mark Title
                            if (instructionNode.Attributes["title"] != null)
                            {
                                if (instructionNode.Attributes["title"].Value != workPackage)
                                {
                                    if (instuctionTextL1.InnerHtml.StartsWith("[b]") == false)
                                    {
                                        instuctionTextL1.InnerHtml = String.Format("[b]{0}[/b]", instructionNode.Attributes["title"].Value) + instuctionTextL1.InnerHtml;
                                    }
                                }
                                workPackage = instructionNode.Attributes["title"].Value;
                            }
                            #endregion
                        }

                        #region Format instrTxt
                        instrTxt = instuctionTextL1.InnerHtml.Trim().Replace("\n", " ");
                        if (instuctionTextL1.Attributes["group"] != null)
                        {
                            instrTxt = instrTxt.Contains(":") ? instrTxt.Replace(":", " :") : instrTxt + " :";
                            instrTxt = Regex.Replace(instrTxt, @"Group.+?:", "");
                            if (instrTxt == "")
                            {
                                instrTxt = instuctionTextL1.InnerHtml.Trim().Replace("\n", " ");
                            }
                        }
                        instrTxt = instrTxt.Trim(new char[] { ' ', ':' });
                        instrTxt = HttpUtility.HtmlDecode(instrTxt);
                        instrTxt = Regex.Replace(instrTxt, @" {2,500}", " ");

                        instrTxt = Regex.Replace(instrTxt, @"[^>""]FIGURE\s+?\d{1,}[\,]{0,1}\s+?and\s+?\d{1,}", delegate(Match m)
                        {
                            string gg = m.Value.ToUpper();
                            gg = gg.Replace("FIGURES", "").Trim();
                            gg = gg.Replace("FIGURE", "");
                            gg = Regex.Replace(gg, @"\d{1,}", delegate(Match m2)
                            {
                                return "FIGURE " + m2.Value;
                            }, RegexOptions.IgnoreCase);
                            return gg;

                        }, RegexOptions.IgnoreCase);
                        #endregion

                        instuctionTextL1.InnerHtml = instrTxt;

                        instrunctionCollection.Add(group, instructionNumberSTR + ".::" + instrTxt);

                        #region Get Tables used in Instruction
                        HtmlNodeCollection tableNodesColl = null;
                        List<HtmlNode> tableNodes = new List<HtmlNode>();
                        tableNodesColl = instuctionTextL1.ParentNode.SelectNodes("table");
                        if (tableNodesColl != null)
                        {
                            tableNodes = tableNodesColl.ToList();
                            //tableNodes = tableNodesColl.Cast<HtmlNode>().Where(fd => fd.Attributes["group"] == null ? true :
                            //(fd.Attributes["group"].Value.Contains(group) || fd.Attributes["group"].Value.Contains("ALL") ||
                            //fd.Attributes["group"].Value.Contains(groupPrams[0] + "-C0"))
                            //).ToList();
                        }
                        updateTableHisto(group + "_" + instructionNumberSTR, tableNodes);
                        #endregion

                        List<HtmlNode> figureNodes = new List<HtmlNode>();
                        HtmlNodeCollection htmlColl = instuctionTextL1.ParentNode.SelectNodes(".//refint[@reftype='Figure']");

                        #region Redundant
                        //HtmlNodeCollection htmlCollfTbles = instuctionTextL1.ParentNode.SelectNodes("./table//refint[@reftype='Figure']");
                        //if (htmlCollfTbles != null)
                        //{
                        //    figureNodes.AddRange(htmlCollfTbles.ToList());
                        //} 
                        #endregion

                        if (htmlColl != null)
                        {
                            figureNodes.AddRange(htmlColl.ToList());
                        }
                        #region No longer Used Xmlreader used instead of HtmlDocument
                        //else
                        //{
                        //    MatchCollection figureMatches = Regex.Matches(instrTxt.ToUpper(), @"FIGURE\s\d{1,10000}.+?[a-zA-z\.]{0,1}", RegexOptions.IgnoreCase);
                        //    int ss = figureMatches.Count;
                        //    HtmlNodeCollection graphicCol = instuctionTextL1.OwnerDocument.DocumentNode.SelectNodes("//graphic");
                        //    if (graphicCol != null)
                        //    {
                        //        foreach (Match figMatch in figureMatches)
                        //        {
                        //            string gg = figMatch.Value;
                        //            MatchCollection figNbr = Regex.Matches(gg, @"\d{1,100000}");
                        //            figNbr.Cast<Match>().ToList().ForEach(delegate(Match mFigMatch)
                        //            {
                        //                figureNodes.AddRange(collectFigures(mFigMatch.Value, graphicCol, figureNodes, instuctionTextL1));
                        //            });
                        //        }
                        //    }
                        //} 
                        #endregion

                        #region Redundant
                        //if (tableNodes.Count > 0)
                        //{
                        //    foreach (HtmlNode tableNode in tableNodes)
                        //    {
                        //        htmlColl = tableNode.SelectNodes(".//refint[@reftype='Figure']");
                        //        if (htmlColl != null)
                        //        {
                        //            figureNodes.AddRange(htmlColl.ToList());
                        //        }
                        //        else
                        //        {
                        //            string tableText = tableNode.InnerHtml;
                        //            MatchCollection figureMatches = Regex.Matches(tableText.ToUpper(), @"FIGURE\s\d{1,10000}.+?[a-zA-z\.]{0,1}", RegexOptions.IgnoreCase);
                        //            HtmlNodeCollection graphicCol = tableNode.OwnerDocument.DocumentNode.SelectNodes("//graphic");
                        //            if (graphicCol != null)
                        //            {
                        //                foreach (Match figMatch in figureMatches)
                        //                {
                        //                    string gg = figMatch.Value;
                        //                    MatchCollection figNbr = Regex.Matches(gg, @"\d{1,100000}");
                        //                    figNbr.Cast<Match>().ToList().ForEach(delegate(Match mFigMatch)
                        //                    {
                        //                        figureNodes.AddRange(collectFigures(mFigMatch.Value, graphicCol, figureNodes, instuctionTextL1));
                        //                    });
                        //                }
                        //            }
                        //            if (htmlColl != null)
                        //            {
                        //                figureNodes.AddRange(htmlColl.ToList());
                        //            }
                        //        }
                        //    }
                        //} 
                        #endregion

                        if (figureNodes.Count > 0)
                        {
                            #region MyRegion
                            foreach (HtmlNode figureNode in figureNodes)
                            {
                                updateGraphics(figureNode, instuctionTextL1);

                                if (FigureInstrunctions[figureNode.InnerText] == null)
                                {
                                    FigureInstrunctions.Add(figureNode.InnerText, group + "_" + instructionNumberSTR);
                                    InstrunctionsFigures.Add(group + "_" + instructionNumberSTR.ToString(), figureNode.InnerText);
                                }
                                else
                                {
                                    if (FigureInstrunctions.GetValues(figureNode.InnerText).Contains(group + "_" + instructionNumberSTR) == false)
                                    {
                                        FigureInstrunctions.Add(figureNode.InnerText, group + "_" + instructionNumberSTR);
                                        InstrunctionsFigures.Add(group + "_" + instructionNumberSTR.ToString(), figureNode.InnerText);
                                    }
                                }
                            }
                            #endregion
                        }
                        else
                        {
                            #region Figures not taged so use Regex
                            MatchCollection FigureMatches = Regex.Matches(instrTxt, @"(FIGURE(?<a>\s+?)\d{1,})", RegexOptions.IgnoreCase);
                            foreach (Match FigureMatch in FigureMatches)
                            {
                                string fig = FigureMatch.Value.ToUpper();
                                HtmlNode figureNode = findGraphic(fig);
                                if (figureNode != null)
                                {
                                    fig = figureNode.InnerText;
                                    if (FigureInstrunctions[fig] == null)
                                    {
                                        FigureInstrunctions.Add(fig, group + "_" + instructionNumberSTR);
                                        InstrunctionsFigures.Add(group + "_" + instructionNumberSTR.ToString(), fig);
                                    }
                                    else
                                    {
                                        if (FigureInstrunctions.GetValues(fig).Contains(group + "_" + instructionNumberSTR) == false)
                                        {
                                            FigureInstrunctions.Add(fig, group + "_" + instructionNumberSTR);
                                            InstrunctionsFigures.Add(group + "_" + instructionNumberSTR.ToString(), fig);
                                        }
                                    }
                                    updateGraphics(figureNode, instuctionTextL1);
                                }
                            }
                            #endregion
                        }

                        InstructionsWriter.WriteLine(instuctionTextL1.OuterHtml);
                    }
                }
                InstructionsWriter.WriteLine("</wic>");
                InstructionsWriter.WriteLine("</instuctions>");
                InstructionsWriter.Close();
                InstructionsWriter.Dispose();
            }
            #endregion
            saveTables(instrunctionTables);
            updateFiguresIndex(bw, FigureInstrunctions);
        }

        public void saveTables(NameValueCollection instrunction_tables)
        {
            string path = activeSb_dir + "\\Tables";
            string tpath = activeSb_dir + "\\Tables\\tables.xml";
            XmlDocument tables_index = new XmlDocument();
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            if (!File.Exists(path + "\\tables.xml"))
            {
                tables_index.LoadXml("<tables></tables>");
                tables_index.Save(tpath);
            }
            else
            {
                tables_index.Load(path + "\\tables.xml");
            }
            XmlDocument figure_index = new XmlDocument();
            figure_index.Load(activeSb_dir + "\\figure_index.xml");
            foreach (string key in instrunction_tables.Keys)
            {
                string[] pramaInstr = key.Split(new string[] { ".::" }, StringSplitOptions.RemoveEmptyEntries);
                string nbr = pramaInstr[0].Trim().Replace("ALL_", "");
                string[] tableNodes = instrunction_tables.GetValues(key);
                foreach (string tmarkUp in tableNodes)
                {
                    string[] tableMarkUpPrams = tmarkUp.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                    string tableMarkUp = tableMarkUpPrams[1].Trim();
                    string tableID = tableMarkUpPrams[0].Trim();

                    HtmlAgilityPack.HtmlDocument tableSGML = new HtmlAgilityPack.HtmlDocument();
                    tableSGML.LoadHtml(String.Format("{0}", tableMarkUp));
                    string htmlMarkup = convert2HtmlTable(tableSGML.DocumentNode, null);

                    path = activeSb_dir + "\\Tables\\" + tableID + ".html";
                    if (File.Exists(path) == false)
                    {
                        using (StreamWriter sw = new StreamWriter(path))
                        {
                            sw.Write(htmlMarkup);
                            sw.Close();
                            sw.Dispose();
                        }
                    }


                    XmlNode tableNode = tables_index.DocumentElement.SelectSingleNode("//table[@id='" + tableID + "']");
                    if (tableNode == null)
                    {
                        tableNode = tables_index.CreateElement("table");
                        tableNode.Attributes.Append(tables_index.CreateAttribute("id")).Value = tableID;
                        tableNode.InnerXml = "<instr/>";
                        tableNode.SelectSingleNode("instr").InnerXml = nbr;
                        tables_index.DocumentElement.AppendChild(tableNode);
                    }
                    else
                    {
                        tableNode.SelectSingleNode("instr").InnerXml += "," + nbr;
                    }

                    tables_index.Save(tpath);
                    string group = "ALL";
                    if (tableSGML.DocumentNode.SelectSingleNode("table").Attributes["group"] != null)
                    {
                        group = tableSGML.DocumentNode.SelectSingleNode("table").Attributes["group"].Value;
                    }
                    #region MyRegion
                    string instrTxt = tableMarkUp;
                    MatchCollection FigureMatches = Regex.Matches(instrTxt, @"(FIGURE(?<a>\s+)\d{1,5})", RegexOptions.IgnoreCase);
                    foreach (Match FigureMatch in FigureMatches)
                    {
                        string fig = FigureMatch.Value.ToUpper();
                        XmlNodeList idx_figs = figure_index.DocumentElement.SelectNodes("//fig[text()[contains(.,'" + fig + "')]]");
                        if (idx_figs.Count > 0)
                        {
                            string[] groups = group.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                            string refid = idx_figs[0].Attributes["refid"].Value;
                            HtmlNode ff = new HtmlNode(HtmlNodeType.Element, new HtmlAgilityPack.HtmlDocument(), 0);
                            ff.Attributes.Add("refid", refid);
                            ff.InnerHtml = fig.ToUpper();
                            foreach (string g in groups)
                            {
                                FigureInstrunctions.Add(FigureMatch.Value, g.Trim() + "_" + nbr);
                            }
                            updateGraphics(ff, null);

                        }

                    }
                    #endregion

                }
            }
        }


        public HtmlNode findGraphic(string figureName)
        {
            List<string> filePaths = Directory.GetFiles(activeSb_dir + "\\graphics", "*.xml").ToList();
            foreach (string file in filePaths)
            {
                XmlReader reader = XmlReader.Create(file);
                reader.ReadToDescendant("graphic");
                reader.ReadToDescendant("title");
                string pattern = figureName.Replace(" ", @".+?")+":";
                string title = reader.ReadInnerXml().ToUpper();
                if (Regex.Match(title, pattern, RegexOptions.IgnoreCase).Success)
                {
                    reader.Close();
                    HtmlAgilityPack.HtmlDocument ff = new HtmlAgilityPack.HtmlDocument();
                    ff.Load(file);
                    HtmlNode r = ff.DocumentNode.ChildNodes[0];
                    string key = r.Attributes["key"].Value;
                    r.InnerHtml = figureName;
                    r.Attributes.Add("refid", key);
                    return r;
                }

            }
            return null;
        }
        
        public void updateFiguresIndex(BackgroundWorker bw, NameValueCollection figureInstrunctions)
        {

            XmlDocument figure_index = new XmlDocument();
            figure_index.Load(activeSb_dir + "\\figure_index.xml");

            XmlDocument figuresIndexDOC = new XmlDocument();
            string pathFigures = activeSb_dir + "\\FIGURES\\";
            if (Directory.Exists(pathFigures) == false)
            {
                Directory.CreateDirectory(pathFigures);
            }



            pathFigures = pathFigures + "Figures.xml";
            if (File.Exists(pathFigures) == false)
            {
                figuresIndexDOC.LoadXml("<NewDataSet/>");
            }
            else
            {
                figuresIndexDOC.Load(pathFigures);
            }
            StringBuilder sb = new StringBuilder();

            
            string[] iFigs = figureInstrunctions.Keys.Cast<string>().ToArray();
            Array.Sort(iFigs, new FigureSorter());
            foreach (string str in iFigs)
            {
                string[] graphics = FigureGraphics.GetValues(str);
                foreach (string val in figureInstrunctions.GetValues(str))
                {
                    string[] figParams = val.Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);

                    string tableList_figs = "";
                    string graphicStringList = "";
                    string figId = str.Replace("FIGURE", "").Trim();

                    XmlNode figureNode = figuresIndexDOC.SelectSingleNode("//Figure[@id='" + figId + "']");
                    if (figureNode == null)
                    {
                        figureNode = figuresIndexDOC.CreateElement("Figure");
                        figureNode.Attributes.Append(figuresIndexDOC.CreateAttribute("id")).Value = figId;

                        XmlNode figure = figuresIndexDOC.CreateElement("figure");
                        XmlNode graphic = figuresIndexDOC.CreateElement("graphic");
                        XmlNode group = figuresIndexDOC.CreateElement("group");
                        XmlNode instruction = figuresIndexDOC.CreateElement("instruction");
                        XmlNode tables = figuresIndexDOC.CreateElement("tables");

                        figureNode.AppendChild(figure);
                        figureNode.AppendChild(graphic);
                        figureNode.AppendChild(group);
                        figureNode.AppendChild(instruction);
                        figureNode.AppendChild(tables);

                        if (graphics !=  null)
                        {
                            foreach (string graphicName in graphics)
                            {
                                XmlNode sheet = figuresIndexDOC.CreateElement("sheet");
                                sheet.Attributes.Append(figuresIndexDOC.CreateAttribute("id")).Value = graphicName;
                                string[] tableIds = FigureTables.GetValues(str + "_" + graphicName);
                                if (tableIds != null)
                                {
                                    tableIds = tableIds.Distinct().ToArray();
                                    foreach(string tid in tableIds)
                                    {
                                        XmlNode table = figuresIndexDOC.CreateElement("table");
                                        table.Attributes.Append(figuresIndexDOC.CreateAttribute("id")).Value = tid;
                                        sheet.AppendChild(table);
                                    }                                                                        
                                }
                                figureNode.AppendChild(sheet);
                            }
                        }
                        else
                        {
                            graphicStringList = "";
                        }

                        string figNbr = str.Replace("FIGURE", "").Trim();
                        tables.InnerXml = tableList_figs;
                        group.InnerXml = figParams[0];
                        figure.InnerXml = figNbr;
                        graphic.InnerXml = graphicStringList;
                        instruction.InnerXml = figParams[0] + "::" + figParams[1];


                        
                        XmlNode refidNode = figure_index.DocumentElement.SelectSingleNode("//fig[text() = 'FIGURE " + figNbr + "']");
                        if (refidNode != null)
                        {
                            XmlDocument graphicDoc = new XmlDocument();
                            graphicDoc.Load(activeSb_dir + "\\graphics\\" + refidNode.Attributes["refid"].Value+".xml");
                            refidNode = graphicDoc.DocumentElement.SelectSingleNode("//title");
                            if (refidNode != null)
                            {
                                Match regMatch = Regex.Match(refidNode.InnerText, @"\[.+\]");
                                if (regMatch.Success)
                                {
                                  string combindedGropus=  decodeGroupTitle(regMatch.Value);
                                  group.InnerXml = combindedGropus;
                                }
                            }
                        }
                        
                        
                        figuresIndexDOC.DocumentElement.AppendChild(figureNode);
                    }
                    else
                    {
                        if (!figureNode.SelectSingleNode("instruction").InnerXml.Contains(figParams[0] + "::" + figParams[1]))
                        {
                            figureNode.SelectSingleNode("instruction").InnerXml += ("," + figParams[0] + "::" + figParams[1]);
                        }
                        if (!figureNode.SelectSingleNode("group").InnerXml.Contains(figParams[0]))
                        {
                            if (figureNode.SelectSingleNode("group").InnerXml == "")
                            {
                                figureNode.SelectSingleNode("group").InnerXml = figParams[0];
                            }
                            
                        }
                    }



                    sb.Append(val + ", ");
                    bw.ReportProgress(13, new string[] { val, str });

                }
                sb.AppendLine("=".PadLeft(200, '='));
                sb.AppendLine();
            }
            figuresIndexDOC.Save(pathFigures);
            GC.Collect();



            bw.ReportProgress(4, sb.ToString());
            sb.Clear();

        }
        public void getAppendix()
        {
            
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant("append");
            while (reader.Name == "append")
            {                
               string key = reader.GetAttribute("key");
               string data = reader.ReadOuterXml();
               String f = activeSb_dir + "\\" + key + ".xml";
               XmlDocument xdoc = new XmlDocument();
               xdoc.LoadXml(data);

               XmlNodeList graphicSheets = xdoc.SelectNodes("//sheet");
               foreach (XmlNode graphicSheet in graphicSheets)
               {
                   if (graphicSheet.Attributes["gnbr"] == null)
                   {
                       continue;
                   }
                   string sheetFile = activeSb_dir + "\\" + graphicSheet.Attributes["gnbr"].Value + ".xml";
                   string grFileVal = "";
                   if (doctypeEntity[graphicSheet.Attributes["gnbr"].Value] != null)
                   {
                       grFileVal = doctypeEntity[graphicSheet.Attributes["gnbr"].Value].ToString();
                       grFileVal = grFileVal.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries).Last();
                       graphicSheet.Attributes.Append(xdoc.CreateAttribute("img")).Value = grFileVal;
                   }

                   string refid = graphicSheet.Attributes["key"].Value;
                   if (refid.Contains("-"))
                   {
                       refid = refid.Split(new[] { "-" }, StringSplitOptions.RemoveEmptyEntries).Last();
                       refid = Regex.Replace(refid, @"\d{1,}\Z", delegate(Match m)
                       {
                           return "_" + m.Value;
                       });
                       refid = refid.Substring(1);
                   }
                   graphicSheet.Attributes.Append(xdoc.CreateAttribute("imgalt")).Value = refid;

                   using (StreamWriter sw = new StreamWriter(sheetFile))
                   {
                       sw.Write(graphicSheet.OuterXml);
                       sw.Close();
                       sw.Dispose();
                   }
               }


               xdoc.Save(f);
            }

            reader.Close();
        }
        public void getSbGraphics()
        {

            String f = activeSb_dir + "\\graphics";
            if (Directory.Exists(f) == false)
            {
                Directory.CreateDirectory(f);
            }
            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant("graphic");
            while (match)
            {
                XmlReader inner = reader.ReadSubtree();
                inner.Read();                
                while (inner.EOF == false && inner.Name == "graphic")
                {
                    string key = reader.GetAttribute("key").ToUpper() + ".xml";
                    data = inner.ReadOuterXml();
                    string graphicFile = string.Format("{0}\\{1}", f, key);


                    XmlDocument fig = new XmlDocument();
                    fig.LoadXml(data);
                    XmlNode title = fig.DocumentElement.SelectSingleNode("//title");
                    if (title != null)
                    {
                        Match ConfigGroup = Regex.Match(title.InnerText.Trim(), @"Config.+\s(?<a>.+)\:");
                        if (ConfigGroup.Success)
                        {
                            string configs = ConfigGroup.Groups["a"].Value;
                            configs = configs.Trim().Replace("thru", "-");
                            configs = expandEffectivity_str(configs, "-C1");
                            fig.DocumentElement.Attributes.Append(fig.CreateAttribute("group")).Value = configs;
                        }
                        ConfigGroup = null;
                    }
                    

                    XmlNodeList tables = fig.SelectNodes("//table");
                    if (tables != null)
                    {
                        foreach (XmlNode table in tables)
                        {
                            if (table.Attributes["id"] != null)
                            {
                                title = fig.SelectSingleNode("//title");
                                table.AppendChild(title);
                                String path = activeSb_dir + "\\graphics\\" + table.Attributes["id"].Value + ".html";
                                convert2HtmlTable(table, path);
                                path = "";
                            }
                        }
                    }
                    if (isAirbus)
                    {
                        XmlNodeList graphicSheets = fig.SelectNodes("//sheet");
                        foreach (XmlNode graphicSheet in graphicSheets)
                        {
                            if (graphicSheet.Attributes["gnbr"] == null)
                            {
                                continue;
                            }
                            string sheetFile = activeSb_dir + "\\" + graphicSheet.Attributes["gnbr"].Value + ".xml";
                            string grFileVal = "";
                            if (doctypeEntity[graphicSheet.Attributes["gnbr"].Value] != null)
                            {
                                grFileVal = doctypeEntity[graphicSheet.Attributes["gnbr"].Value].ToString();
                                grFileVal = grFileVal.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries).Last();
                                graphicSheet.Attributes.Append(fig.CreateAttribute("img")).Value = grFileVal;
                            }

                            string refid = graphicSheet.Attributes["key"].Value;
                            if (refid.Contains("-"))
                            {
                                refid = refid.Split(new[] { "-" }, StringSplitOptions.RemoveEmptyEntries).Last();
                                refid = Regex.Replace(refid, @"\d{1,}\Z", delegate(Match m)
                                {
                                    return "_" + m.Value;
                                });
                                refid = refid.Substring(1);
                            }
                            graphicSheet.Attributes.Append(fig.CreateAttribute("imgalt")).Value = refid;

                            using (StreamWriter sw = new StreamWriter(sheetFile))
                            {
                                sw.Write(graphicSheet.OuterXml);
                                sw.Close();
                            }
                            sheetFile = sheetFile = "";
                        }
                        graphicSheets = null;
                    }
                    tables =  null;
                    fig.Save(graphicFile);
                    fig = new XmlDocument();
                    GC.Collect();
                    GC.Collect();
                }
                while (reader.Name != "graphic" )
                {
                   match = reader.Read();
                }
                match = reader.ReadToFollowing("graphic");
                
            }

            reader.Close();
        }
        public void getSbInstrs_sb()
        {
            String f = activeSb_dir + "\\ainstr";
            if (Directory.Exists(f) == false)
            {
                Directory.CreateDirectory(f);
            }
            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            NameValueCollection stepsNbrs = new NameValueCollection();
            reader.ReadToDescendant("sb");

            bool match = reader.ReadToDescendant("instsect");
            int wic_part = 1;
            while (match && reader.Name == "instsect")
            {
                
                string key = reader.GetAttribute("key").ToString();
                match = reader.ReadToDescendant("title");
                string title = reader.ReadInnerXml().ToUpper();
                title = title.Replace(" ", "_");
                if (reader.Name != "list1")
                {
                    match = reader.ReadToNextSibling("list1");
                }
                if (match)
                {
                    if (reader.Name == "list1" && !reader.EOF && reader.IsStartElement())
                    {
                        string group = "All";
                        string listlevel = "1";
                        string stnbr = "1";
                        int step = 0;
                        XmlReader list1 = reader.ReadSubtree();
                        int currentLevel = 0;
                        #region MyRegion
                        while (list1.EOF == false)
                        {
                            if (Regex.Match(list1.Name, @"list\d{1,}").Success && reader.IsStartElement())
                            {
                                listlevel = Regex.Match(list1.Name, @"list(?<a>\d{1,})").Groups["a"].Value;
                                int level = int.Parse(listlevel);
                                group = level == 1 ? "All" : group;
                                list1.Read();
                            }
                            else if (Regex.Match(list1.Name, @"l(?<a>\d{1,})item").Success && reader.IsStartElement())
                            {
                                listlevel = Regex.Match(list1.Name, @"l(?<a>\d{1,})item").Groups["a"].Value;
                                int level = int.Parse(listlevel);
                                if (level == 1)
                                {
                                    step++;
                                }
                                else if (level > currentLevel)
                                {
                                    string[] steps = stnbr.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                                    Array.Resize(ref steps, level);
                                    steps[level - 1] = "1";
                                    stnbr = String.Join(".", steps);
                                }
                                else if (level < currentLevel)
                                {
                                    int ff = int.Parse(stnbr.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries)[level - 1]);
                                    ff++;
                                    string[] steps = stnbr.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                                    Array.Resize(ref steps, level);
                                    steps[level - 1] = ff.ToString();
                                    stnbr = String.Join(".", steps);
                                }
                                else if (level == currentLevel)
                                {
                                    int ff = int.Parse(stnbr.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries)[level - 1]);
                                    ff++;
                                    string[] steps = stnbr.Split(new[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                                    Array.Resize(ref steps, level);
                                    steps[level - 1] = ff.ToString();
                                    stnbr = String.Join(".", steps);
                                }
                                currentLevel = level;
                                list1.Read();
                            }
                            else if (list1.Name != "" && reader.IsStartElement())
                            {
                                XmlDocument node = new XmlDocument();
                                string nodeName = list1.Name;
                                data = list1.ReadOuterXml();
                                if (data != "")
                                {

                                    XmlDocument instrtable = new XmlDocument();
                                    instrtable.LoadXml(data);

                                    XmlNodeList tables = instrtable.SelectNodes(".//table");
                                    if (tables != null)
                                    {
                                        foreach (XmlNode tbl in tables)
                                        {
                                            HtmlAgilityPack.HtmlDocument tdoc = new HtmlAgilityPack.HtmlDocument();
                                            tdoc.OptionWriteEmptyNodes = true;
                                            tdoc.OptionAutoCloseOnEnd = true;
                                            tdoc.OptionOutputAsXml = true;
                                            tdoc.LoadHtml(tbl.OuterXml);
                                            XmlNode newTable = instrtable.CreateElement("div");                                            
                                            string htmlTbl = convert2HtmlTable(tdoc.DocumentNode, null);                                            
                                            tdoc.LoadHtml(htmlTbl);
                                            newTable.InnerXml = tdoc.DocumentNode.InnerHtml;
                                            tbl.ParentNode.ReplaceChild(newTable, tbl);
                                        }
                                    }

                                    data = instrtable.OuterXml;
                                    string data2 = instrtable.InnerText;
                                    MatchCollection configMatch = Regex.Matches(data2, @"config\.{0,1}.\d{1,}", RegexOptions.IgnoreCase);
                                    if (configMatch.Count > 0 && nodeName == "para")
                                    {
                                        group = "";
                                        List<string> grps = new List<string>();

                                        foreach (Match m in configMatch)
                                        {
                                            data2 = m.Value;
                                            group = String.Join(",", Regex.Matches(data2, @"\d{1,}").Cast<Match>().Select(fd => fd.Value.PadLeft(3, '0') + "-C1").ToList());
                                            grps.Add(group);
                                        }
                                        group = String.Join(",", grps);
                                    }
                                    node.LoadXml("<para/>");
                                    node.DocumentElement.InnerXml = data;
                                    if (node.InnerText != "")
                                    {
                                        node.DocumentElement.Attributes.Append(node.CreateAttribute("group")).Value = group;
                                        node.DocumentElement.Attributes.Append(node.CreateAttribute("listlevel")).Value = listlevel.ToString();
                                        node.DocumentElement.Attributes.Append(node.CreateAttribute("step")).Value = stnbr;
                                        stepsNbrs.Add(stnbr, node.OuterXml);
                                    }
                                }
                            }
                            else
                            {
                                list1.Read();
                            }
                        }
                        #endregion
                        key = "wicpart_" + title + "_" + wic_part.ToString().PadLeft(3,'0') + ".xml";
                        using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", f, key), false))
                        {
                            sw.Write("<instuctions><wic id='" + wic_part + "'>");
                            foreach (string keyStep in stepsNbrs.Keys)
                            {
                                HtmlNode node = HtmlNode.CreateNode("<para/>");
                                node.Attributes.Add("step", keyStep);
                                List<string> groups = new List<string>();
                                foreach (string step_xml in stepsNbrs.GetValues(keyStep))
                                {
                                    HtmlNode childNode = HtmlNode.CreateNode(step_xml);

                                    childNode.ChildNodes.ToList().ForEach(fd => fd.Attributes.Add("group", childNode.Attributes["group"].Value));

                                    groups.Add(childNode.Attributes["group"].Value);
                                    node.AppendChildren(childNode.ChildNodes);
                                }
                                node.Attributes.Add("group", String.Join(",", groups.Distinct().ToList()));
                                sw.Write(node.OuterHtml);
                            }
                            sw.Write("</wic></instuctions>");
                            sw.Close();
                        }
                        wic_part++;
                    }
               
                }
                while (reader.IsStartElement() == false && reader.EOF == false)
                {
                    match = true;
                    reader.ReadEndElement();
                }
            }
            reader.Close();
        }
        public void getSbInstrs()
        {
            if (SBNum.Contains("A300") || SBNum.Contains("A310"))
            {
                getSbInstrs_sb();
                return;
            }

            String f = activeSb_dir + "\\ainstr";
            if (Directory.Exists(f) == false)
            {
                Directory.CreateDirectory(f);
            }
            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);

            reader.ReadToDescendant("sb");
      
            bool match = reader.ReadToDescendant("instsect");

            while ("AWIN" != reader.GetAttribute("key").ToUpper() || !match)
            {
                match = reader.ReadToNextSibling("instsect");
                if (!match) { break; }
            }

            bool alt_list1 = false;
            if (reader.GetAttribute("key") == null)
            {
                reader = XmlReader.Create(activeSb, readerSetting);
                reader.ReadToDescendant("sb");
                match = reader.ReadToDescendant("instsect");
                match = reader.ReadToDescendant("title");
                string d = reader.ReadInnerXml().ToUpper();

                while ("WORK INSTRUCTIONS" != d || !match)
                {
                    match = reader.ReadToFollowing("instsect");

                    if (!match)
                    {
                        break;
                    }
                    match = reader.ReadToDescendant("title");
                    d = reader.ReadInnerXml().ToUpper();
                }

                alt_list1 = match = "WORK INSTRUCTIONS" == d;

            }
            else
            {
                match = "AWIN" == reader.GetAttribute("key").ToUpper();
            }



            if (match)
            {
                int wic_part = 1;
                if (alt_list1)
                {
                    if (reader.Name != "list1")
                    {
                        match = reader.ReadToNextSibling("list1");
                    }
                }
                else
                {
                    match = reader.ReadToDescendant("list1");
                }
                if (match)
                {
                    while (reader.Name == "list1")
                    {
                        string key = "wicpart_" + wic_part.ToString() + ".xml";
                        data = (reader.ReadOuterXml());

                        using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", f, key), false))
                        {
                            sw.Write(data);
                            sw.Close();
                        }
                        wic_part++;
                    }
                }
                else
                {
                    string key = "wicpart.xml";
                    data = (reader.ReadOuterXml());
                    using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", f, key), false))
                    {
                        sw.Write(data);
                        sw.Close();
                    }
                }
            }
            reader.Close();
        }

        public void getSbInstrsGen()
        {

            String f = activeSb_dir + "\\ainstr_gen";
            if (Directory.Exists(f) == false)
            {
                Directory.CreateDirectory(f);
            }
            string data = "";
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");

            bool match = reader.ReadToDescendant("instsect");

            while ("AGI" != reader.GetAttribute("key").ToUpper() || !match)
            {
                match = reader.ReadToNextSibling("instsect");
                if (!match) { break; }
            }
            if (match && "AGI" == reader.GetAttribute("key").ToUpper())
            {
                string key = "general.xml";
                data = reader.ReadOuterXml();
                using (StreamWriter sw = new StreamWriter(string.Format("{0}\\{1}", f, key), false))
                {
                    sw.Write(data);
                    sw.Close();
                }
            }
            reader.Close();
        }

        public NameValueCollection getEffectivity(BackgroundWorker bw, out List<string> configurations)
        {
            string data = "";
            configurations = new List<string>();
            HtmlAgilityPack.HtmlDocument effxrefDoc = new HtmlAgilityPack.HtmlDocument();
            String f = activeSb_dir + "\\effectivity.sect";
            bool airbusLogic = isAirbus;
            if (File.Exists(f))
            {
                effxrefDoc.Load(f);
                HtmlNode configuration = effxrefDoc.DocumentNode.SelectSingleNode("//table");
                configuration = effxrefDoc.DocumentNode.SelectSingleNode(".//table//*[text()='CONFIGURATION']"); 
                HtmlNode airbusNode = effxrefDoc.DocumentNode.SelectSingleNode("//*[text()='Configuration by MSN']");
                if (airbusLogic)
                {
                    #region MyRegion
                    NameValueCollection configHisto = new NameValueCollection();
                    HtmlNodeCollection effbyOp = effxrefDoc.DocumentNode.SelectNodes("//*[text()='Effectivity by Operator']");
                    HtmlNodeCollection effByKit = effxrefDoc.DocumentNode.SelectNodes("//*[text()='Effectivity by MSN and Kit/Configuration']");
                    HtmlNodeCollection configCols = effxrefDoc.DocumentNode.SelectNodes("//*[text()='CONFIGURATION']");
                    NameValueCollection custMSN = new NameValueCollection();
                    HtmlNodeCollection msnbrAll = effxrefDoc.DocumentNode.SelectNodes("//msnbr");
                    HtmlNodeCollection custNodes = effxrefDoc.DocumentNode.SelectNodes("//cus");
                    if (custNodes != null)
                    {
                        foreach (HtmlNode cust_X in custNodes)
                        {
                            HtmlNode rowX = cust_X.Ancestors("tr").First();
                            HtmlNodeCollection msnbrs_X = rowX.SelectNodes(".//msnbr");
                            if (msnbrs_X != null)
                            {
                                msnbrs_X.ToList().ForEach(fd => custMSN.Add(cust_X.InnerText, fd.InnerText));
                                msnbrs_X.ToList().ForEach(fd => configHisto.Add("1", "SN" + fd.InnerText));
                            }
                        }
                    }
                    if (configCols != null)
                    {
                        #region MyRegion
                        if (effByKit != null)
                        {
                            configHisto = new NameValueCollection();
                            foreach (HtmlNode config_X in configCols)
                            {

                                int colidx = config_X.Ancestors("tr").First().ChildNodes.IndexOf(config_X.Ancestors("td").First());
                                colidx++;
                                HtmlNodeCollection configCells = config_X.Ancestors("table").First().SelectNodes(".//td[" + colidx.ToString() + "]");
                                if (configCells != null)
                                {
                                    int testInt = 0;
                                    List<string> configVals = configCells.Select(fd => fd.InnerText).Where(fd => int.TryParse(fd, out testInt)).ToList();
                                    if (configVals.Count == 0)
                                    {
                                        configVals.Add("1");
                                    }
                                    foreach (string configVal_X in configVals)
                                    {
                                        string configVal = configVal_X.TrimStart(new[] { '0' }).Trim();
                                        List<string> msn_list = new List<string>();
                                        HtmlNode table_X = config_X.Ancestors("table").First();
                                        HtmlNode msnPara = table_X.ParentNode.PreviousSibling;
                                        HtmlNodeCollection genrange = msnPara.SelectNodes(".//genrange");
                                        HtmlNodeCollection msnbrVals = msnPara.SelectNodes(".//msnbr");
                                        if (msnbrVals != null)
                                        {
                                            msn_list.AddRange(msnbrVals.Select(fd => fd.InnerText).ToList());
                                        }
                                        if (genrange != null)
                                        {
                                            msn_list.AddRange(
                                                genrange.SelectMany(fd =>

                                                    expandEffectivity_StringList(
                                                    String.Join("-", Regex.Matches(fd.InnerHtml, @"\d{1,}").Cast<Match>().Select(fdx => fdx.Value).ToArray())).Cast<string>()

                                                    ).ToList());
                                        }
                                        msn_list = msn_list.Distinct().ToList();
                                        msn_list.Sort();
                                        msn_list.ForEach(fd => configHisto.Add(configVal, "SN" + fd));
                                    }
                                }
                            }
                        }
                        else if (effbyOp != null)
                        {
                            configHisto = new NameValueCollection();
                            #region MyRegion
                            if (airbusNode != null)
                            {
                                airbusNode = effxrefDoc.DocumentNode.SelectSingleNode("//*[text()='Effectivity by Operator']");
                                airbusNode = airbusNode.ParentNode.SelectSingleNode(".//table//p[text()='OPERATOR']");
                                if (airbusNode != null)
                                {
                                    configuration = airbusNode.Ancestors("table").ToArray()[0];
                                }

                                if (configuration != null)
                                {
                                    configuration = effxrefDoc.DocumentNode.SelectSingleNode("//*[text()='CONFIGURATION']");
                                    if (configuration != null)
                                    {
                                        airbusLogic = configuration.Ancestors("table") != null;
                                        configuration = airbusLogic ? configuration.Ancestors("table").Cast<HtmlNode>().ToList()[0] : null;
                                    }
                                }
                            }
                            #endregion
                            #region MyRegion

                            configuration = effxrefDoc.DocumentNode.SelectSingleNode("//cus[text()='" + airline_operator + "']");
                            if (configuration != null)
                            {

                                HtmlNodeCollection msnNbrs = configuration.Ancestors("td").ToArray()[0].NextSibling.SelectNodes(".//msnbr");
                                configuration = effxrefDoc.DocumentNode.SelectSingleNode("//td/p[text()='CONFIGURATION']");
                                if (configuration != null)
                                {
                                    var dsdvsd = configuration.Ancestors("table").Cast<HtmlNode>().ToList();
                                    airbusLogic = configuration.Ancestors("table") != null;
                                    configuration = airbusLogic ? configuration.Ancestors("table").Cast<HtmlNode>().ToList()[0] : null;
                                    if (configuration != null)
                                    {
                                        foreach (HtmlNode msnX in msnNbrs)
                                        {
                                            try
                                            {
                                                string msnNbr = msnX.InnerHtml;
                                                List<HtmlNode> configs = configuration.SelectNodes(".//msnbr[text()='" + msnNbr + "']").ToList();
                                                configs = configs.ToList().Select(fd => fd.Ancestors("td").ToArray()[0].NextSibling).ToList();
                                                if (configs.Count > 0)
                                                {
                                                    configs.ForEach(fx =>
                                                        fx.InnerText.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries).ToList().ForEach(fd =>
                                                        configHisto.Add(fd.TrimStart(new[] { '0' }), "SN" + msnNbr)));
                                                }
                                                else
                                                {
                                                    configHisto.Add("1", "SN" + msnNbr);
                                                }
                                            }
                                            catch
                                            {
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                        }

                        #endregion
                    }

                    else if (msnbrAll != null)
                    {
                        configHisto = new NameValueCollection();
                        msnbrAll.ToList().ForEach(fd => configHisto.Add("1", "SN" + fd.InnerText));
                    }
                    #region MyRegion
                    configurations = configHisto.Keys.Cast<string>().Select(fd => fd.PadLeft(3, '0') + "-C1").ToList();
                    configurations.Sort();
                    sortVariableNames(configHisto);

                    List<string> groupdedConfigs = configurations.Select(fd => string.Format("<config>{0}</config>", fd)).ToList();
                    XmlDocument configDoc = new XmlDocument();
                    configDoc.LoadXml("<configuration/>");
                    configDoc.DocumentElement.InnerXml = String.Join("", groupdedConfigs.ToArray().Distinct());
                    configDoc.Save(activeSb_dir + "\\configurations.xml");
                    return configHisto;
                    #endregion 
                    #endregion
                }

                if (configuration == null || isAirbus)
                {
                    return null;
                }
                if (configuration != null)
                {
                    configuration = configuration.Ancestors("table").First();
                    configuration = configuration.SelectSingleNode("tbody");
                    #region MyRegion
                    configurations = new List<string>();
                    HtmlNodeCollection cells = configuration.SelectNodes(".//td[@colname='1']");

                    if (cells != null)
                    {
                        foreach (HtmlNode cell_x in cells)
                        {
                            HtmlNode cell = cell_x;
                            string group = cell.InnerText.Trim().PadLeft(3, '0');
                            int configs = Math.Max(0, int.Parse(cell.Attributes["rowspan"].Value) - 1);
                            if (configs == 0)
                            {
                                configurations.Add(group + "-C1");
                            }
                            else if (configs == 1)
                            {
                                if (cell.ParentNode.NextSibling != null)
                                {
                                    string config = cell.ParentNode.NextSibling.SelectSingleNode("td").InnerText;
                                    configurations.Add(group + "-C" + config);
                                }
                                else
                                {
                                    configurations.Add(group + "-C1");
                                }
                            }
                            else if (configs > 1)
                            {
                                while (configs > 0)
                                {
                                    if (cell.ParentNode.NextSibling != null)
                                    {
                                        cell = cell.ParentNode.NextSibling.SelectSingleNode("td");
                                        string config = cell.InnerText;
                                        configurations.Add(group + "-C" + config);
                                    }
                                    configs--;
                                }
                            }

                        }
                    }
                    #endregion
                }
            }
            configurations = configurations.Distinct().ToList();
            configurations.Sort();
            #region MyRegion
            //List<string> groupdedConfigs = configurations.Select(fd => string.Format("<config>{0}</config>", fd)).ToList();
            //XmlDocument configDoc = new XmlDocument();
            //configDoc.LoadXml("<configuration/>");
            //configDoc.DocumentElement.InnerXml = String.Join("", groupdedConfigs.ToArray().Distinct());
            //configDoc.Save(activeSb_dir + "\\configurations.xml");            
            #endregion
            XmlReader reader = XmlReader.Create(activeSb, readerSetting);
            reader.ReadToDescendant("sb");
            bool match = reader.ReadToDescendant("plansect");
            while ("EFF" != reader.GetAttribute("sectname").ToUpper() || !match)
            {
                match = reader.ReadToNextSibling("plansect");
                if (!match)
                {
                    break;
                }
            }
            if (match)
            {
                data = reader.ReadOuterXml();
            }
            reader.Close();
            effxrefDoc.LoadHtml(data);
            HtmlNode effxref = effxrefDoc.DocumentNode.SelectSingleNode("//table[@id='insetTbl']");
            NameValueCollection TailGroupHisto = new NameValueCollection();
            if (effxref != null)
            {
                while (effxref.Name.ToLower() != "table")
                {
                    effxref = effxref.ParentNode;
                }
                HtmlNodeCollection tail_Groups = effxref.SelectNodes(".//tbody/row");


                foreach (HtmlNode row in tail_Groups)
                {
                    HtmlNodeCollection tailNodes = row.SelectNodes("entry/para/effxref/effdata//sunit");
                    if (tailNodes != null)
                    {
                        foreach (HtmlNode tailNode in tailNodes)
                        {
                            string tail = tailNode.InnerText;
                            string group = row.SelectSingleNode("entry[@colname='2']/para").InnerText;
                            TailGroupHisto.Add(group, tail);
                        }
                    }
                }
            }
            else
            {
                HtmlNodeCollection tail_Groups = effxrefDoc.DocumentNode.SelectNodes(".//row");
                if (tail_Groups != null)
                {
                    foreach (HtmlNode row in tail_Groups)
                    {
                        HtmlNodeCollection tailNodes = row.SelectNodes("entry/para/effxref/effdata//sunit");
                        if (tailNodes == null) { continue; }
                        foreach (HtmlNode tailNode in tailNodes)
                        {
                            string tail = tailNode.InnerText;
                            string group = row.SelectSingleNode("entry[@colname='2']/para").InnerText;
                            TailGroupHisto.Add(group, tail);
                        }
                    }
                }
            }
            FedExGroups = new List<string>();
            TailGroupHisto = sortVariableNames(TailGroupHisto);
            return TailGroupHisto;
        }

        public HtmlNode getInstrustions_All()
        {
            String f = activeSb_dir + "\\ainstr";
            HtmlAgilityPack.HtmlDocument fullWic = new HtmlAgilityPack.HtmlDocument();
            StringBuilder wicsBuilder = new StringBuilder();
            if (Directory.Exists(f) != false)
            {
                string[] wicFiles = Directory.GetFiles(f, "*.xml").ToArray();
                int wicIdx = 1;
                wicFiles = wicFiles.OrderBy(fd => int.Parse(Regex.Match(Path.GetFileNameWithoutExtension(fd), @"\d{1,100}").Value)).ToArray();
                wicFiles.ToList().ForEach(delegate(string fd)
                {
                    HtmlAgilityPack.HtmlDocument ff = new HtmlAgilityPack.HtmlDocument();
                    ff.Load(fd);
                    ff.DocumentNode.SelectNodes("//para").ToList().ForEach(delegate(HtmlNode nd)
                    {
                        nd.Attributes.Add("wic_idx", wicIdx.ToString());
                    });
                    wicsBuilder.AppendLine(ff.DocumentNode.InnerHtml);
                    wicIdx++;
                });
                fullWic.LoadHtml("<wic>" + wicsBuilder.ToString() + "</wic>");
            }
            return fullWic.DocumentNode;
        }

        public void ProcessBoeingSB_XML(HtmlAgilityPack.HtmlDocument sbXML, BackgroundWorker bw)
        {

            List<string> configurations = new List<string>();
            NameValueCollection TailGroupHisto = new NameValueCollection();
            TailGroupHisto = getEffectivity(bw, out configurations);
            HtmlNode instsect = null;
            if (SBNum.Contains("A300") || SBNum.Contains("A310"))
            {
                instsect = getInstrustions_All();
                if (instsect.SelectSingleNode("//instuctions") != null)
                {
                    using (StreamWriter InstructionsWriter = new StreamWriter(activeSb_dir + "\\instructions.xml", false))
                    {
                        InstructionsWriter.Write(instsect.InnerHtml);
                        InstructionsWriter.Close();
                    }
                }
                
                return;
            }

            instsect = getInstrustions_All();
            HtmlNodeCollection instuctions = instsect.SelectNodes(".//list1//l1item");
            

            GC.Collect();

            if (instuctions == null) { return; }

            getInstuctions(instuctions, configurations, bw);

            GC.Collect();

            StringBuilder sb = new StringBuilder();

            #region For Progress Reporting in GUI
            //foreach (string group in instrunctionCollection.Keys)
            //{
            //    sb.AppendLine("Group " + group);
            //    string[] group_Instuctions = instrunctionCollection.GetValues(group);
            //    foreach (string instr in group_Instuctions)
            //    {

            //        string instNbr = instr.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries)[0];
            //        sb.Append(instNbr.Trim().Trim(new char[] { '.' }));
            //        sb.Append(",\t");
            //    }
            //    sb.AppendLine();
            //    sb.AppendLine("=".PadLeft(200, '='));
            //    sb.AppendLine();

            //}
            //bw.ReportProgress(2, sb.ToString());

            //sb.Clear();
            //string[] instrKeys = instrCollection.Keys.Cast<string>().ToArray();
            //Array.Sort(instrKeys, new InstrunctionSorter2());
            //foreach (string str in instrKeys)
            //{
            //    sb.Append(str + ". ");
            //    sb.AppendLine(instrCollection[str] + "\r\n");
            //    sb.AppendLine("=".PadLeft(200, '='));
            //    sb.AppendLine();
            //}
            //bw.ReportProgress(3, sb.ToString());
            //sb.Clear(); 
            #endregion


            bw.ReportProgress(12, "");


            string gpath = Application.StartupPath + "\\cgms";
            Directory.CreateDirectory(gpath);



            
            bw.ReportProgress(10, ConfigurationGroups.Where(fd => fd.Contains("ALL") == false).ToList());
            

            #region deprecated
            //ZipArchive zip = ZipFile.Open(compresedFile, ZipArchiveMode.Update);
            //ZipArchiveEntry entry =  zip.GetEntry("cgms\\" + SBNum + ".zip");
            //entry.Delete();

            //HtmlNodeCollection materialTablesColl = sbXML.DocumentNode.SelectNodes("//matinfo//table");
            //List<HtmlNode> materialTables = new List<HtmlNode>();
            //if(materialTablesColl != null)
            //{
            //    materialTables = materialTablesColl.ToList();
            //}

            // groupTables_Material_Media = new NameValueCollection();
            // groupTables_Material_KITS = new NameValueCollection();
            // groupTables_Material_WIRES = new NameValueCollection();
            //foreach (HtmlNode tableNode in materialTables)
            //{
            //    HtmlNode tableTitle = tableNode.SelectSingleNode("title");
            //    HtmlNode tableBody = tableNode.SelectSingleNode("tgroup");
            //    StringBuilder tableFooter = new StringBuilder();
            //    if (tableNode.SelectNodes("ftnote") != null)
            //    {
            //        List<HtmlNode> ftnotes = tableNode.SelectNodes("ftnote").ToList();                    
            //        ftnotes.ForEach(delegate(HtmlNode note)
            //        {
            //            note.SelectSingleNode("para").InnerHtml = "(" + note.SelectSingleNode("para").InnerHtml.Trim() + ") ";
            //            note.InnerHtml = "<entry colname=\"1\">" + note.InnerHtml + "</entry>";
            //            tableFooter.Append(note.OuterHtml.Replace("ftnote", "row"));

            //        });


            //    }
            //    if (tableBody != null)
            //    {
            //        HtmlNode dsd = tableBody.SelectSingleNode(".//tbody");
            //        dsd.InnerHtml += tableFooter.ToString();
            //        HtmlNode thead = tableNode.SelectSingleNode("tgroup//thead/row/entry");
            //        string tableType = null;
            //        if (thead != null)
            //        {
            //             tableType = tableNode.SelectSingleNode("tgroup//thead/row/entry").InnerText.Trim().ToUpper();
            //        }


            //        string tableGroup = tableTitle != null ? tableTitle.InnerText : "ALL";
            //        List<object> groups = new List<object>();
            //        if (tableGroup == "ALL" || tableGroup.ToLower().Contains("group") == false)
            //        {
            //            groups = new List<object>();
            //            foreach (string groupKey in TailGroupHisto.Keys)
            //            {
            //                groups.Add(groupKey.Trim());
            //            }
            //        }
            //        else
            //        {
            //            groups = expandEffectivity_Array(tableGroup.Replace("Group", "").Replace(":", ""));
            //        }
            //        foreach (object grp in groups)
            //        {
            //            if (tableType == null) { continue; }
            //            if (tableType.StartsWith("MEDIA SET"))
            //            {
            //                groupTables_Material_Media.Add(grp.ToString(), tableBody.InnerHtml);
            //                groupTables_Material_KITS.Add(grp.ToString(), tableBody.OuterHtml);

            //            }
            //            else if (tableType.StartsWith("KIT"))
            //            {                     
            //                groupTables_Material_KITS.Add(grp.ToString(), tableBody.InnerHtml);
            //            }
            //            else if (tableType.Contains("PART NUMBER"))
            //            {
            //                HtmlAgilityPack.HtmlDocument dfdf = new HtmlAgilityPack.HtmlDocument();
            //                dfdf.LoadHtml(tableBody.InnerHtml);
            //                HtmlNode theader = dfdf.DocumentNode.SelectSingleNode("//thead/row");
            //                HtmlNode tbody = dfdf.DocumentNode.SelectSingleNode("//tbody");

            //                HtmlNode topRow = dfdf.CreateElement("row");
            //                topRow.InnerHtml = theader.InnerHtml.Trim();
            //                tbody.InsertBefore(topRow, tbody.FirstChild);
            //                topRow = dfdf.CreateElement("row");

            //                groupTables_Material_KITS.Add(grp.ToString(), tbody.OuterHtml);
            //            }
            //            else if (tableType.StartsWith("WIRE LIST") || tableType.Contains("LIST - CUT AND CODE AS REQUIRED"))
            //            {
            //                groupTables_Material_WIRES.Add(grp.ToString(), tableBody.InnerHtml);
            //            }
            //        }
            //    }
            //}
            //GC.Collect();
            //saveMaterialData_WIRES(TailGroupHisto.Keys.Cast<string>().ToArray(), bw);
            //GC.Collect();
            //saveMaterialData_KITS(TailGroupHisto.Keys.Cast<string>().ToArray(), bw);

            //try
            //{
            //    HtmlNode manpowerData = null;
            //    //manpowerData = sbXML.DocumentNode.SelectNodes("//plansect/title").SingleOrDefault(fd => fd.InnerHtml.Trim().ToLower() == "manpower");
            //    //if (manpowerData != null)
            //    //{
            //    //    manpowerData = manpowerData.ParentNode;
            //    //    manHours = ExtractManHoursXML(manpowerData);
            //    //    SaveManHoursAllGroups();
            //    //    bw.ReportProgress(15, Collection2String(manHours));
            //    //}
            //}
            //catch { 
            //} 
            #endregion
        }

        public string convert2HtmlTable(XmlNode tnode, string saveloc)
        {
            HtmlNode n = (new HtmlAgilityPack.HtmlDocument()).CreateElement("table");
            n.InnerHtml = tnode.InnerXml;
            return convert2HtmlTable(n, saveloc);
        }

        public string convert2HtmlTable(HtmlNode tnode, string saveloc)
        {
            tnode.InnerHtml = Regex.Replace(tnode.InnerHtml, @"COL[A-Z]*(?<a>\d{1,})", delegate(Match m) { return m.Groups["a"].Value; }, RegexOptions.IgnoreCase);
            HtmlNode ttable = tnode.Name.ToLower() == "table" ? tnode : tnode.SelectSingleNode("table");
            HtmlNode t_title = tnode.SelectSingleNode(".//title");
            HtmlNode tgroup = tnode.SelectSingleNode(".//tgroup");
            HtmlNode theadsb = tnode.SelectSingleNode(".//thead");
            HtmlNode tbodysb = tnode.SelectSingleNode(".//tbody");
            HtmlNodeCollection colspec = tnode.SelectNodes(".//colspec");
            HtmlNodeCollection spanSpecs = tnode.SelectNodes(".//spanspec");
            HtmlNodeCollection allRows = tnode.SelectNodes(".//row");

            int rowpanBase = 1;
            bool useMorerows = tbodysb.SelectNodes(".//entry/@morerows") != null;
            if (useMorerows)
            {
                //rowpanBase = tnode.SelectNodes(".//entry/@morerows").Select(fd => int.Parse(fd.Attributes["morerows"].Value)).Min();
            }
            HtmlNodeCollection ftnote = tnode.SelectNodes(".//ftnote");

            HtmlAgilityPack.HtmlDocument tabFigDoc = new HtmlAgilityPack.HtmlDocument();
            tabFigDoc.OptionWriteEmptyNodes = true;
            tabFigDoc.OptionAutoCloseOnEnd = true;            
            tabFigDoc.LoadHtml("<table/>");

            HtmlNode table = tabFigDoc.CreateElement("table");
            HtmlNode thead = tabFigDoc.CreateElement("thead");
            HtmlNode tfoot = tabFigDoc.CreateElement("table");
            HtmlNode tbody = tabFigDoc.CreateElement("tbody");

            ttable.Attributes.ToList().ForEach(delegate(HtmlAttribute attr)
            {

                table.Attributes.Add(attr.Name, attr.Value);

            });

            HtmlNodeCollection rows;

            rows = ttable.SelectNodes(".//row");
            for (int rowNbr = 0; rowNbr < rows.Count; rowNbr++)
            {
                HtmlNode row = rows[rowNbr];
              HtmlNodeCollection cells = row.SelectNodes(".//entry");
                for (int cellIdx = 0; cellIdx < cells.Count; cellIdx++)
                {
                    #region MyRegion
                    HtmlNode cell = cells[cellIdx];
                    if (cell.Attributes["colname"] != null)
                    {
                        cell.Attributes.Add("calrowspan", "true");
                    }
                    cell.Attributes.Add("touched", "true");
                    cell.Attributes.Add("cellIdx", cellIdx.ToString());
                    if (cell.Attributes["namest"] != null)
                    {
                        int nameend = int.Parse(cell.Attributes["nameend"].Value);
                        string cspans = cell.Attributes["namest"].Value + "-" + nameend.ToString();
                        List<object> colspaces = expandEffectivity_Array(cspans);
                        cell.Attributes.Add("colname", cell.Attributes["namest"].Value);
                        cell.Attributes.Add("cspase", String.Join(",", colspaces.Select(fd => fd.ToString())));
                    }
                    else if (cell.Attributes["spanname"] != null)
                    {
                        #region MyRegion
                        string specName = cell.Attributes["spanname"].Value;

                        HtmlNode spanspec = spanSpecs.SingleOrDefault(fd => fd.Attributes["spanname"].Value == specName);
                        if (spanspec != null)
                        {
                            int namest = int.Parse(spanspec.Attributes["namest"].Value);
                            int nameend = int.Parse(spanspec.Attributes["nameend"].Value);
                            string cspans = namest.ToString() + "-" + nameend.ToString();
                            List<object> colspaces = expandEffectivity_Array(cspans);
                            cell.Attributes.Add("cspase", String.Join(",", colspaces.Select(fd => fd.ToString())));
                            cell.Attributes.Add("colname", namest.ToString());
                            cell.Attributes.Add("namest", namest.ToString());
                            cell.Attributes.Add("nameend", nameend.ToString());
                        }
                        else
                        {
                            MatchCollection spanTo = Regex.Matches(cell.Attributes["spanname"].Value, @"\d{1,}");
                            if (spanTo.Count == 2)
                            {
                                int namest = int.Parse(spanTo[0].Value);
                                int nameend = int.Parse(spanTo[1].Value) + 1;
                                string cspans = namest.ToString() + "-" + nameend.ToString();
                                List<object> colspaces = expandEffectivity_Array(cspans);
                                cell.Attributes.Add("cspase", String.Join(",", colspaces.Select(fd => fd.ToString())));
                                cell.Attributes.Add("colname", namest.ToString());
                                cell.Attributes.Add("namest", namest.ToString());
                                cell.Attributes.Add("nameend", nameend.ToString());
                            }
                        } 
                        #endregion
                    }
                    else if (cell.Attributes["colname"] != null)
                    {

                        int indexof = cell.ParentNode.ChildNodes.IndexOf(cell);
                        int namest = indexof + 1;
                        bool testName = int.TryParse(cell.Attributes["colname"].Value, out namest);
                        if (testName)
                        {
                            if (indexof == 0 && namest > 1)
                            {
                                indexof++;
                                while (indexof < namest)
                                {
                                    HtmlNode emptyCell = HtmlNode.CreateNode("<entry></entry>");
                                    emptyCell.Attributes.Add("colname", (indexof).ToString());
                                    emptyCell.Attributes.Add("dynode", "true");
                                    emptyCell.Attributes.Add("cspase", indexof.ToString());
                                    emptyCell.Attributes.Add("namest", indexof.ToString());
                                    emptyCell.Attributes.Add("nameend", indexof.ToString());
                                    cell.ParentNode.InsertBefore(emptyCell, cell);
                                    cells = row.SelectNodes(".//entry");
                                    indexof++;
                                    cellIdx++;
                                }
                            }


                            if (cell.NextSibling != null)
                            {
                                if (cell.NextSibling.Attributes["colname"] != null)
                                {
                                    int nameend = int.Parse(cell.NextSibling.Attributes["colname"].Value);

                                    if (nameend > namest + 1)
                                    {
                                        HtmlNode emptyCell = HtmlNode.CreateNode("<entry></entry>");
                                        emptyCell.Attributes.Add("colname", (namest + 1).ToString());
                                        emptyCell.Attributes.Add("cspase", (namest + 1).ToString());
                                        emptyCell.Attributes.Add("dynode", "true");
                                        cell.ParentNode.InsertAfter(emptyCell, cell);
                                        cells = row.SelectNodes(".//entry");
                                    }
                                }
                                string cspans = namest.ToString() + "-" + namest.ToString();
                                List<object> colspaces = expandEffectivity_Array(cspans);
                                cell.Attributes.Add("cspase", String.Join(",", colspaces.Select(fd => fd.ToString())));
                            }
                            else
                            {
                                string cspans = namest.ToString() + "-" + namest.ToString();
                                List<object> colspaces = expandEffectivity_Array(cspans);
                                cell.Attributes.Add("cspase", String.Join(",", colspaces.Select(fd => fd.ToString())));
                            }
                        }
                        else
                        {
                            cell.Attributes.Add("colname", namest.ToString());
                            cell.Attributes.Add("cspase", namest.ToString());
                        }
                        cell.Attributes.Add("namest", namest.ToString());
                        cell.Attributes.Add("nameend", namest.ToString());
                    }
                    else
                    {
                        int namest = cell.ParentNode.ChildNodes.IndexOf(cell) + 1;
                        cell.Attributes.Add("colname", namest.ToString());
                        cell.Attributes.Add("cspase", namest.ToString());
                        cell.Attributes.Add("namest", namest.ToString());
                        cell.Attributes.Add("nameend", namest.ToString());
                        continue;
                    }
                    #endregion
                }
            }


            if (theadsb != null)
            {
                rows = theadsb.SelectNodes(".//row");
                int gg = int.Parse(tgroup.Attributes["cols"].Value);

                foreach (HtmlNode row in rows)
                {
                    HtmlNode tr = tabFigDoc.CreateElement("tr");

                    List<HtmlNode> cells = row.SelectNodes(".//entry").ToList();

                    cells = cells.OrderBy(fd => fd.Attributes["namest"] != null ? int.Parse(fd.Attributes["namest"].Value) : 0).ToList();
                    int rowspan = 1;
                    int cellCnt = 1;
                    foreach (HtmlNode cell in cells)
                    {
                        HtmlNode td = tabFigDoc.CreateElement("td");

                        int colName = cell.Attributes["namest"] != null ? (int.Parse(cell.Attributes["namest"].Value)) : cellCnt;
                        int nameend = cell.Attributes["nameend"] != null ? (int.Parse(cell.Attributes["nameend"].Value) + 1) : colName + 1;



                        td.Attributes.Add("colspan", (nameend - colName).ToString());
                        rowspan = cell.Attributes["morerows"] != null ? int.Parse(cell.Attributes["morerows"].Value) + rowpanBase : 1;
                        string morerows = rowspan.ToString();
                        cell.Attributes.Add("rowspan", morerows);
                        td.Attributes.Add("rowspan", morerows);
                        td.Attributes.Add("namest", colName.ToString());
                        td.Attributes.Add("nameend", (nameend - 1).ToString());
                        td.InnerHtml = cell.InnerHtml.Trim();
                        tr.AppendChild(td);
                        cellCnt++;
                    }
                    thead.AppendChild(tr);
                }
            }


            #region ftnote
            if (ftnote != null)
            {
                rows = ftnote;
                tfoot.Attributes.Add("class", "fnote");

                foreach (HtmlNode row in rows)
                {
                    HtmlNode tr = tabFigDoc.CreateElement("tr");

                    List<HtmlNode> cells = row.SelectNodes("para").ToList();
                    HtmlNode td = tabFigDoc.CreateElement("td");
                    td.InnerHtml = row.InnerHtml.Trim().Replace("<para", "<span").Replace("para>", "span>");
                    tr.AppendChild(td);
                    tfoot.AppendChild(tr);
                }
            }
            #endregion


            rows = tbodysb.SelectNodes(".//row");
            int rowIdx = 0;

            foreach (HtmlNode row in rows)
            {
                HtmlNode tr = tabFigDoc.CreateElement("tr");
                List<HtmlNode> cells = row.SelectNodes(".//entry").ToList();
                int rowspan = 1;
                int colspan = 1;

                cells = cells.OrderBy(fd => fd.Attributes["namest"] != null ? int.Parse(fd.Attributes["namest"].Value) : 0).ToList();
                rowIdx++;
                int colIdx = 1;
                foreach (HtmlNode cell in cells)
                {
                    int rowSpanCalculated = 0;
                    HtmlNode td = tabFigDoc.CreateElement("td");

                    int colName = cell.Attributes["namest"] != null ? (int.Parse(cell.Attributes["namest"].Value)) : colIdx;
                    int nameend = cell.Attributes["nameend"] != null ? (int.Parse(cell.Attributes["nameend"].Value) + 1) : colName + 1;

                    if (cell.Attributes["spanname"] != null)
                    {
                        MatchCollection spanTo = Regex.Matches(cell.Attributes["spanname"].Value, @"\d{1,}");
                        if (spanTo.Count == 2)
                        {
                            colName = int.Parse(spanTo[0].Value);
                            nameend = int.Parse(spanTo[1].Value) + 1;
                        }
                    }

                    colspan = (nameend - colName);

                    bool calrowspan = cell.Attributes["calrowspan"] != null;
                    td.Attributes.Add("colspan", colspan.ToString());
                    td.Attributes.Add("namest", colName.ToString());
                    td.Attributes.Add("nameend", (nameend + 1).ToString());
                    rowspan = cell.Attributes["morerows"] != null ? int.Parse(cell.Attributes["morerows"].Value) + rowpanBase : 1;
                    if (!useMorerows || calrowspan)
                    {
                        string colnbr = cell.Attributes["colname"].Value;
                        HtmlNode rowNextSibling = row.NextSibling;
                        int i = 0;
                        while (rowNextSibling != null)
                        {

                            rowSpanCalculated = rowNextSibling.NodeType.ToString() == HtmlNodeType.Element.ToString() ? rowSpanCalculated + 1 : rowSpanCalculated;
                            HtmlNodeCollection entries = rowNextSibling.SelectNodes(".//entry");
                            if (entries != null)
                            {
                                HtmlNode testNode = entries.FirstOrDefault(fd => fd.Attributes["cspase"].Value.Contains(colnbr));
                                if (testNode != null)
                                {
                                    rowspan = rowSpanCalculated;
                                    break;
                                }
                                else if (i == 0 && rowNextSibling == rowNextSibling.ParentNode.ChildNodes.Last())
                                {
                                    rowspan = rowSpanCalculated + 1;
                                }
                            }

                            i++;
                            rowNextSibling = rowNextSibling.NextSibling;
                        }
                        td.Attributes.Add("colname", colnbr);
                    }
                    else
                    {
                        td.Attributes.Add("colname", colIdx.ToString());
                    }

                    string morerows = rowspan.ToString();
                    cell.Attributes.Add("rowspan", morerows);
                    td.Attributes.Add("rowspan", morerows);

                    td.InnerHtml = cell.InnerHtml.Trim();
                    tr.AppendChild(td);
                    colIdx++;
                }



                tbody.AppendChild(tr);
            }
            table.Attributes.Add("class", "sgml_table");
            table.AppendChild(thead);

            table.AppendChild(tbody);

            string html = table.OuterHtml;
            table.Attributes.Add("group", "ALL");
            if (t_title != null)
            {

                table.Attributes["group"].Value = t_title.Attributes["group"] != null ? t_title.Attributes["group"].Value : "ALL";
                html = "<h1>" + t_title.InnerHtml + "</h1>" + table.OuterHtml;
            }
            html = tfoot.InnerHtml.Trim() != "" ? html + "<hr/>" + tfoot.OuterHtml : html;
            html = "<div>" + html + "</div>";

            if (saveloc != null)
            {
                using (StreamWriter sw = new StreamWriter(saveloc, false))
                {
                    sw.WriteLine(html);
                    sw.Close();
                    sw.Dispose();
                }
            }
            GC.Collect();


            return html;

        }

        private NameValueCollection sortVariableNames(NameValueCollection variableGroups)
        {
            backgroundWorker1.ReportProgress(7, "");
            NameValueCollection sortedvariableGroups = new NameValueCollection();
            NameValueCollection VariblePreFix = new NameValueCollection();
            Hashtable groupLookUp = new Hashtable();
            List<Object> SBGroups = new List<object>();
            List<Object> sortedExpandedEffectivityINTs = new List<object>();
            List<int> sbGroups = new List<int>();
            foreach (string groupName in variableGroups.Keys)
            {
                groupLookUp = new Hashtable();
                if (groupName == "") { continue; }
                string[] grpVariables = variableGroups.GetValues(groupName);
                int groupX = 0;
                if (int.TryParse(groupName, out groupX))
                {
                    if (!sbGroups.Contains(groupX))
                    {
                        sbGroups.Add(groupX);
                    }
                }
                sortedExpandedEffectivityINTs = new List<object>();
                string variblePrefix = grpVariables[0].Trim().Substring(0, 2);
                VariblePreFix = new NameValueCollection();

                foreach (string entry in grpVariables)
                {
                    variblePrefix = entry.Trim().Substring(0, 2);
                    string name = entry.Trim();
                    string vGroup = groupName.Trim();

                    if (SBGroups.Contains(int.Parse(vGroup.Trim())) == false)
                    {
                        SBGroups.Add(int.Parse(vGroup.Trim()));
                    }

                    if (name.Contains("-") || name.Contains("–"))
                    {
                        List<Object> expandedEffectivityINTs = expandEffectivity(name.Replace(variblePrefix, ""));

                        if (expandedEffectivityINTs != null)
                        {
                            expandedEffectivityINTs.Sort();
                            foreach (int Nbr in expandedEffectivityINTs)
                            {
                                if (!sortedExpandedEffectivityINTs.Contains(Nbr))
                                {
                                    sortedExpandedEffectivityINTs.Add(Nbr);
                                    string VariableNBR = variblePrefix + Nbr.ToString().PadLeft(3, '0');                                    
                                    VariblePreFix.Add(variblePrefix, VariableNBR);
                                    groupLookUp.Add(VariableNBR, vGroup);
                                }
                            }
                        }
                    }
                    else
                    {
                        int Nbr = int.Parse(name.Replace(variblePrefix, ""));
                        string VariableNBR = variblePrefix + Nbr.ToString().PadLeft(3, '0');                        
                        if (!groupLookUp.Contains(VariableNBR))
                        {
                            VariblePreFix.Add(variblePrefix, VariableNBR);
                            sortedExpandedEffectivityINTs.Add(Nbr);
                            groupLookUp.Add(VariableNBR, vGroup);
                        }
                    }
                }
                sbGroups.Sort();
                sortedExpandedEffectivityINTs.Sort();
                foreach (string vprefix in VariblePreFix.Keys)
                {
                    List<string> tails = VariblePreFix.GetValues(vprefix).ToList();
                    tails.Sort();
                    foreach (string tail in tails)
                    {
                        string VariableNBR = tail;
                        if (groupLookUp.ContainsKey(VariableNBR))
                        {
                            string VariableGroup = groupLookUp[VariableNBR].ToString();
                            sortedvariableGroups.Add(VariableGroup, VariableNBR);
                        }
                    }
                }
            }

            #region Groups.xml
            System.Data.DataTable ds = new System.Data.DataTable();

            
            variableGroups = new NameValueCollection();

            foreach (int Group in SBGroups)
            {
                string colName = "Group" + Group.ToString().PadLeft(3, '0');
                ds.Columns.Add(colName);
                StringBuilder GroupNbrs = new StringBuilder();

                int rowIdx = 0;



                foreach (string Nbr in sortedvariableGroups.GetValues(Group.ToString()))
                {
                    variableGroups.Add(Group.ToString(), Nbr);
                    string FDXTail = Nbr;
                    GroupNbrs.Append(FDXTail + ", ");                    
                    if (ds.Rows.Count <= rowIdx)
                    {
                        ds.Rows.Add(new string[] { "" });
                    }
                    ds.Rows[rowIdx++].SetField(ds.Columns[colName], FDXTail);
                }
            }            
            string path = activeSb_dir + "\\GROUPS\\";
            Directory.CreateDirectory(path);
            path = path + "Groups.xml";
            ds.TableName = "Groups";
            ds.WriteXml(path, XmlWriteMode.WriteSchema);
            #endregion

            return variableGroups;
        }

        #endregion   
    }
   

}

