//#define counting
using System;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Data;
using ICSharpCode.SharpZipLib.Zip;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Xml.Xsl;

namespace ZAGS_KLADR_OKATO
{
    class Program
    {
        // базовый каталог
        public static string strBaseDir = Environment.CurrentDirectory;
        // временный каталог
        public static string strTempDir = "Temp";
        // приставки населенных пунктов
        public static string[,] strAffix = new string[,] {  { " \\(РАЙОН\\)", " район" }, 
                                                            { " \\(Р\\-Н\\)", " район" }, 
                                                            { " \\(С\\.?\\)", "село" }, 
                                                            { " \\(Д\\.?\\)", "деревня" }, 
                                                            { " \\(П\\.Г\\.Т\\.\\)", "пгт" }, 
                                                            { " \\(ПГТ\\)", "пгт" }, 
                                                            { " \\(Р\\.?П\\.?\\)", "пгт" }, 
                                                            { " \\(СОВХОЗ\\)", "посёлок совхоза" }, 
                                                            { " \\(С/Х\\)", "посёлок совхоза" }, 
                                                            { " \\(ПОС\\. С/З\\)", "посёлок совхоза" }, 
                                                            { " \\(С/З\\)", "посёлок совхоза" }, 
                                                            { " \\(СТ\\.\\)", "посёлок железнодорожн" }, 
                                                            { " \\(ПОС\\. Ж\\.Д\\.СТ\\.\\)", "посёлок железнодорожн" }, 
                                                            { " \\(ПОС\\.?\\)", "посёлок" },
                                                            { " \\(П\\.?\\)", "посёлок" } };
        // поселки городского подчинения
        public static string[] strExclusion = new string[] {
            "посёлок Кармалка",
            "деревня Алань",
            "деревня Дмитриевка",
            "деревня Ильинка",
            "посёлок Ерыклинский",
            "" };
        // сегодня
        public static string strToday = "_" + DateTime.Today.ToShortDateString();

        static void Main(string[] args)
        {
            string rl = "";
            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "0":
                    case "1":
                    case "2":
                    case "3":
                    case "4":
                    case "s":
                        rl = args[0];
                        break;
                    default:
                        Console.Write("0 - все, 1 - разводы, 2 - браки, 3 - рожденные, 4 - смерти, s - new: ");
                        rl = Console.ReadLine();
                        break;
                }
            }
            else
            {
                Console.Write("0 - все, 1 - разводы, 2 - браки, 3 - рожденные, 4 - смерти, s - new: ");
                rl = Console.ReadLine();
            }
            Console.WriteLine(DateTime.Now);
            if (rl == "0" || rl == "1")
            {
                MrrgDvrc("Divorce");
                Console.WriteLine(DateTime.Now + " Divorces done!");
            }
            if (rl == "0" || rl == "2")
            {
                MrrgDvrc("Marriage");
                Console.WriteLine(DateTime.Now + " Marriages done!");
            }
            if (rl == "0" || rl == "3")
            {
                MrrgDvrc("Birth");
                Console.WriteLine(DateTime.Now + " Borns done!");
            }
            if (rl == "0" || rl == "4")
            {
                MrrgDvrc("Death");
                Console.WriteLine(DateTime.Now + " Deaths done!");
            }
            Console.Write("Press 'Enter'!"); Console.ReadLine();
        }

        public static void MrrgDvrc(string strAct)
        {
            Regex rgxKLADR = new Regex(@"\d{13}", RegexOptions.IgnoreCase);
            Regex rgxRegion = new Regex(@"\d{2}0{11}", RegexOptions.IgnoreCase);
            Regex rgxZAGS = new Regex(@"\""\d{13}\""", RegexOptions.IgnoreCase);
            try
            {
                OleDbConnection connDBF = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strBaseDir + ";Extended Properties=DBASE IV;Persist Security Info=False;");
                connDBF.Open();
                if (!Directory.Exists(Path.Combine(strBaseDir, strTempDir)))
                    Directory.CreateDirectory(Path.Combine(strBaseDir, strTempDir));
                if (!Directory.Exists(Path.Combine(strBaseDir, strAct + strToday)))
                    Directory.CreateDirectory(Path.Combine(strBaseDir, strAct + strToday));
                // распаковка
#if !counting
                string sZip = Directory.GetFiles(strBaseDir, "fsgs_" + strAct + "*.zip")[0];
                UncompressZip(sZip);
#endif

                #region // создаем xml (чтение файлов из временной папки)
#if !counting
                string[] FileEntr = Directory.GetFiles(Path.Combine(strBaseDir, strTempDir));
                foreach (string fPath in FileEntr)
                {// цикл, пока файлы есть

                    FileStream fsRead = new FileStream(fPath, FileMode.Open);
                    string fName = Path.GetFileName(fPath);
                    // избавляемся от нечитаемых символов
                    string[] splitArr = fName.Split('_');
                    if (splitArr.Length > 7)
                    {
                        fName = "";
                        for (int i = 0; i < 8; i++)
                            fName += splitArr[i] + "_";
                        fName += ".xml";
                    }
                    FileStream fsWrite = new FileStream(Path.Combine(strBaseDir + "\\" + strAct + strToday, fName), FileMode.Create);

                    using (StreamReader srXML = new StreamReader(fsRead))
                    {
                        using (StreamWriter swXML = new StreamWriter(fsWrite))
                        {
                            string line;
                            string newLine;
                            string strOKATO = "";
                            string strKLADR = "";
                            while ((line = srXML.ReadLine()) != null)
                            {
                                newLine = line;
                                Match mKLADR = rgxKLADR.Match(line);
                                while (mKLADR.Success)
                                {
                                    strOKATO = OKATO(connDBF, mKLADR.Value);
                                    newLine = rgxKLADR.Replace(line, strOKATO);
                                    if (!rgxZAGS.IsMatch(line)) // если ЗАГС, то нужен ОКАТО, а не ТЕРСОН
                                    {
                                        strKLADR = mKLADR.Value;
                                        if (rgxRegion.IsMatch(line)) // если это субъект РФ, то нужны 2 первых цифры
                                            newLine = newLine.Replace(strOKATO, strOKATO.Substring(0, 2));
                                        else // если это адрес, то по КЛАДР узнаем ИМЯ (ОКАТО уже узнали), потом узнаем ТЕРСОН
                                            newLine = rgxKLADR.Replace(line, TERSON(connDBF, strOKATO, N_A_M_E(connDBF, strKLADR)));
                                    }
                                    mKLADR = mKLADR.NextMatch();
                                }
                                swXML.WriteLine(newLine);
                            }
                        }
                    }
                    fsWrite.Close();
                    fsRead.Close();
                }
#endif
                #endregion

                #region // заполняем txt (чтение файлов из папки с xml)

#if counting
                ////string[] FileEntr = Directory.GetFiles(strBaseDir + "\\" + strAct + strToday + "\\");
                string[] FileEntr = Directory.GetFiles(strBaseDir + "\\" + "fsgs_" + strAct + "_март\\");
#else
                FileEntr = Directory.GetFiles(Path.Combine(strBaseDir, strAct + strToday));
#endif
                string txtFileName = Path.Combine(strBaseDir + "\\" + strAct + strToday, strAct + strToday + ".txt");

                //***********************************
                string listPath = Path.Combine(strBaseDir, strAct + strToday + ".xml");
                if (File.Exists(listPath)) File.Delete(listPath);
                StreamWriter listXML = new StreamWriter(listPath, true, Encoding.GetEncoding("UTF-8"));
                listXML.WriteLine("<?xml version='1.0'?>");
                listXML.WriteLine("<report>");

                foreach (string fPath in FileEntr)// цикл, пока файлы есть
                    if (fPath != txtFileName)
                    {
#if counting
                        XmlDocument xB = new XmlDocument();
                        xB.Load(fPath);
                        XmlNode father;
                        XmlNode fatherCtznshp;
                        XmlNode motherCtznshp;
                        XmlNode root = xB.DocumentElement;
                        father = root.SelectSingleNode("descendant::row[@code='56']/col");
                        fatherCtznshp = root.SelectSingleNode("descendant::row[@code='35']/col");
                        motherCtznshp = root.SelectSingleNode("descendant::row[@code='45']/col");
                        string fthr;
                        string fthctzn;
                        string mthctzn;
                        fthr = father.LastChild == null ? "-1" : father.LastChild.InnerText.ToString();
                        fthctzn = fatherCtznshp.LastChild == null ? "-1" : fatherCtznshp.LastChild.InnerText.ToString();
                        mthctzn = motherCtznshp.LastChild == null ? "-1" : motherCtznshp.LastChild.InnerText.ToString();
                        listXML.WriteLine("<p filename='" + fPath + "' father='" + fthr + "' fatherCtznshp='" + fthctzn + "' motherCtznshp='" + mthctzn + "'/>");
#else
                        listXML.WriteLine("<p filename='" + fPath + "' act='" + FileEntr.ToString() + "'/>");
#endif
                    }
                listXML.WriteLine("</report>");
                listXML.Close();
                XsltSettings settings = new XsltSettings(true, true);
                XmlUrlResolver resolver = new XmlUrlResolver();
                resolver.Credentials = System.Net.CredentialCache.DefaultCredentials;
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(strAct + ".xsl", settings, resolver);
                xslt.Transform(listPath, strAct + strToday + ".txt");
                //**************************************
                #region
                //if (File.Exists(txtFileName)) File.Delete(txtFileName);
                //StreamWriter swTXT = new StreamWriter(txtFileName, true, Encoding.GetEncoding("windows-1251"));
                //int Counter = 1;
                //foreach (string fPath in FileEntr)
                //{// цикл, пока файлы есть
                //    if (fPath != txtFileName)
                //    {
                //        //***********************************
                //        //listXML.WriteLine("<p filename='" + Path.GetFileName(fPath) + "'/>");
                //        //listXML.WriteLine("<p filename='" + fPath + "'/>");
                //        //***********************************

                //        FileStream fsRead = new FileStream(fPath, FileMode.Open);
                //        //Console.WriteLine(fPath);
                //        //string fName = DC(Path.GetFileName(fPath), "UTF-8", "iso-8859-5");

                //        XmlReader ZAGSreader = XmlReader.Create(fsRead);
                //        while (ZAGSreader.Read())
                //        {
                //            switch (ZAGSreader.LocalName)
                //            {
                //                case "item":
                //                    if (ZAGSreader.GetAttribute("name") == "okato")
                //                    {
                //                        txtOZ = ZAGSreader.GetAttribute("value"); // ОКАТО ЗАГС
                //                        txtLine[1] = txtOZ.Length > 7 ? txtOZ.Substring(0, 8) : txtOZ;
                //                        txtBorn[1] = txtOZ.Length > 7 ? txtOZ.Substring(0, 8) : txtOZ;
                //                        txtRZ = txtOZ.Length > 4 ? txtOZ.Substring((txtOZ.Substring(2, 3) == "401" ? 5 : 2), 3) : txtOZ; // район ЗАГС
                //                        if (strAct == "Divorce")
                //                        {
                //                            txtLine[32] = txtOZ;
                //                            txtLine[33] = txtRZ;
                //                        }
                //                        else if (strAct == "Marriage")
                //                        {
                //                            txtLine[30] = txtOZ;
                //                            txtLine[31] = txtRZ;
                //                        }
                //                        else if (strAct == "Death")
                //                        {
                //                            txtLine[18] = txtOZ;
                //                            txtLine[19] = txtRZ;
                //                        }
                //                        else if (strAct == "Birth")
                //                        {
                //                            txtBorn[59] = txtOZ;
                //                            txtBorn[61] = txtRZ;
                //                        }
                //                        else // s
                //                        {
                //                            txtLine[18] = txtOZ;
                //                            txtLine[19] = txtRZ;
                //                        }
                //                    }
                //                    else if (ZAGSreader.GetAttribute("name") == "akts")
                //                    {
                //                        txtLine[3] = ZAGSreader.GetAttribute("value");
                //                        txtBorn[3] = ZAGSreader.GetAttribute("value");
                //                    }
                //                    else if (ZAGSreader.GetAttribute("name") == "name")
                //                        txtName = ZAGSreader.GetAttribute("value");
                //                    break;
                //                case "row":
                //                    {
                //                        if (strAct == "Marriage")
                //                        {
                //                            #region // Marriage
                //                            switch (ZAGSreader.GetAttribute("code"))
                //                            {
                //                                case "1": // месяц регистрации
                //                                    ZAGSreader.Read(); txtLine[5] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "2": // год регистрации
                //                                    ZAGSreader.Read(); txtLine[6] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                //**********************************************************
                //                                case "3": //дата рождения мужа
                //                                    ZAGSreader.Read(); txtLine[7] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "4": // месяц рождения мужа
                //                                    ZAGSreader.Read(); txtLine[8] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "5": // год рождения мужа
                //                                    ZAGSreader.Read(); txtLine[9] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "38": // возраст мужа
                //                                    ZAGSreader.Read(); txtLine[10] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "9": // код субъекта РФ мужа
                //                                    ZAGSreader.Read(); txtLine[11] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "6": // гражданство мужа
                //                                    ZAGSreader.Read(); txtLine[12] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "7": // код гражданство мужа
                //                                    ZAGSreader.Read(); txtLine[13] = ZAGSreader.ReadElementString();
                //                                    if (txtLine[12] == "" | txtLine[12] == null) txtLine[12] = txtLine[13]; break;
                //                                case "8": // субъект РФ мужа
                //                                    //ZAGSreader.Read(); txtLine[..] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "10": // ТЕРСОН мужа
                //                                    ZAGSreader.Read(); txtLine[14] = ZAGSreader.ReadElementString();
                //                                    txtLine[15] = Name2TERSON(conn2, txtLine[14]); break; // ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
                //                                    //txtLine[15] = txtLine[14]; break; 
                //                                case "11": // код семейного положения мужа
                //                                    ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    txtLine[17] = txtLine[16];
                //                                    break;
                //                                //*************************************************************
                //                                case "12": // дата рождения жены
                //                                    ZAGSreader.Read(); txtLine[18] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "13": // месяц рождения жены
                //                                    ZAGSreader.Read(); txtLine[19] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "14": // год рождения жены
                //                                    ZAGSreader.Read(); txtLine[20] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "39": // возраст жены
                //                                    ZAGSreader.Read(); txtLine[21] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "18": // код субъекта РФ жены
                //                                    ZAGSreader.Read(); txtLine[22] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "15": // гражданство жены
                //                                    ZAGSreader.Read(); txtLine[23] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "16": // код гражданство жены
                //                                    ZAGSreader.Read(); txtLine[24] = ZAGSreader.ReadElementString();
                //                                    txtLine[23] = txtLine[24]; break;
                //                                case "17": // субъект РФ жены
                //                                    //ZAGSreader.Read(); txtLine[..] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "19": // ТЕРСОН жены
                //                                    ZAGSreader.Read(); txtLine[25] = ZAGSreader.ReadElementString();
                //                                    txtLine[26] = Name2TERSON(conn2, txtLine[25]); break; // ~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~
                //                                    //txtLine[26] = txtLine[25]; break;
                //                                case "20": // 
                //                                    //ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    //txtLine[17] = txtLine[16];
                //                                    break;
                //                                case "21": // 
                //                                    //ZAGSreader.Read(); txtLine[14] = ZAGSreader.ReadElementString();
                //                                    //txtLine[15] = txtLine[14];
                //                                    break;
                //                                case "22": // 
                //                                    //ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    //txtLine[17] = txtLine[16];
                //                                    break;
                //                                case "23": // код семейного положения жены
                //                                    ZAGSreader.Read(); txtLine[28] = ZAGSreader.ReadElementString();
                //                                    txtLine[27] = txtLine[28];
                //                                    break;
                //                                case "24": // количество общих детей
                //                                    ZAGSreader.Read(); txtLine[29] = ZAGSreader.ReadElementString();
                //                                    break;
                //                            }
                //                            #endregion
                //                        }
                //                        else if (strAct == "Divorce")
                //                        {
                //                            #region // divorce
                //                            switch (ZAGSreader.GetAttribute("code"))
                //                            {
                //                                case "1": // месяц регистрации
                //                                    ZAGSreader.Read(); txtLine[5] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "2": // год регистрации
                //                                    ZAGSreader.Read(); txtLine[6] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                //**********************************************************
                //                                case "3": // дата рождения мужа
                //                                    ZAGSreader.Read(); txtLine[7] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "4": // месяц рождения мужа
                //                                    ZAGSreader.Read(); txtLine[8] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "5": // год рождения мужа
                //                                    ZAGSreader.Read(); txtLine[9] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "38": // возраст мужа
                //                                    ZAGSreader.Read(); txtLine[10] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "9": // код субъекта РФ мужа
                //                                    ZAGSreader.Read(); txtLine[11] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "6": // гражданство мужа
                //                                    ZAGSreader.Read(); txtLine[12] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "7": // код гражданство мужа
                //                                    ZAGSreader.Read(); txtLine[13] = ZAGSreader.ReadElementString();
                //                                    if (txtLine[12] == "" | txtLine[12] == null) txtLine[12] = txtLine[13]; break;
                //                                case "8": // субъект РФ мужа
                //                                    //ZAGSreader.Read(); txtLine[..] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "10": // ТЕРСОН мужа
                //                                    ZAGSreader.Read(); txtLine[14] = ZAGSreader.ReadElementString();
                //                                    txtLine[15] = Name2TERSON(conn2, txtLine[14]); break;
                //                                case "11": // код семейного положения мужа
                //                                    //ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    //txtLine[17] = txtLine[16];
                //                                    break;
                //                                //*************************************************************
                //                                case "12": // дата рождения жены
                //                                    ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "13": // месяц рождения жены
                //                                    ZAGSreader.Read(); txtLine[17] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "14": // год рождения жены
                //                                    ZAGSreader.Read(); txtLine[18] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "39": // возраст жены
                //                                    ZAGSreader.Read(); txtLine[19] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "18": // код субъекта РФ жены
                //                                    ZAGSreader.Read(); txtLine[20] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "15": // гражданство жены
                //                                    ZAGSreader.Read(); txtLine[21] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "16": // код гражданство жены
                //                                    ZAGSreader.Read(); txtLine[22] = ZAGSreader.ReadElementString();
                //                                    txtLine[21] = txtLine[22]; break;
                //                                case "17": // субъект РФ жены
                //                                    //ZAGSreader.Read(); txtLine[..] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "19": // ТЕРСОН жены
                //                                    ZAGSreader.Read(); txtLine[23] = ZAGSreader.ReadElementString();
                //                                    txtLine[24] = Name2TERSON(conn2, txtLine[23]); break;
                //                                case "21": // месяц регистрации брака
                //                                    ZAGSreader.Read(); txtLine[26] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "22": // год регистрации брака
                //                                    ZAGSreader.Read(); txtLine[27] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "23": // число прекращения брака
                //                                    ZAGSreader.Read(); txtLine[28] = ZAGSreader.ReadElementString();
                //                                    //txtLine[27] = txtLine[28];
                //                                    break;
                //                                case "24": // месяц прекращения брака
                //                                    ZAGSreader.Read(); txtLine[29] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "25": // год прекращения брака
                //                                    ZAGSreader.Read(); txtLine[30] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "20": // количество общих детей
                //                                    ZAGSreader.Read(); txtLine[31] = ZAGSreader.ReadElementString();
                //                                    //txtLine[17] = txtLine[16];
                //                                    break;
                //                            }
                //                            #endregion
                //                        }
                //                        else if (strAct == "Birth")
                //                        {
                //                            #region // Birth
                //                            switch (ZAGSreader.GetAttribute("code"))
                //                            {
                //                                case "1": // месяц регистрации
                //                                    ZAGSreader.Read(); txtBorn[4] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "2": // год регистрации
                //                                    ZAGSreader.Read(); txtBorn[5] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "6": // пол
                //                                    ZAGSreader.Read(); txtBorn[6] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "3": // день рождения
                //                                    ZAGSreader.Read(); txtBorn[8] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "4": // месяц рождения
                //                                    ZAGSreader.Read(); txtBorn[9] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "5": // год рождения
                //                                    ZAGSreader.Read(); txtBorn[10] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "11": // кол-во родившихся детей
                //                                    ZAGSreader.Read(); txtBorn[12] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "12": // жив/мертв
                //                                    ZAGSreader.Read(); txtBorn[13] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "24": // основное заболевание новорожденного
                //                                    ZAGSreader.Read(); txtBorn[15] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "25": // другое заболевание новорожденного
                //                                    ZAGSreader.Read(); txtBorn[17] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "26": // основное заболевание матери
                //                                    ZAGSreader.Read(); txtBorn[19] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "27": // другое заболевание матери
                //                                    ZAGSreader.Read(); txtBorn[21] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "28": // другие обстоятельства
                //                                    ZAGSreader.Read(); txtBorn[23] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "29": // ребенок по счету
                //                                    ZAGSreader.Read(); txtBorn[25] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "32": // день рождения отца
                //                                    ZAGSreader.Read(); txtBorn[27] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "33": // месяц рождения отца
                //                                    ZAGSreader.Read(); txtBorn[28] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "34": // год рождения отца
                //                                    ZAGSreader.Read(); txtBorn[29] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "35": // код гражданства отца
                //                                    ZAGSreader.Read(); txtBorn[31] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "37": // субъект место жительства отца
                //                                    ZAGSreader.Read(); txtBorn[32] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "38": // место жительства отца + код места жительства отца
                //                                    ZAGSreader.Read(); txtBorn[33] = ZAGSreader.ReadElementString();
                //                                    txtBorn[34] = Name2TERSON(conn2, txtBorn[33]);
                //                                    break;
                //                                case "42": // день рождения матери
                //                                    ZAGSreader.Read(); txtBorn[36] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "43": // месяц рождения матери
                //                                    ZAGSreader.Read(); txtBorn[37] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "44": // год рождения матери
                //                                    ZAGSreader.Read(); txtBorn[38] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "45": // код гражданства матери
                //                                    ZAGSreader.Read(); txtBorn[40] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "52": // субъект места жительства матери
                //                                    ZAGSreader.Read(); txtBorn[41] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "53": // место жительства матери + код места жительства матери
                //                                    ZAGSreader.Read(); txtBorn[42] = ZAGSreader.ReadElementString();
                //                                    txtBorn[43] = Name2TERSON(conn2, txtBorn[42]);
                //                                    break;
                //                                case "47": // код образования матери
                //                                    ZAGSreader.Read(); txtBorn[44] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "48": // код занятости матери
                //                                    ZAGSreader.Read(); txtBorn[46] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "56": // сведения об отце указаны
                //                                    ZAGSreader.Read(); txtBorn[47] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "14": // заболевания матери
                //                                    ZAGSreader.Read(); txtBorn[48] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "15": // заболевания матери (кратк)
                //                                    ZAGSreader.Read(); txtBorn[49] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "21": // перинатальная причина
                //                                    ZAGSreader.Read(); txtBorn[50] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "22": // перинатальная причина (кратк)
                //                                    ZAGSreader.Read(); txtBorn[51] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "30": // вес ребенка
                //                                    ZAGSreader.Read(); txtBorn[52] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "61": // месяц регистрации брака
                //                                    ZAGSreader.Read(); txtBorn[55] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "62": // год регистрации брака
                //                                    ZAGSreader.Read(); txtBorn[56] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "70": // возраст отца (расчетный)
                //                                    ZAGSreader.Read(); txtBorn[57] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "71": // возраст матери (расчетный)
                //                                    ZAGSreader.Read(); txtBorn[58] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "57": // статус
                //                                    ZAGSreader.Read(); txtBorn[60] = ZAGSreader.ReadElementString();
                //                                    break;
                //                            }
                //                            #endregion
                //                        }
                //                        else if (strAct == "Death")
                //                        {
                //                            #region // Death
                //                            switch (ZAGSreader.GetAttribute("code"))
                //                            {
                //                                case "1": // 1
                //                                    ZAGSreader.Read(); txtLine[3] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "2": // 2
                //                                    ZAGSreader.Read(); txtLine[4] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "3": // 3
                //                                    ZAGSreader.Read(); txtLine[5] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "4": // 4
                //                                    ZAGSreader.Read(); txtLine[6] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "5": // 5
                //                                    ZAGSreader.Read(); txtLine[7] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "6": // 6
                //                                    ZAGSreader.Read(); txtLine[8] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "11": // 11
                //                                    ZAGSreader.Read(); txtLine[9] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "12": // 12
                //                                    ZAGSreader.Read(); txtLine[10] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "19": // 19
                //                                    ZAGSreader.Read(); txtLine[11] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "20": // 20
                //                                    ZAGSreader.Read(); txtLine[12] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "21": // 21
                //                                    ZAGSreader.Read(); txtLine[13] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "72": // 72
                //                                    ZAGSreader.Read(); txtLine[14] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "75": // 75
                //                                    ZAGSreader.Read(); txtLine[15] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "76": // 76
                //                                    ZAGSreader.Read(); txtLine[16] = ZAGSreader.ReadElementString();
                //                                    break;
                //                                case "77": // 77
                //                                    ZAGSreader.Read(); txtLine[17] = ZAGSreader.ReadElementString();
                //                                    break;
                //                            }
                //                            #endregion
                //                        }
                //                    }
                //                    break;
                //            }
                //        }

                //        //using (swTXT)
                //        //{
                //        string txt = txtName + "#0#" + txtLine[1] + "#" + Counter.ToString();
                //        string txtB = txtName + "#0#" + txtBorn[1] + "#" + Counter.ToString();
                //        //string txt = "0#" + txtLine[1] + "#" + Counter.ToString();
                //        switch (strAct)
                //        {
                //            case "Marriage":
                //            case "Death":
                //            case "Divorce":
                //                for (int i = 3; i < (txtLine.Length - (strAct == "Divorce" ? 0 : 2)); i++)
                //                {
                //                    txt += "#" + txtLine[i];
                //                    txtLine[i] = "";
                //                }
                //                break;
                //            case "Birth":
                //                for (int i = 3; i < txtBorn.Length; i++)
                //                {
                //                    txtB += "#" + txtBorn[i];
                //                    txtBorn[i] = "";
                //                }
                //                txt = txtB;
                //                break;
                //        }
                //        swTXT.WriteLine(txt);
                //        Counter++;
                //        //}
                //    }
                //}
                //swTXT.Close();
                //Console.ReadLine();
                #endregion
                if (Directory.Exists(strBaseDir + "\\Temp\\"))
                    Directory.Delete(strBaseDir + "\\Temp\\", true);
                #endregion
            }
            catch (OleDbException exception)
            {
                for (int i = 0; i < exception.Errors.Count; i++)
                    Console.WriteLine("Index #" + i + "\n" +
                     "Message: " + exception.Errors[i].Message + "\n" +
                     "Native: " + exception.Errors[i].NativeError.ToString() + "\n" +
                     "Source: " + exception.Errors[i].Source + "\n" +
                     "SQL: " + exception.Errors[i].SQLState + "\n");
                Console.ReadLine();
            }

        }

        public static string Name2TERSON(OleDbConnection connKLADR, string strNAME)
        {
            string strSQL="";
            Regex rgxGorod = new Regex(@" \(г\.?\)$", RegexOptions.IgnoreCase);
            Regex rgxRayon = new Regex(@" (\(район\))|(\(р-н\))$", RegexOptions.IgnoreCase);
            string[] strN;
            strN = strNAME.Split(new string[] {", "}, StringSplitOptions.None);
            if (rgxGorod.IsMatch(strN[0])) // если ГОРОД
            {
                strN[0] = "г. " + strN[0].Replace(" (Г.)", "").Replace(" (г)", "").Trim();
                if (strN.GetLength(0) > 1) // если КАЗАНЬ(0) с районом(1)
                {
                    strN[1] = Up2LowNames(rgxRayon.Replace(strN[1], "район").Trim());
                    // TERSON до 2015
                    // strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + Up2LowNames(strN[0]) + "' AND NAME='" + strN[1] + "'";
                    // Terson 2015
                    strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + Up2LowNames(strN[0]) + "' AND NAME='" + strN[1] + " " + Up2LowNames(strN[0]) + "'";
                }
                else // если город(0) без района
                {
                    strSQL = "SELECT CODE FROM 92_TER WHERE NAME='" + Up2LowNames(strN[0].Trim()) + "'";
                }
            }
            else if (rgxRayon.IsMatch(strN[0])) // если РАЙОН(0)
            {
                strN[0] = Up2LowNames(rgxRayon.Replace(strN[0], "").Trim());
                if (strN.Length < 2) return strNAME; // если РАЙОН без запятой
                if (rgxGorod.IsMatch(strN[1])) // если РАЙОН(0) + ГОРОД(1)
                {
                    strN[1] = "г. " + Up2LowNames(strN[1].Replace(" (Г.)", "").Replace(" (г)", "").Trim());
                    // TERSON до 2015
                    // strSQL = "SELECT CODE FROM 92_TER WHERE Left(CODE, 2)='92' AND NAME='" + strN[1] + "'";
                    // TERSON 2015
                    strSQL = "SELECT CODE FROM 92_TER WHERE Left(CODE, 2)='92' AND NAME='" + strN[1] + " " + strN[0] + " район'";
                }
                else // если РАЙОН(0) + село(1)
                {
                    strN[1] = Up2LowNames(ReformNames(strN[1]).Trim());
                    // TERSON до 2015
                    // strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + strN[0] + "' AND NAME='" + strN[1] + " " + "'";
                    // TERSON 2015
                    strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + strN[0] + "' AND NAME='" + strN[1] + " " + strN[0] + " район'";
                }
            }
            else // если что-то непонятное
                return strNAME;

            OleDbCommand comm = connKLADR.CreateCommand();
            comm.CommandText = strSQL;
            comm.CommandType = CommandType.Text;
            OleDbDataReader reader = comm.ExecuteReader();
            string strN2T = "";
            if (reader.HasRows) //true
            {
                reader.Read();
                strN2T = reader.GetValue(0).ToString();
            }
            reader.Close();
            return strN2T;
        }

        /// <summary>замена приставки из массива strAffix</summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public static string ReformNames(string strName)
        {
            for (int i = 0; i < strAffix.GetLength(0); i++)
            {
                Regex rgxAffix = new Regex(@"" + strAffix[i, 0], RegexOptions.IgnoreCase);
                if (rgxAffix.IsMatch(strName))
                {
                    strName = strAffix[i, 1] + " " + rgxAffix.Replace(strName, "");
                    return strName;
                }
            }
            return strName;
        }

        public static string Up2LowNames(string strName) // меняет регистр букв, если "ВСЕ ЗАГЛАВНЫЕ", на "Первые Заглавные, Остальные Строчные" и убирает лишние пробелы
        {
            Regex rgxU2L = new Regex (@"[А-Я][А-Я]+", RegexOptions.None);
            Regex rgxSpaces = new Regex(@"\s+");
            MatchCollection matches = rgxU2L.Matches(strName);
            foreach (Match match in matches)
                strName = strName.Replace(match.ToString().Substring(1), match.ToString().Substring(1).ToLower());
            strName = rgxSpaces.Replace(strName, " ");
            return strName;
        }

        public static string OKATO(OleDbConnection connKLADR, string strKLADR)
        {
            OleDbCommand comm = connKLADR.CreateCommand();
            comm.CommandText = "SELECT NEWCODE FROM ALTNAMES WHERE OLDCODE='" + strKLADR + "'";
            comm.CommandType = CommandType.Text;
            OleDbDataReader reader = comm.ExecuteReader();

            OleDbCommand commKLADR = connKLADR.CreateCommand();
            if (reader.HasRows) //true
            {
                reader.Read();
                string strCODE = reader.GetString(0);
                commKLADR.CommandText = "SELECT OCATD FROM KLADR WHERE CODE='" + reader.GetString(0) + "'";
            }
            else
            {
                commKLADR.CommandText = "SELECT OCATD FROM KLADR WHERE CODE='" + strKLADR + "'";
            }
            commKLADR.CommandType = CommandType.Text;
            OleDbDataReader rdKLADR = commKLADR.ExecuteReader();
            rdKLADR.Read();
            string strOKATO;
            if (rdKLADR.HasRows) strOKATO = rdKLADR.GetString(0); else strOKATO = "92000000";
            rdKLADR.Close();
            reader.Close();
            return strOKATO;
        }

        public static string N_A_M_E(OleDbConnection connKLADR, string strKLADR)
        {
            OleDbCommand comm = connKLADR.CreateCommand();
            comm.CommandText = "SELECT NEWCODE FROM ALTNAMES WHERE OLDCODE='" + strKLADR + "'";
            comm.CommandType = CommandType.Text;
            OleDbDataReader reader = comm.ExecuteReader();

            OleDbCommand commKLADR = connKLADR.CreateCommand();
            if (reader.HasRows) //true
            {
                reader.Read();
                string strCODE = reader.GetString(0);
                commKLADR.CommandText = "SELECT NAME FROM KLADR WHERE CODE='" + reader.GetString(0) + "'";
            }
            else
            {
                commKLADR.CommandText = "SELECT NAME FROM KLADR WHERE CODE='" + strKLADR + "'";
            }
            commKLADR.CommandType = CommandType.Text;
            OleDbDataReader rdKLADR = commKLADR.ExecuteReader();
            rdKLADR.Read();
            string strNAME;
            if (rdKLADR.HasRows) strNAME = rdKLADR.GetString(0); else strNAME = "";
            rdKLADR.Close();
            reader.Close();
            return strNAME;
        }

        public static string TERSON(OleDbConnection connKLADR, string strOKATO, string strNAME)
        {
            OleDbCommand comm = connKLADR.CreateCommand();
            comm.CommandText = "SELECT CODE FROM 92_TER WHERE Left(CODE, 11) = '" + strOKATO + "'";
            comm.CommandType = CommandType.Text;
            OleDbDataReader reader = comm.ExecuteReader();
            string strTERSON = "";
            if (reader.HasRows) //true
            {
                reader.Read();
                strTERSON = reader.GetValue(0).ToString();
            }
            reader.Close();
            return strTERSON;
        }

        public static string DC(string strString, string e1, string e2) //попытка смены кодировки имен файлов
        {
            // Create two different encodings.
            Encoding enc1 = Encoding.GetEncoding(e1);
            Encoding enc2 = Encoding.GetEncoding(e2);
            //Encoding enc1 = Encoding.Unicode;

            // Convert the string into a byte[].
            byte[] Bytes2 = enc2.GetBytes(strString);

            // Perform the conversion from one encoding to the other.
            byte[] Bytes1 = Encoding.Convert(enc2, enc1, Bytes2);

            // Convert the new byte[] into a char[] and then into a string.
            // This is a slightly different approach to converting to illustrate
            // the use of GetCharCount/GetChars.
            char[] Chars1 = new char[enc1.GetCharCount(Bytes1, 0, Bytes1.Length)];
            enc1.GetChars(Bytes1, 0, Bytes1.Length, Chars1, 0);
            string String1 = new string(Chars1);
            return String1;
        }

        public static void ValidationHandler(object sender, ValidationEventArgs args)
        {
            Console.WriteLine("Validation error!");
            Console.WriteLine("\tSeverity:{0}", args.Severity);
            Console.WriteLine("\tMessage:{0}", args.Message);
            Log(strBaseDir, "1У", "Validation error!\n\tSeverity:" + args.Severity + "\n\tMessage:" + args.Message);
        }

        private static void ListZipContent(string sFile)
        {
            ZipFile zip = new ZipFile(File.OpenRead(sFile));
            foreach (ZipEntry entry in zip)
                Console.WriteLine(entry.Name);
        }

        private static void UncompressZip(string sFile)
        {
            string sPath = strBaseDir + strTempDir;
            ZipInputStream zipIn = new ZipInputStream(File.OpenRead(sFile));
            ZipEntry entry;

            while ((entry = zipIn.GetNextEntry()) != null)
            {
                FileStream streamWriter = File.Create(@sPath + entry.Name);
                long size = entry.Size;
                byte[] data = new byte[size];
                while (true)
                {
                    size = zipIn.Read(data, 0, data.Length);
                    if (size > 0) streamWriter.Write(data, 0, (int)size);
                    else break;
                }
                streamWriter.Close();
            }
        }

        public static void Log(string logPath, string logFile, string logMessage) // запись протокола
        {
            using (StreamWriter sw = File.AppendText(Path.Combine(logPath, DateTime.Now.Date.ToString("dd_MM_yyyy") + "_protocol.log")))
            {
                sw.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                sw.WriteLine(" :{0}", logFile + " :: " + logMessage);
            }
            Console.WriteLine(logFile + " :: " + logMessage);
        }
    }
}
