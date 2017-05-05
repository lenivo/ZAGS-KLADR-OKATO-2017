<?xml version="1.0" encoding="utf-8"?>
<!--<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">-->
<!--<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl">-->
  <xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" exclude-result-prefixes="xsl in lang user" xmlns:in="http://www.composite.net/ns/transformation/input/1.0" xmlns:lang="http://www.composite.net/ns/localization/1.0" xmlns:f="http://www.composite.net/ns/function/1.0" xmlns="http://www.w3.org/1999/xhtml" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:my-scripts">

  <xsl:output method="text" encoding="windows-1251" indent="no"/>
  <msxsl:script language="C#" implements-prefix="user">
    <msxsl:assembly name="System.Web" />
    <msxsl:assembly name="System.Data" />
    <msxsl:using namespace="System.Web" />
    <msxsl:using namespace="System.Text.RegularExpressions" />
    <msxsl:using namespace="System.Data.OleDb" />
    <![CDATA[
    public static string[,] strAffix = new string[,] { { " \\(РАЙОН\\)", " район" }, { " \\(Р\\-Н\\)", " район" }, { " \\(С\\.?\\)", "село" }, { " \\(Д\\.?\\)", "деревня" }, { " \\(П\\.Г\\.Т\\.\\)", "пгт" }, { " \\(ПГТ\\)", "пгт" }, { " \\(Р\\.П\\.\\)", "пгт" }, { " \\(СОВХОЗ\\)", "посёлок совхоза" }, { " \\(С/Х\\)", "посёлок совхоза" }, { " \\(ПОС\\. С/З\\)", "посёлок совхоза" }, { " \\(С/З\\)", "посёлок совхоза" }, { " \\(СТ\\.\\)", "посёлок железнодорожн" }, { " \\(ПОС\\. Ж\\.Д\\.СТ\\.\\)", "посёлок железнодорожн" }, { " \\(П\\.?\\)", "посёлок" }, { " \\(ПОС\\.?\\)", "посёлок" } };

    public string Name2Terson(string strNAME)
    {
      OleDbConnection connKLADR = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Environment.CurrentDirectory + ";Extended Properties=DBASE IV;Persist Security Info=False;");
      connKLADR.Open();
      string strSQL="";
      Regex rgxGorod = new Regex(@" \(г\.?\)$", RegexOptions.IgnoreCase);
      Regex rgxRayon = new Regex(@" (\(район\))|(\(р-н\))$", RegexOptions.IgnoreCase);
      string[] strN;
      strN = strNAME.Split(new string[] {", "}, StringSplitOptions.None);
      if (rgxGorod.IsMatch(strN[0])) // если ГОРОД
      {
          strN[0] = "г. " + strN[0].Replace(" (г)", "").Replace(" (Г.)", "").Trim();
          if (strN.GetLength(0) > 1) // если КАЗАНЬ с районом
          {
              strN[1] = Up2LowNames(rgxRayon.Replace(strN[1], " район"));
                  // TERSON до 2015
                  // strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + Up2LowNames(strN[0]) + "' AND NAME='" + strN[1] + "'";
                  // Terson 2015
                  strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + Up2LowNames(strN[0]) + "' AND NAME='" + strN[1] + " " + Up2LowNames(strN[0]) + "'";
          }
          else // если город без района
          {
              strSQL = "SELECT CODE FROM 92_TER WHERE NAME='" + Up2LowNames(strN[0]) + "'";
          }
      }
      else if (rgxRayon.IsMatch(strN[0])) // если РАЙОН
      {
          strN[0] = Up2LowNames(rgxRayon.Replace(strN[0], "")).Trim();
          if (strN.Length < 2) return strNAME; // если РАЙОН без запятой
          if (rgxGorod.IsMatch(strN[1])) // если РАЙОН + ГОРОД
          {
              strN[1] = "г. " + Up2LowNames(strN[1].Replace(" (г)", "").Replace(" (Г.)", "")).Trim();
              // TERSON до 2015
              // strSQL = "SELECT CODE FROM 92_TER WHERE Left(CODE, 2)='92' AND NAME='" + strN[1] + "'";
              // TERSON 2015
              strSQL = "SELECT CODE FROM 92_TER WHERE Left(CODE, 2)='92' AND NAME='" + strN[1] + " " + strN[0] + " район'";
          }
          else // если РАЙОН + село
          {
              strN[1] = Up2LowNames(ReformNames(strN[1])).Trim();
              // TERSON до 2015
              // strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + strN[0] + "' AND NAME='" + strN[1] + " " + "'";
              // TERSON 2015
              strSQL = "SELECT CODE FROM 92_TER, RAYON WHERE Left(CODE, 5)=KOD AND RAY = '" + strN[0] + "' AND NAME='" + strN[1] + " " + strN[0] + " район'";
          }
      }
      else // если что-то непонятное
      {
        connKLADR.Close();
        return strNAME;
      }
      OleDbCommand comm = connKLADR.CreateCommand();
      comm.CommandText = strSQL;
      //comm.CommandType = CommandType.Text;
      OleDbDataReader reader = comm.ExecuteReader();
      string strN2T = "";
      if (reader.HasRows) //true
      {
          //Console.WriteLine(reader.FieldCount); // 1 
          reader.Read();
          strN2T = reader.GetValue(0).ToString();
      }
      reader.Close();
      connKLADR.Close();
      return strN2T;

    }
      
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

    public static string Up2LowNames(string strName) // меняет регистр букв на "Первые Заглавные, Остальные Строчные" и убирает лишние пробелы
    {
        //Regex rgxUL = new Regex(@"([\s-]|^)([a-z0-9-_]+)",RegexOptions.IgnoreCase);
        //string NewName = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(strName.ToLower());
        Regex rgxU2L = new Regex (@"[А-Я][А-Я]+", RegexOptions.None);
        Regex rgxSpaces = new Regex(@"\s+");
        MatchCollection matches = rgxU2L.Matches(strName);
        foreach (Match match in matches)
            strName = strName.Replace(match.ToString().Substring(1), match.ToString().Substring(1).ToLower());
        strName = rgxSpaces.Replace(strName, " ");
        //Console.WriteLine(strName + ":::" + NewName);
        return strName;
    }]]>
  </msxsl:script>

  <xsl:template match="/">
    <xsl:text>ZAGS#CodeZAGS#OKATO#np#Pos#Acts#1#2#3#4#5#6#9#10#11#12#12TERSON#19#20#21#72#75#76#77#78&#xA;</xsl:text>
    <xsl:for-each select="/report/p">
      <xsl:apply-templates select="document(@filename)/report">
        <xsl:with-param name="pos" select="position()"/>
      </xsl:apply-templates>
    </xsl:for-each>
  </xsl:template>
  <!-- VERSION 2014-12-10 -->
  <xsl:template match="report">
    <xsl:param name="pos"/>
    <xsl:value-of select="title/item[@name='name']/@value"/><xsl:text>#</xsl:text><!--[] наименование ЗАГС-->
    <!--<xsl:value-of select="substring(title/item[@name='okato']/@value,1,8)"/>--><xsl:text>###</xsl:text><!--[] код ЗАГС (8 знаков), ОКАТО, np-->
    <xsl:value-of select="$pos"/><xsl:text>#</xsl:text><!--[] счетчик актов-->
    <xsl:value-of select="title/item[@name='akts']/@value"/><xsl:text>#</xsl:text><!--[] akts-->
    <xsl:value-of select="sections/section/row[@code='1']/col"/><xsl:text>#</xsl:text><!--[] месяц регистрации-->
    <xsl:value-of select="sections/section/row[@code='2']/col"/><xsl:text>#</xsl:text><!--[] год регистрации-->
    <xsl:value-of select="sections/section/row[@code='3']/col"/><xsl:text>#</xsl:text><!--[] день рождения-->
    <xsl:value-of select="sections/section/row[@code='4']/col"/><xsl:text>#</xsl:text><!--[] месяц рождения-->
    <xsl:value-of select="sections/section/row[@code='5']/col"/><xsl:text>#</xsl:text><!--[] год рождения-->
    <xsl:value-of select="sections/section/row[@code='6']/col"/><xsl:text>#</xsl:text><!--[] пол-->
    <xsl:value-of select="sections/section/row[@code='9']/col"/><xsl:text>#</xsl:text><!--[] место рождения-->
    <xsl:value-of select="sections/section/row[@code='10']/col"/><xsl:text>#</xsl:text><!--[] место смерти-->
    <xsl:value-of select="sections/section/row[@code='11']/col"/><xsl:text>#</xsl:text><!--[] субъект места жительства-->
    <xsl:value-of select="sections/section/row[@code='12']/col"/><xsl:text>#</xsl:text><!--[] место жительства-->
    <xsl:value-of select="user:Name2Terson(sections/section/row[@code='12']/col)"/><xsl:text>#</xsl:text><!--[] ТЕРСОН место жительства-->
    <xsl:value-of select="sections/section/row[@code='19']/col"/><xsl:text>#</xsl:text><!--[] день смерти-->
    <xsl:value-of select="sections/section/row[@code='20']/col"/><xsl:text>#</xsl:text><!--[] месяц смерти-->
    <xsl:value-of select="sections/section/row[@code='21']/col"/><xsl:text>#</xsl:text><!--[] год смерти-->
    <xsl:value-of select="sections/section/row[@code='72']/col"/><xsl:text>#</xsl:text><!--[] № свидетельства-->
    <xsl:value-of select="sections/section/row[@code='75']/col"/><xsl:text>#</xsl:text><!--[] день регистрации-->
    <xsl:value-of select="sections/section/row[@code='76']/col"/><xsl:text>#</xsl:text><!--[] месяц регистрации-->
    <xsl:value-of select="sections/section/row[@code='77']/col"/><xsl:text>#</xsl:text><!--[] год регистрации-->
    <xsl:value-of select="sections/section/row[@code='78']/col"/><xsl:text>&#xA;</xsl:text><!--[] код гражданства-->
  </xsl:template>
</xsl:stylesheet>
