using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Xml;
using System.Xml.Linq;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace VMS.TPS
{
    public class Script
    {

        // fix special character break *, run in plan sum, calculate age, increase font

        //public FileInfo fileinfo = new FileInfo(@"\\varnwcdcpap001.nyumc.org\va_data$\filedata\ProgramData\Vision\PublishedScripts\dx.html");
        public FileInfo fileinfo = new FileInfo(@"\\varwcdcpimgvm01.nyumc.org\va_transfer\DotDecimal\dx.html");
        //public FileInfo fileinfo = new FileInfo(@"H:\dx.html");
        public StreamWriter streamWriter;
        public FileStream fileStream;

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {

            string html = GrabDx(context);

            PlanSetup ps = context.PlanSetup;
            PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
            PlanningItem pi = ps != null ? (PlanningItem)ps : (PlanningItem)psum;

            string temp = System.Environment.GetEnvironmentVariable("TEMP");
            
            string htmlPath = string.Format("{0}\\{1}({2}){3}.html", temp,
              context.Patient.LastName, context.Patient.Id, pi.Id);
            using (System.IO.FileStream file = new System.IO.FileStream
              (htmlPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            { }
            //MessageBox.Show(htmlPath);
            //htmlPath = @"\\varnwcdcpap001.nyumc.org\va_data$\filedata\ProgramData\Vision\PublishedScripts\dx.html";
            //System.IO.File.WriteAllText(htmlPath, html);


            //XmlWriterSettings settings = new XmlWriterSettings();
            //settings.Indent = true;
            //settings.IndentChars = ("\t");
            //System.IO.MemoryStream mStream = new System.IO.MemoryStream();
            //using (XmlWriter writer = XmlWriter.Create(mStream, settings))
            //{
            //    // generate DVHs in an HTML report for selected structures
            //    ExportDVHs(context, writer);

            //    // done writing
            //    writer.Flush();
            //    mStream.Flush();

            //    // write the XML file.
            //    string temp = System.Environment.GetEnvironmentVariable("TEMP");
            //    string htmlPath = string.Format("{0}\\{1}({2})-plan-{3}.html", temp,
            //      context.Patient.LastName, context.Patient.Id, context.PlanSetup.Id);
            //    using (System.IO.FileStream file = new System.IO.FileStream
            //      (htmlPath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            //    {
            //        // Have to rewind the MemoryStream in order to read its contents. 
            //        mStream.Position = 0;
            //        mStream.CopyTo(file);
            //        file.Flush();
            //        file.Close();
            //    }

            WriteHtml(html);

            // 'Start' generated HTML file to launch browser window
            //System.Diagnostics.Process.Start(htmlPath);
            System.Diagnostics.Process.Start(fileinfo.FullName);
            // Sleep for a few seconds to let internet browser window to start
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
        }

        public string GrabDx(ScriptContext context)
        {

//            string sql = "SELECT DISTINCT dbo.pt.patient_ser, LTRIM(RTRIM(dbo.pt.pt_last_name)) + ', ' + LTRIM(RTRIM(dbo.pt.pt_first_name)) AS Patient, variansystem.dbo.Patient.PatientId AS MRN, dbo.dx_typ.dx_typ_desc, CAST(dbo.Diagnosis.DateStamp AS DATE) as DxDate, dbo.Diagnosis.HstryUserName AS DxUser, dbo.Diagnosis.DiagnosisTableName, dbo.Diagnosis.DiagnosisCode, dbo.Diagnosis.Description,                          dbo.Diagnosis.DiagnosisType, dbo.icdo_site_cd.icdo_desc, dbo.icdo_site_typ.icdo_site_desc, CAST(dbo.pt_dx.dx_desc AS nvarchar(max)) AS dx_desc "
//+ "FROM dbo.icdo_site_cd INNER JOIN dbo.icdo_site_typ ON dbo.icdo_site_cd.icdo_site_typ = dbo.icdo_site_typ.icdo_site_typ INNER JOIN dbo.pt INNER JOIN dbo.pt_dx ON dbo.pt.pt_id = dbo.pt_dx.pt_id INNER JOIN dbo.dx_typ ON dbo.pt_dx.dx_typ = dbo.dx_typ.dx_typ INNER JOIN dbo.Diagnosis ON dbo.pt_dx.diagnosis_ser = dbo.Diagnosis.DiagnosisSer INNER JOIN dbo.pt_dx_cncr ON dbo.pt_dx.pt_id = dbo.pt_dx_cncr.pt_id AND dbo.pt_dx.dx_id = dbo.pt_dx_cncr.pt_dx_id ON dbo.icdo_site_cd.icdo_site_cd = dbo.pt_dx_cncr.icdo_site_cd AND dbo.icdo_site_cd.icdo_site_seq = dbo.pt_dx_cncr.icdo_site_seq, [variansystem].[dbo].Patient WHERE dbo.pt.patient_ser = [variansystem].[dbo].Patient.PatientSer AND"
//+ "[variansystem].[dbo].Patient.PatientId = '" + context.Patient.Id + "' ";

            string sql = "SELECT DISTINCT dbo.pt.patient_ser, LTRIM(RTRIM(dbo.pt.pt_last_name)) + ', ' + LTRIM(RTRIM(dbo.pt.pt_first_name)) AS Patient, variansystem.dbo.Patient.PatientId AS MRN, dbo.dx_typ.dx_typ_desc, CASE WHEN dbo.Diagnosis.DateStamp is null then '' else CAST(dbo.Diagnosis.DateStamp AS DATE) end AS DxDate, case when dbo.Diagnosis.HstryUserName is null then '' else dbo.Diagnosis.HstryUserName end AS DxUser, case when dbo.Diagnosis.DiagnosisTableName is null then '' else dbo.Diagnosis.DiagnosisTableName end as DiagnosisTableName, case when dbo.Diagnosis.DiagnosisCode is null then '' else dbo.Diagnosis.DiagnosisCode end as DiagnosisCode, case when dbo.Diagnosis.Description is null then '' else dbo.Diagnosis.Description end as Description, case when dbo.Diagnosis.DiagnosisType is null then '' else dbo.Diagnosis.DiagnosisType end as DiagnosisType, case when dbo.icdo_site_cd.icdo_desc is null then '' else dbo.icdo_site_cd.icdo_desc end as icdo_desc, case when dbo.icdo_site_typ.icdo_site_desc is null then '' else dbo.icdo_site_typ.icdo_site_desc end as icdo_site_desc, CASE WHEN CAST(dbo.pt_dx.dx_desc AS nvarchar(MAX)) = '' THEN CAST(dbo.pt_dx.clinical_desc AS nvarchar(MAX)) ELSE CAST(dbo.pt_dx.dx_desc AS nvarchar(MAX)) END AS dx_desc "
                + " FROM dbo.dx_typ RIGHT OUTER JOIN dbo.icdo_site_cd RIGHT OUTER JOIN dbo.pt_dx_cncr ON dbo.icdo_site_cd.icdo_site_cd = dbo.pt_dx_cncr.icdo_site_cd AND dbo.icdo_site_cd.icdo_site_seq = dbo.pt_dx_cncr.icdo_site_seq RIGHT OUTER JOIN dbo.pt INNER JOIN dbo.pt_dx ON dbo.pt.pt_id = dbo.pt_dx.pt_id INNER JOIN variansystem.dbo.Patient ON dbo.pt.patient_ser = variansystem.dbo.Patient.PatientSer ON dbo.pt_dx_cncr.pt_id = dbo.pt_dx.pt_id AND dbo.pt_dx_cncr.pt_dx_id = dbo.pt_dx.dx_id LEFT OUTER JOIN dbo.icdo_site_typ ON dbo.icdo_site_cd.icdo_site_typ = dbo.icdo_site_typ.icdo_site_typ ON dbo.dx_typ.dx_typ = dbo.pt_dx.dx_typ LEFT OUTER JOIN dbo.Diagnosis ON dbo.pt_dx.diagnosis_ser = dbo.Diagnosis.DiagnosisSer WHERE (variansystem.dbo.Patient.PatientId = '" + context.Patient.Id + "')";

            DataTable dt = AriaSQL(sql, "enm");

            DataTable RxTable;
            string Rxsql = "SELECT DISTINCT dbo.Course.CourseId, dbo.Prescription.PrescriptionName, dbo.Prescription.Site, dbo.Prescription.PhaseType, CASE dbo.Prescription.SimulationNeeded WHEN 1 THEN 'YES' WHEN 0 THEN 'NO' END AS[CT Simulation], dbo.Prescription.Status, dbo.Prescription.HstryUserName AS[Approved By], dbo.Prescription.HstryDateTime AS[Approval Date], "
                    + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'DOSE PER FRACTION' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS DosePerFx, dbo.Prescription.NumberOfFractions, "
                    + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'TOTAL DOSE' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS TotalDose, "
                    + "(SELECT ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'VOLUME ID' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS Volume, "
                    + "substring((SELECT ',' + PropertyValue AS[text()] FROM dbo.PrescriptionProperty PP Where PP.PropertyType = 1 AND PP.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Energies, "
                    + "substring((SELECT ',' + PP2.PropertyValue AS [text()] FROM dbo.PrescriptionProperty PP2 WHERE PP2.PropertyType = 2 AND PP2.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Mode, dbo.Prescription.Technique, dbo.Prescription.Notes FROM dbo.PrescriptionProperty RIGHT OUTER JOIN dbo.Patient INNER JOIN dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer INNER JOIN dbo.TreatmentPhase ON dbo.Course.CourseSer = dbo.TreatmentPhase.CourseSer INNER JOIN dbo.Prescription ON dbo.TreatmentPhase.TreatmentPhaseSer = dbo.Prescription.TreatmentPhaseSer INNER JOIN dbo.PrescriptionAnatomy ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionAnatomy.PrescriptionSer INNER JOIN dbo.PrescriptionAnatomyItem ON dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer ON dbo.PrescriptionProperty.PrescriptionSer = dbo.Prescription.PrescriptionSer ";

            RxTable = AriaSQL(Rxsql + " WHERE (dbo.PrescriptionAnatomy.AnatomyRole = 2) AND (dbo.Patient.PatientId = '" + context.Patient.Id + "')"
                + "ORDER BY dbo.Prescription.HstryDateTime DESC", "system");


            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body>");
            strHTMLBuilder.Append("<div>");
            strHTMLBuilder.Append("<H3>" + context.Patient.Name + "</H3>");
            strHTMLBuilder.Append("<H4>DOB: " + Convert.ToDateTime(context.Patient.DateOfBirth).ToString("d") + " (" + ToAgeString(Convert.ToDateTime(context.Patient.DateOfBirth)) + ")" + "</H4>");
            strHTMLBuilder.Append("<H4>Gender: " + context.Patient.Sex.ToString());
            strHTMLBuilder.Append("</div>");

            if (dt.Rows.Count == 0)
            {
                strHTMLBuilder.Append("<div><h4>No diagnoses could be found.</h4></div>");
            }
            foreach (DataRow dr in dt.Rows)
            {

                strHTMLBuilder.Append("<div style=\"width:100%;\">");
                strHTMLBuilder.Append("<H3>Diagnoses</H3>");
                strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='plum' style='font-family:Arial; font-size:larger'>");

                // Build the column names
                strHTMLBuilder.Append("<tr >");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("DxDate");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("DxUser");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("icdo_site_desc");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("DiagnosisCode");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("icdo_desc");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append("Description");
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("</tr>");

                // Fill with data
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(Convert.ToDateTime(dr["DxDate"]).ToString("d"));
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(dr["DxUser"].ToString());
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(dr["icdo_site_desc"].ToString());
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(dr["DiagnosisCode"].ToString());
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(dr["icdo_desc"].ToString());
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(dr["Description"].ToString());
                strHTMLBuilder.Append("</td>");
                strHTMLBuilder.Append("</tr>");
                strHTMLBuilder.Append("</table>");
                strHTMLBuilder.Append(dr["dx_desc"].ToString());
                strHTMLBuilder.Append("<br/>");
            }

            System.Text.StringBuilder RxList = new System.Text.StringBuilder();
            strHTMLBuilder.Append("<h3>Prescriptions</h3>");
            foreach (System.Data.DataRow r in RxTable.Rows)
            {
                RxList.Append("<div style=\"width:100%; float:left;\">");
                string strcolor = "";
                string strcolor2 = "";
                if (r["Status"].ToString() != "Approved")
                {
                    strcolor = "<span style = \"color:Red;font-weight:bold;\">";
                    strcolor2 = "</span>";
                }
                RxList.Append("<div style=\"width: 100%;\"><h3>" + r["CourseId"].ToString() + " - " + r["Site"].ToString() + " - " + strcolor + r["Status"].ToString() + strcolor2 + " (" + r["Approved By"].ToString() + " - " + r["Approval Date"].ToString() + ")</h3><tr>");
                RxList.Append("<div style=\"width: 25 %; float:left;\">");
                RxList.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='plum' style='font-family:Arial;'>");
                RxList.Append("<tr><td>Rx Name:</td><td>" + r["PrescriptionName"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Fractions:</td><td>" + r["NumberOfFractions"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Prescribe to:</td><td>" + r["Volume"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Total Dose:</td><td>" + r["TotalDose"].ToString() + " cGy" + "</td></tr>");
                RxList.Append("<tr><td>Dose/fx:</td><td>" + r["DosePerFx"].ToString() + " cGy" + "</td></tr>");

                RxList.Append("</table></div>");

                RxList.Append("<div style=\"width: 25 %; float:left;\">");
                RxList.Append("<table border='1px' cellpadding='1' cellspacing='1' bgcolor='plum' style='font-family:Arial;'>");
                RxList.Append("<tr><td>Mode:</td><td>" + r["Mode"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Energies:</td><td>" + r["Energies"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Technique:</td><td>" + r["Technique"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>CT Sim:</td><td>" + r["CT Simulation"].ToString() + "</td></tr>");
                RxList.Append("<tr><td>Notes:</td><td>" + r["Notes"].ToString() + "</td></tr>");

                RxList.Append("</table></div></div>");

                RxList.Append("</div>");
                //foreach (DataColumn c in RxTable.Columns)
                //{
                //    RxList.Append(c.ColumnName.ToString() + ": " + r[c.ColumnName].ToString() + "   " + '\n');
                //}
                //RxList.Append("<br/>");
            }
            strHTMLBuilder.Append(RxList.ToString());
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;
        }

        public DataTable AriaSQL(string query, string db)
        {
            //string out_string = "";
            string connectionString = "";
            if (db == "enm")
            {
                connectionString = "Data Source=varwcdcpdbavm01;Persist Security Info=True;Password=reports;User ID=reports;Initial Catalog=varianenm";
            }
            else {
                connectionString = "Data Source=varwcdcpdbavm01;Persist Security Info=True;Password=reports;User ID=reports;Initial Catalog=variansystem";
            }

            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlDataReader reader = null;
                SqlCommand command = new SqlCommand(query, con);
                con.Open();
                reader = command.ExecuteReader();
                dt.Load(reader);
                //if (reader.HasRows)k
                //{
                //    while (reader.Read())
                //    {
                //        out_string += reader["MachineId"].ToString() + "   " + Convert.ToDateTime(reader["ScheduledStartTime"]).ToString("M/d/yyyy h:mm tt") + '\n';
                //    }
                //}
                con.Close();
            }
            return dt;
        }

        void WriteHtml(string strComments)
        {
            OpenFile();
            streamWriter.WriteLine(strComments);
            CloseFile();
        }

        void OpenFile()
        {
            DateTime dt = DateTime.Now;

            fileStream = new FileStream(fileinfo.FullName, FileMode.Create, FileAccess.Write);
            streamWriter = new StreamWriter(fileStream);
        }

        void CloseFile()
        {
            streamWriter.Close();
            fileStream.Close();
            GC.Collect();
        }

        public static string ToAgeString(DateTime dob)
        {
            // http://stackoverflow.com/questions/3054715/c-sharp-calculate-accurate-age
            DateTime today = DateTime.Today;

            int months = today.Month - dob.Month;
            int years = today.Year - dob.Year;

            if (today.Day < dob.Day)
            {
                months--;
            }

            if (months < 0)
            {
                years--;
                months += 12;
            }

            int days = (today - dob.AddMonths((years * 12) + months)).Days;

            return string.Format("{0} year{1}, {2} month{3} and {4} day{5}",
                                 years, (years == 1) ? "" : "s",
                                 months, (months == 1) ? "" : "s",
                                 days, (days == 1) ? "" : "s");
        }
    }
}