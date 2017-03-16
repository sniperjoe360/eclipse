using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Collections.Generic;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Text.RegularExpressions;
using System.Data.SqlClient;

/*  Plan Check script written by David Barbee 
    DISCLAIMER
    STOP!  USE THIS SCRIPT AT YOUR OWN RISK!!!
    
    Author is not responsible for any errors, issues, misadministrations, or general outcomes or results of using this script.
    
    This code is still in testing and cannot possibly catch every single possible planning error, so please be diligent when 
    checking your plans.

    Many of the triggered catches are specific to the author's institution and may not be reflective of conventions used at your institution.
*/

namespace VMS.TPS

{

    public class Script
    {

        #region Variables and functions

        string ErrorLog = "";

        public Boolean VMAT;

        public FileInfo fi = new FileInfo(@"\\varwcdcpimgvm01.nyumc.org\va_data$\filedata\ProgramData\Vision\PublishedScripts\PlanCheckLog.log");
        public StreamWriter streamWriter;
        public FileStream fileStream;

        public Structure TargetStruct;

        struct BeamOverview
        {
            public string Name;
            public string Issues;
        }

        struct RV_Struct
        {
            public string Name;
            public bool Empty;
            public List<int> SliceList;
            public IEnumerable<int> MissingSliceList;
            public bool MissingSlices;
            public int MinSlice;
            public int MaxSlice;
            public bool DensityOvrd;
            public double Density;
            public bool HasLaterality;
            public string Laterality;
            public void InitSlice()
            {
                SliceList = new List<int>();
            }
            public bool Target;
        }

        List<string[]> beamnames = new List<string[]>();

        public string Vec2Str(VVector v)
        {
            return v.x.ToString() + ", " + v.y.ToString() + ", " + v.z.ToString();
        }
        public VVector VecAdd(VVector a, VVector b)
        {
            return new VVector(a.x + b.x, a.y + b.y, a.z + b.z);
        }
        public VVector VecSubtract(VVector a, VVector b)
        {
            return new VVector(a.x - b.x, a.y - b.y, a.z - b.z);
        }
        public VVector VecRound(VVector a, int d)
        {
            return new VVector(Math.Round(a.x, d), Math.Round(a.y, d), Math.Round(a.z, d));
        }
        public VVector VecMultiply(VVector a, double d)
        {
            return new VVector(a.x * d, a.y * d, a.z * d);
        }

        public string ReturnShiftDirections(string patorientation, VVector shift)
        {
            string out1 = "";
            switch (patorientation)
            {
                case "HeadFirstSupine":
                    if (shift.x < 0) { out1 += Math.Abs(shift.x).ToString() + " Right, "; } else { out1 += Math.Abs(shift.x).ToString() + " Left, "; }
                    if (shift.y < 0) { out1 += Math.Abs(shift.y).ToString() + " Ant, "; } else { out1 += Math.Abs(shift.y).ToString() + " Post, "; }
                    if (shift.z < 0) { out1 += Math.Abs(shift.z).ToString() + " Inf, "; } else { out1 += Math.Abs(shift.z).ToString() + " Sup, "; }
                    break;
                case "HeadFirstProne":
                    if (shift.x < 0) { out1 += Math.Abs(shift.x).ToString() + " Right, "; } else { out1 += Math.Abs(shift.x).ToString() + " Left, "; }
                    if (shift.y < 0) { out1 += Math.Abs(shift.y).ToString() + " Ant, "; } else { out1 += Math.Abs(shift.y).ToString() + " Post, "; }
                    if (shift.z < 0) { out1 += Math.Abs(shift.z).ToString() + " Inf, "; } else { out1 += Math.Abs(shift.z).ToString() + " Sup, "; }
                    //if (shift.x < 0) { out1 += Math.Abs(shift.x).ToString() + " Right, "; } else { out1 += Math.Abs(shift.x).ToString() + " Left, "; }
                    //if (shift.y < 0) { out1 += Math.Abs(shift.y).ToString() + " Post, "; } else { out1 += Math.Abs(shift.y).ToString() + " Ant, "; }
                    //if (shift.z < 0) { out1 += Math.Abs(shift.z).ToString() + " Inf, "; } else { out1 += Math.Abs(shift.z).ToString() + " Sup, "; }
                    break;
                case "FeetFirstSupine":
                    if (shift.x < 0) { out1 += Math.Abs(shift.x).ToString() + " Right, "; } else { out1 += Math.Abs(shift.x).ToString() + " Left, "; }
                    if (shift.y < 0) { out1 += Math.Abs(shift.y).ToString() + " Ant, "; } else { out1 += Math.Abs(shift.y).ToString() + " Post, "; }
                    if (shift.z < 0) { out1 += Math.Abs(shift.z).ToString() + " Sup, "; } else { out1 += Math.Abs(shift.z).ToString() + " Inf, "; }
                    break;
                case "FeetFirstProne":
                    if (shift.x < 0) { out1 += Math.Abs(shift.x).ToString() + " Right, "; } else { out1 += Math.Abs(shift.x).ToString() + " Left, "; }
                    if (shift.y < 0) { out1 += Math.Abs(shift.y).ToString() + " Post, "; } else { out1 += Math.Abs(shift.y).ToString() + " Ant, "; }
                    if (shift.z < 0) { out1 += Math.Abs(shift.z).ToString() + " Sup, "; } else { out1 += Math.Abs(shift.z).ToString() + " Inf, "; }
                    break;
            }

            out1 = out1.Substring(0, out1.LastIndexOf(", "));
            return out1;
        }

        public string FindCTCouch(Image I, PlanSetup p)
        {
            string out_string = "CT Scan Search:" + '\n';

            // Switch search direction if prone
            int s = 1;
            if (p.TreatmentOrientation.ToString().Contains("Prone")) { s = -1; }

            // find the beam isocenter and work your way down until the +20 -20 is uniform
            List<double> SD = new List<double> { };
            VVector iso = p.Beams.First().IsocenterPosition;
            double[] x_prof = new double[512];
            double[] z_prof = new double[512];
            double min_ave, max_ave, max_sd, search;
            min_ave = -800;
            max_ave = -300;
            max_sd = 100;
            search = 60;

            // This gives the image origin, X_center pos, User origin, and Iso of the 1st beam
            string test_string = Vec2Str(I.Origin) + '\n' + (I.Origin.x + I.XRes * I.XSize / 2).ToString() + '\n' + Vec2Str(I.UserOrigin) + '\n' + Vec2Str(p.Beams.First().IsocenterPosition);

            for (int i = 0; i < I.YSize; i++)  // this loops through depth, moving down (orientation dependent)
            {
                VVector iso_sub20 = new VVector(-search, iso.y + s * i * I.YRes, iso.z);
                VVector iso_add20 = new VVector(search, iso.y + s * i * I.YRes, iso.z);
                ImageProfile ImProf = I.GetImageProfile(iso_sub20, iso_add20, x_prof);
                double average = x_prof.Average();
                double sumOfSquaresOfDifferences = x_prof.Select(val => (val - average) * (val - average)).Sum();
                double sd = Math.Sqrt(sumOfSquaresOfDifferences / x_prof.Length);

                if ((average < max_ave) & (average > min_ave) & (sd < max_sd))
                {  // search left to right
                    VVector y_search_add10 = new VVector(0, iso.y + s * i * I.YRes, iso.z + search);  // 0 to get the center of the image
                    VVector y_search_sub10 = new VVector(0, iso.y + s * i * I.YRes, iso.z - search);
                    ImageProfile ImProf2 = I.GetImageProfile(y_search_sub10, y_search_add10, z_prof);
                    double average2 = z_prof.Average();
                    double sumOfSquaresOfDifferences2 = z_prof.Select(val => (val - average2) * (val - average2)).Sum();
                    double sd2 = Math.Sqrt(sumOfSquaresOfDifferences2 / z_prof.Length);
                    //MessageBox.Show(average2.ToString() + " " + max_ave.ToString() + " " + min_ave.ToString() + '\n' + sd2.ToString() + " " + max_sd.ToString());
                    if ((average2 < max_ave) & (average2 > min_ave) & (sd2 < max_sd))
                    {  // search in and out
                       //out_string = (0.5 * 0.1 * I.YRes * SD.IndexOf(SD.Min())).ToString() + " TT measurement.";
                       //out_string += ((iso.y - I.YRes * i) / 10).ToString() + " TT measurement" + '\n';
                       //MessageBox.Show("TT = " + ((I.YRes * i - 3) / 10).ToString("F1") + " cm"); // + '\n' + average.ToString() + "  " + sd.ToString() + '\n';
                        out_string += "TT = " + ((I.YRes * i - 0) / 10).ToString("F1") + " cm"; // + '\n' + average.ToString() + "  " + sd.ToString() + '\n';
                        break;  // leave the first couch you find.  Searching down from iso
                    }

                }
                SD.Add(sd);
                //out_string += sd.ToString() + ", ";
            }
            //out_string = (0.5 * 0.1 * I.YRes * SD.IndexOf(SD.Min())).ToString() + " TT measurement.";

            if (!out_string.Contains("TT = "))  // if the previous script couldn't find anything, assume dosemax table and search up down as well.
            {

            }
            return out_string;

            //VVector iso_sub20 = new VVector(iso.x - 150, iso.y - 80, iso.z);
            //VVector iso_add20 = new VVector(iso.x + 150, iso.y - 80, iso.z);
            //ImageProfile ImProf = I.GetImageProfile(iso_sub20, iso_add20, prof);

            //return string.Join(",", prof.Select(x => x.ToString()).ToArray());
            ////MessageBox.Show(I.XRes + ", " + I.YRes + ", " + I.ZRes);

        }

        public string FindPortraitCouch(PlanSetup p, Structure S)
        {
            double minZ = S.MeshGeometry.Positions.Min(a => a.Z);
            //double maxZ = S.MeshGeometry.Positions.Max(a => a.Z);
            //string first_temp = (body.MeshGeometry.Positions.First(a => a.Z == minZ).X / 10).ToString() + ", " + (body.MeshGeometry.Positions.First(a => a.Z == minZ).Y / 10).ToString() + ", " + (body.MeshGeometry.Positions.First(a => a.Z == minZ).Z / 10).ToString();
            //string last_temp = (body.MeshGeometry.Positions.Last(a => a.Z == minZ).X / 10).ToString() + ", " + (body.MeshGeometry.Positions.Last(a => a.Z == minZ).Y / 10).ToString() + ", " + (body.MeshGeometry.Positions.Last(a => a.Z == minZ).Z / 10).ToString();

            //return "Portrait = " + ((S.MeshGeometry.Positions.Last(a => a.Z == minZ).Y-I.UserOrigin.y)/10).ToString("F1") + " cm";

            // THIS WAY MIGHT BE BETTER, RATHER than looking at the last point in the slice, it looks at the max point in the contour
            //MessageBox.Show(string.Join(",", (S.MeshGeometry.Positions.Where(a => a.Z == minZ).Max(b => b.Y))));
            //return "Portrait = " + ((S.MeshGeometry.Positions.Where(a => a.Z == minZ).Max(b => b.Y) - I.UserOrigin.y) / 10).ToString("F1") + " cm";

            // Just give me the min along the sagittal plane
            VVector iso = p.Beams.First().IsocenterPosition;
            //return "MinBody = " + ((S.MeshGeometry.Positions.Where(a => a.Z == minZ).Max(b => b.Y) - iso.y) / 10).ToString("F1") + " cm";

            try
            {
                if (p.TreatmentOrientation.ToString().Contains("Prone"))
                {
                    return "MinBody = " + ((iso.y - S.MeshGeometry.Positions.Where(a => Math.Sqrt(a.X * a.X) < 3).Min(b => b.Y)) / 10).ToString("F1") + " cm";
                }
                else
                {
                    return "MinBody = " + ((S.MeshGeometry.Positions.Where(a => Math.Sqrt(a.X * a.X) < 3).Max(b => b.Y) - (iso.y)) / 10).ToString("F1") + " cm";
                }
            }
            catch
            {
                // this will shift the search radius around the x isocenter, rather than the center of the table.  In the situation where there is no body contour along x=0 (two legs)
                return "MinBody = " + ((S.MeshGeometry.Positions.Where(a => Math.Sqrt((a.X - iso.x) * (a.X - iso.x)) < 3).Max(b => b.Y) - (iso.y)) / 10).ToString("F1") + " cm";
            }
        }

        public string FindMaxCouch(PlanSetup p, Structure S)
        {
            VVector iso = p.Beams.First().IsocenterPosition;
            if (p.TreatmentOrientation.ToString().Contains("Prone"))
            {
                return "CouchSurface = " + ((iso.y - S.MeshGeometry.Positions.Max(a => a.Y)) / 10).ToString("F1") + " cm";
            }
            else
            {
                return "CouchSurface = " + ((S.MeshGeometry.Positions.Min(a => a.Y) - iso.y) / 10).ToString("F1") + " cm";
            }
        }

        public string CheckAnatomicFieldName(string ptorient, double gantry, string fldid)
        {
            string returnstring = "";
            gantry = Math.Round(gantry, 0);
            string origID = fldid;
            fldid = fldid.Replace(" ", "").Replace("_", "").ToLower(); // make the field ID lowercase and remove all whitespace to help with the matching
            List<string> match = new List<string>();
            List<string> AP = new List<string> { "ap", "ant" };
            List<string> LLAT = new List<string> { "ll", "ltlat", "llat", "leftlat", "ltla" };
            List<string> PA = new List<string> { "pa" };
            List<string> RLAT = new List<string> { "rl", "rtlat", "rlat", "rightlat", "rtla" };
            List<string> LAO = new List<string> { "lao" };
            List<string> LPO = new List<string> { "lpo" };
            List<string> RAO = new List<string> { "rao" };
            List<string> RPO = new List<string> { "rpo" };

            string cName = "";
            int sector = 1;
            if ((gantry > 0) & (gantry < 90)) { sector = 1; }
            if ((gantry > 90) & (Math.Round(gantry) < 180)) { sector = 2; }
            if ((Math.Round(gantry) > 180) & (gantry < 270)) { sector = 3; }
            if ((gantry > 270) & (gantry < 360)) { sector = 4; }
            if (gantry == 0) { sector = 5; }
            if (gantry == 90) { sector = 6; }
            if (gantry == 180) { sector = 7; }
            if (gantry == 270) { sector = 8; }

            switch (ptorient)
            {
                case "HeadFirstSupine":
                    switch (sector)
                    {
                        case 1:
                            match = LAO; cName = "LAO"; break;
                        case 2:
                            match = LPO; cName = "LPO"; break;
                        case 3:
                            match = RPO; cName = "RPO"; break;
                        case 4:
                            match = RAO; cName = "RAO"; break;
                        case 5:
                            match = AP; cName = "AP"; break;
                        case 6:
                            match = LLAT; cName = "LLAT"; break;
                        case 7:
                            match = PA; cName = "PA"; break;
                        case 8:
                            match = RLAT; cName = "RLAT"; break;
                    }
                    break;
                case "HeadFirstProne":
                    switch (sector)
                    {
                        case 1:
                            match = RPO; cName = "RPO"; break;
                        case 2:
                            match = RAO; cName = "RAO"; break;
                        case 3:
                            match = LAO; cName = "LAO"; break;
                        case 4:
                            match = LPO; cName = "LPO"; break;
                        case 5:
                            match = PA; cName = "PA"; break;
                        case 6:
                            match = RLAT; cName = "RLAT"; break;
                        case 7:
                            match = AP; cName = "AP"; break;
                        case 8:
                            match = LLAT; cName = "LLAT"; break;
                    }
                    break;
                case "FeetFirstSupine":
                    switch (sector)
                    {
                        case 1:
                            match = RAO; cName = "RAO"; break;
                        case 2:
                            match = RPO; cName = "RPO"; break;
                        case 3:
                            match = LPO; cName = "LPO"; break;
                        case 4:
                            match = LAO; cName = "LAO"; break;
                        case 5:
                            match = AP; cName = "AP"; break;
                        case 6:
                            match = RLAT; cName = "RLAT"; break;
                        case 7:
                            match = PA; cName = "PA"; break;
                        case 8:
                            match = LLAT; cName = "LLAT"; break;
                    }
                    break;
                case "FeetFirstProne":
                    switch (sector)
                    {
                        case 1:
                            match = LPO; cName = "LPO"; break;
                        case 2:
                            match = LAO; cName = "LAO"; break;
                        case 3:
                            match = RAO; cName = "RAO"; break;
                        case 4:
                            match = RPO; cName = "RPO"; break;
                        case 5:
                            match = PA; cName = "PA"; break;
                        case 6:
                            match = LLAT; cName = "LLAT"; break;
                        case 7:
                            match = AP; cName = "AP"; break;
                        case 8:
                            match = RLAT; cName = "RLAT"; break;
                    }
                    break;
            }
            if ((match.Any(fldid.Contains)) | fldid.Contains("cbct") | fldid.Contains("bolus"))
            {
                if (fldid.Contains("cbct")) // check to see if the cbct field has gantry 0, doesn't matter too much.
                {
                    if (gantry != 0)
                    {
                        returnstring = origID + " does not have gantry 0.";
                    }
                }
            }
            else
            {
                returnstring = origID + " is not labeled correctly.  Should be: " + cName;
            }
            return returnstring;
        }

        public string CheckSetupField(string patorientation, string fldid, double gantry)
        {
            string origID = fldid;
            fldid = fldid.Replace(" ", "").ToLower(); // make the field ID lowercase and remove all whitespace to help with the matching
            List<string> match = new List<string>();
            List<string> AP = new List<string> { "ap", "ant" };
            List<string> LLAT = new List<string> { "ll", "ltlat", "llat", "leftlat" };
            List<string> PA = new List<string> { "pa" };
            List<string> RLAT = new List<string> { "rl", "rtlat", "rlat", "rightlat" };

            // assumes that setup fields will be orthogonal
            string cName = "";
            switch (patorientation)
            {
                case "HeadFirstSupine":
                    switch (gantry.ToString())
                    {
                        case "0":
                            match = AP; cName = "AP";
                            break;
                        case "90":
                            match = LLAT; cName = "LLAT";
                            break;
                        case "180":
                            match = PA; cName = "PA";
                            break;
                        case "270":
                            match = RLAT; cName = "RLAT";
                            break;
                    }
                    break;
                case "HeadFirstProne":
                    switch (gantry.ToString())
                    {
                        case "180":
                            match = AP; cName = "AP";
                            break;
                        case "270":
                            match = LLAT; cName = "LLAT";
                            break;
                        case "0":
                            match = PA; cName = "PA";
                            break;
                        case "90":
                            match = RLAT; cName = "RLAT";
                            break;
                    }
                    break;
                case "FeetFirstSupine":
                    switch (gantry.ToString())
                    {
                        case "0":
                            match = AP; cName = "AP";
                            break;
                        case "270":
                            match = LLAT; cName = "LLAT";
                            break;
                        case "180":
                            match = PA; cName = "PA";
                            break;
                        case "90":
                            match = RLAT; cName = "RLAT";
                            break;
                    }
                    break;
                case "FeetFirstProne":
                    switch (gantry.ToString())
                    {
                        case "180":
                            match = AP; cName = "AP";
                            break;
                        case "90":
                            match = LLAT; cName = "LLAT";
                            break;
                        case "0":
                            match = PA; cName = "PA";
                            break;
                        case "270":
                            match = RLAT; cName = "RLAT";
                            break;
                    }
                    break;
            }
            if ((match.Any(fldid.Contains)) | fldid.Contains("cbct") | fldid.Contains("bolus"))
            {
                return null;
            }
            else
            {
                return origID + " is not labeled correctly." + '\n' + "Should be: " + cName;
            }
        }

        public float CheckOverlap(Structure oar, Structure ptv)
        {
            float overlap = 0;

            System.Windows.Media.Media3D.MeshGeometry3D a = oar.MeshGeometry;
            System.Windows.Media.Media3D.MeshGeometry3D p = ptv.MeshGeometry;

            System.Windows.Media.Media3D.GeometryModel3D g = new System.Windows.Media.Media3D.GeometryModel3D(p, new System.Windows.Media.Media3D.DiffuseMaterial());            //p.Positions

            //System.Windows.Media.Media3D. cg = new System.Windows.Media.CombinedGeometry(a, p);
            return overlap;
        }

        public string CheckIsland(Structure struc, int planes)
        {
            for (int i = 0; i < planes; i++)
            {
                if (struc.GetContoursOnImagePlane(i).Count() != 0)
                {
                    VVector[][] slice = struc.GetContoursOnImagePlane(i);
                }
            }

            return "";
        }

        public string CheckReferencePoints(PlanSetup p)
        {
            String refpoints = "";
            string[,] RParray = new string[p.Beams.Count(), p.Beams.First().FieldReferencePoints.Count()];
            int m, n;
            m = 0; n = 0;
            foreach (Beam b in p.Beams)
            {
                n = 0;
                foreach (var bb in b.FieldReferencePoints)
                {
                    RParray[m, n] = (Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F2");
                    refpoints += "Id: " + bb.ReferencePoint.Id + ", IsPrimary: " + bb.IsPrimaryReferencePoint + ", EffDepth: " + (Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F1") + ", Dose: " + bb.FieldDose + '\n';
                    n = n + 1;
                }
                m = m + 1;
            }
            string out_string = "BEAM_ID, ";
            for (int i = 0; i < p.Beams.First().FieldReferencePoints.Count(); i++)
            {
                out_string += p.Beams.First().FieldReferencePoints.ElementAt(i).ReferencePoint.Id + ", ";
            }
            out_string += '\n';
            for (int i = 0; i < p.Beams.Count(); i++)
            {
                out_string += p.Beams.ElementAt(i).Id + ": ";
                for (int j = 0; j < p.Beams.First().FieldReferencePoints.Count(); j++)
                {
                    out_string += RParray[i, j] + ", ";
                }
                out_string += '\n';
            }
            //MessageBox.Show(out_string);
            //return refpoints;
            return out_string;
        }

        public void GetFieldNames(ScriptContext context)
        {
            char[] splitter = { ' ', '_', 'g' };

            foreach (Course c in context.Patient.Courses)
            {
                foreach (PlanSetup ps in c.PlanSetups)
                {
                    if (ps.ApprovalStatus == PlanSetupApprovalStatus.PlanningApproved)
                    {
                        foreach (Beam b in ps.Beams)
                        {
                            string[] fieldname = b.Id.Split(splitter);
                            beamnames.Add(new string[] { fieldname.First().ToString(), b.Id, ps.Id });
                        }
                    }
                }
            }
        }

        // fds_1   
        //432_5  
        //rea_12  12_43  1_g180  1_g180_1  fds 1  fds_12  111_111   1_111_111
        // 1_111_111
        public string CheckFieldNumber(ScriptContext context, PlanSetup currentplan, string beamid)
        {
            Regex regex = new Regex(@"_[0-9]+$");
            Match match, match2;

            match2 = regex.Match(beamid);  // is the current beam a subfield?

            string out_string = "";

            foreach (Course c in context.Patient.Courses)
            {
                foreach (PlanSetup ps in c.PlanSetups)
                {
                    //if ((ps.Id != currentplan.Id))
                    //if (((ps.Id == currentplan.Id) & (((ps.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved) | (ps.ApprovalStatus == PlanSetupApprovalStatus.CompletedEarly) | (ps.ApprovalStatus == PlanSetupApprovalStatus.Retired)))) | ((ps.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved) | (ps.ApprovalStatus == PlanSetupApprovalStatus.CompletedEarly) | (ps.ApprovalStatus == PlanSetupApprovalStatus.Retired)))

                    //MessageBox.Show(ps.Id + '\n' + (ps.Id == currentplan.Id).ToString() + " " + (c.Id == context.Course.Id).ToString() + " " + (ps.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved).ToString() + " " + (ps.ApprovalStatus == PlanSetupApprovalStatus.CompletedEarly).ToString() + " " + (ps.ApprovalStatus == PlanSetupApprovalStatus.Retired).ToString());

                    if (((ps.Id == currentplan.Id) & (c.Id == context.Course.Id)) | ((ps.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved) | (ps.ApprovalStatus == PlanSetupApprovalStatus.CompletedEarly) | (ps.ApprovalStatus == PlanSetupApprovalStatus.Retired)))
                    {
                        foreach (Beam b in ps.Beams.Where(o => o.IsSetupField == false))
                        {
                            if (b.Id + ps.Id != beamid + currentplan.Id)  // had to add plan id for the case where the fields are named the same thing
                            {
                                char[] splitter = { ' ', '_', 'g' };
                                double n, n1;
                                string[] fieldname = b.Id.Split(splitter);
                                string[] currentbeam = beamid.Split(splitter);
                                if (double.TryParse(fieldname.First().ToString(), out n) | double.TryParse(currentbeam.First().ToString(), out n1)) // test if the first element in the field name is an integer
                                {
                                    //MessageBox.Show("Current Beam: " + beamid + '\n' + "Test beam: " + b.Id + '\n');
                                    match = regex.Match(b.Id);
                                    if (n == n1)
                                    {  // if the initial beam numbers match
                                        //if (match.Value.ToString() != "") // this indicates a split field, continue if it isn't _num
                                        //{
                                        if ((match.Value.ToString() == "") & (match2.Value.ToString() == ""))
                                        {// if either test or current are subfields {
                                            out_string = out_string + '\n' + "Field #" + n.ToString() + " already used in course: " + c.Id + " in plan: " + ps.Id + " in field: " + b.Id;
                                            //} // update the max number if the current beam is larger
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            return out_string;
        }

        //public string CheckFieldNumber(string beamid,int maxfieldnum)
        //{
        //    string out_string = "";
        //    char[] splitter = { ' ', '_', 'g' };
        //    int n;
        //    string[] fieldname = beamid.Split(splitter);
        //    if (int.TryParse(fieldname.First().ToString(), out n)) // test if the first element in the field name is an integer
        //    {
        //        if (n <= maxfieldnum)
        //        {
        //            out_string = "Field #" + n.ToString() + " already used in previous plan!";
        //        } // update the max number if the current beam is larger
        //    }
        //    return out_string;
        //}

        public string CheckMachines(ScriptContext context, PlanSetup currentplan)
        {
            string mch = "";
            foreach (Course c in context.Patient.Courses)
            {
                if (c == context.Course)
                {
                    foreach (PlanSetup p in c.PlanSetups)
                    {
                        if (p == currentplan) break;
                        foreach (Beam b in p.Beams)
                        {
                            if (b.TreatmentUnit.Id != currentplan.Beams.First().TreatmentUnit.Id)
                            {
                                mch = mch + "Plan: " + p.Id + ", Beam: " + b.Id + ", uses Machine: " + b.TreatmentUnit.Id + '\n';
                            }
                        }
                    }
                }
            }
            return mch;
        }

        public DataTable AriaSQL(string query)
        {
            string connectionString = "Data Source=varwcdcpdbavm01;Persist Security Info=True;Password=reports;User ID=reports;Initial Catalog=variansystem";
            DataTable dt = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                SqlDataReader reader = null;
                SqlCommand command = new SqlCommand(query, con);
                con.Open();
                reader = command.ExecuteReader();
                dt.Load(reader);
                //if (reader.HasRows)
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


        public string Add2BeamSum(string sumstring, string newstring)
        {
            sumstring += newstring + '\n';
            return sumstring;
        }

        public string Add2Log(string sumstring, string newstring)
        {
            sumstring += newstring + ";";
            return sumstring;
        }

        public Tuple<string,DataTable> FindRx(ScriptContext context, string RxStatus)
        {
            DataTable RxTable;

            String sql = "SELECT dbo.Prescription.PrescriptionSer, dbo.Prescription.PrescriptionName, dbo.PlanSetup.PlanSetupId, dbo.Course.CourseId, dbo.Patient.PatientId, dbo.Patient.LastName, dbo.Patient.FirstName FROM dbo.Prescription INNER JOIN dbo.PlanSetup ON dbo.Prescription.PrescriptionSer = dbo.PlanSetup.PrescriptionSer INNER JOIN dbo.Course ON dbo.PlanSetup.CourseSer = dbo.Course.CourseSer INNER JOIN dbo.Patient ON dbo.Course.PatientSer = dbo.Patient.PatientSer "
                + " WHERE PatientId = '" + context.Patient.Id + "' AND CourseId = '" + context.Course.Id + "' AND dbo.PlanSetup.PlanSetupId = '" + context.PlanSetup.Id + "'" ;

            RxTable = AriaSQL(sql);

            sql = "SELECT DISTINCT dbo.Course.CourseId, dbo.Prescription.PrescriptionName, dbo.Prescription.Site, dbo.Prescription.PhaseType, CASE dbo.Prescription.SimulationNeeded WHEN 1 THEN 'YES' WHEN 0 THEN 'NO' END AS[CT Simulation], dbo.Prescription.Status, dbo.Prescription.HstryUserName AS[Approved By], dbo.Prescription.HstryDateTime AS[Approval Date], "
                    + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'DOSE PER FRACTION' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS DosePerFx, dbo.Prescription.NumberOfFractions, "
                    + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'TOTAL DOSE' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS TotalDose, "
                    + "(SELECT TOP 1 ItemType + '= ' + ItemValue + CASE WHEN ItemValueUnit IS NULL THEN '' ELSE ItemValueUnit END AS [ItemValue] FROM dbo.PrescriptionAnatomyItem PAI WHERE (PAI.ItemType = 'VOLUME ID' OR PAI.ItemType = 'DEPTH' OR PAI.ItemType = 'ISODOSE PERCENTAGE') AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS Volume,"
                    + "substring((SELECT ',' + PropertyValue AS[text()] FROM dbo.PrescriptionProperty PP Where PP.PropertyType = 1 AND PP.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Energies, "
                    + "substring((SELECT ',' + PP2.PropertyValue AS [text()] FROM dbo.PrescriptionProperty PP2 WHERE PP2.PropertyType = 2 AND PP2.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Mode, dbo.Prescription.Technique, dbo.Prescription.Notes FROM dbo.PrescriptionProperty RIGHT OUTER JOIN dbo.Patient INNER JOIN dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer INNER JOIN dbo.TreatmentPhase ON dbo.Course.CourseSer = dbo.TreatmentPhase.CourseSer INNER JOIN dbo.Prescription ON dbo.TreatmentPhase.TreatmentPhaseSer = dbo.Prescription.TreatmentPhaseSer INNER JOIN dbo.PrescriptionAnatomy ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionAnatomy.PrescriptionSer INNER JOIN dbo.PrescriptionAnatomyItem ON dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer ON dbo.PrescriptionProperty.PrescriptionSer = dbo.Prescription.PrescriptionSer ";

            if (RxTable != null && RxTable.Rows.Count !=0)
            {
                sql = sql + " WHERE dbo.Prescription.PrescriptionSer = " + RxTable.Rows[0]["PrescriptionSer"].ToString();
                return new Tuple<string,DataTable>("Linked Prescription Found!",AriaSQL(sql));
            }
            else {
                sql = "SELECT DISTINCT dbo.Course.CourseId, dbo.Prescription.PrescriptionName, dbo.Prescription.Site, dbo.Prescription.PhaseType, CASE dbo.Prescription.SimulationNeeded WHEN 1 THEN 'YES' WHEN 0 THEN 'NO' END AS[CT Simulation], dbo.Prescription.Status, dbo.Prescription.HstryUserName AS[Approved By], dbo.Prescription.HstryDateTime AS[Approval Date], "
                        + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'DOSE PER FRACTION' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS DosePerFx, dbo.Prescription.NumberOfFractions, "
                        + "(SELECT CAST(ItemValue * 100.00 AS Integer) AS ItemValue FROM dbo.PrescriptionAnatomyItem PAI WHERE PAI.ItemType = 'TOTAL DOSE' AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS TotalDose, "
                        + "(SELECT TOP 1 ItemType + '= ' + ItemValue + CASE WHEN ItemValueUnit IS NULL THEN '' ELSE ItemValueUnit END AS [ItemValue] FROM dbo.PrescriptionAnatomyItem PAI WHERE (PAI.ItemType = 'VOLUME ID' OR PAI.ItemType = 'DEPTH' OR PAI.ItemType = 'ISODOSE PERCENTAGE') AND PAI.PrescriptionAnatomySer = dbo.PrescriptionAnatomy.PrescriptionAnatomySer) AS Volume,"
                        + "substring((SELECT ',' + PropertyValue AS[text()] FROM dbo.PrescriptionProperty PP Where PP.PropertyType = 1 AND PP.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Energies, "
                        + "substring((SELECT ',' + PP2.PropertyValue AS [text()] FROM dbo.PrescriptionProperty PP2 WHERE PP2.PropertyType = 2 AND PP2.PrescriptionSer = dbo.Prescription.PrescriptionSer For XML Path('')),2,1000) AS Mode, dbo.Prescription.Technique, dbo.Prescription.Notes FROM dbo.PrescriptionProperty RIGHT OUTER JOIN dbo.Patient INNER JOIN dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer INNER JOIN dbo.TreatmentPhase ON dbo.Course.CourseSer = dbo.TreatmentPhase.CourseSer INNER JOIN dbo.Prescription ON dbo.TreatmentPhase.TreatmentPhaseSer = dbo.Prescription.TreatmentPhaseSer INNER JOIN dbo.PrescriptionAnatomy ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionAnatomy.PrescriptionSer INNER JOIN dbo.PrescriptionAnatomyItem ON dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer ON dbo.PrescriptionProperty.PrescriptionSer = dbo.Prescription.PrescriptionSer ";

                if (RxStatus == "Approved")
                {
                    // CHECK database for Rx, create a table for the Rxs
                    sql = sql + "WHERE (dbo.PrescriptionAnatomy.AnatomyRole = 2) AND (dbo.Patient.PatientId = '" + context.Patient.Id + "') AND Status = 'Approved' AND CourseId = '" + context.Course.Id + "' "
                        + "ORDER BY dbo.Prescription.HstryDateTime";
                }
                else
                {
                    // CHECK database for Rx, create a table for the Rxs
                    sql = sql + "WHERE (dbo.PrescriptionAnatomy.AnatomyRole = 2) AND (dbo.Patient.PatientId = '" + context.Patient.Id + "') AND Status = 'Draft' AND CourseId = '" + context.Course.Id + "' "
                        + "ORDER BY dbo.Prescription.HstryDateTime";
                }
                RxTable = AriaSQL(sql);
            }


            //// THIS IS REALLY USEFUL TO PRINT OUT A DATATABLE TO STRING
            System.Text.StringBuilder RxList = new System.Text.StringBuilder();
            foreach (System.Data.DataRow r in RxTable.Rows)
            {
                RxList.Append(r["PrescriptionName"].ToString() + '\n');
                foreach (DataColumn c in RxTable.Columns)
                {
                    RxList.Append(c.ColumnName.ToString() + ": " + r[c.ColumnName].ToString() + "   " + '\n');
                }
                RxList.Append('\n');
            }

            // CHECK if the plan has a "J" in it if it is going to be a junction
            // match static fields that contain "g", check gantry angle against name
            Regex planregex = new Regex(@"J[0-9]+");
            Match planmatch;
            planmatch = planregex.Match(context.PlanSetup.Id);  // this searches if there is a J[0-9]
            Boolean junctionplan = false;
            if (!string.IsNullOrEmpty(planmatch.Value))  // if the field ID CONTAINS g[0-9]
            {
                junctionplan = true;
            }

            // Now match the current plan to the prescription list by dose per fx, total dose, junction in note
            // RETURNS DATAROW
            // 1st try to match dose/fx, ecluding anything that mentions junctions 
            var temp_dt = RxTable.AsEnumerable()
            .Where(r => r.Field<Int32>("DosePerFx").ToString().Equals(Convert.ToInt32(context.PlanSetup.UniqueFractionation.PrescribedDosePerFraction.Dose).ToString())
            && r.Field<Int32>("NumberOfFractions").ToString().Equals(Convert.ToInt32(context.PlanSetup.UniqueFractionation.NumberOfFractions).ToString())
            && !(r.Field<String>("Notes").ToLower().Contains("junction")) && !(r.Field<String>("Notes").ToLower().Contains("jxn"))
            );

            //RxList.Clear();
            //foreach (System.Data.DataRow r in temp_dt.CopyToDataTable().Rows)
            //{
            //    RxList.Append(r["PrescriptionName"].ToString() + '\n');
            //    foreach (DataColumn c in temp_dt.CopyToDataTable().Columns)
            //    {
            //        RxList.Append(c.ColumnName.ToString() + ": " + r[c.ColumnName].ToString() + "   " + '\n');
            //    }
            //    RxList.Append('\n');
            //}
            //MessageBox.Show(RxList.ToString());

            try
            {

                if (temp_dt.Count() == 0) // check if th enotes have any junction or jxn
                {
                    //MessageBox.Show("No match on Dose per Fraction and Number of fractions!" + '\n' + "Looking for junction in notes.");
                    temp_dt = RxTable.AsEnumerable()
                    .Where(r => r.Field<Int32>("DosePerFx").ToString().Equals(Convert.ToInt32(context.PlanSetup.UniqueFractionation.PrescribedDosePerFraction.Dose).ToString())
                    && r.Field<Int32>("NumberOfFractions").ToString().Equals(Convert.ToInt32(context.PlanSetup.UniqueFractionation.NumberOfFractions).ToString())
                    || (r.Field<String>("Notes").ToLower().Contains("junction")) || (r.Field<String>("Notes").ToLower().Contains("jxn"))
                    );
                }

                if (junctionplan && temp_dt.Count() == 0) // still no match :(
                {
                    //MessageBox.Show("No jxn or junction in Rx comment." + '\n' + "Trying to match DosePerFx.");
                    temp_dt = RxTable.AsEnumerable()
                    .Where(r => r.Field<Int32>("DosePerFx").ToString().Equals(Convert.ToInt32(context.PlanSetup.UniqueFractionation.PrescribedDosePerFraction.Dose).ToString())
                    );
                }
                if (temp_dt.Count() > 0)
                {
                    DataTable FoundRx = temp_dt.CopyToDataTable();
                    return new Tuple<string, DataTable>("No Linked Prescription Found!", FoundRx);
                }
                else
                {
                    return new Tuple<string, DataTable>("No Prescription Found!", new DataTable()); 
                }

            }
            catch
            {
                return new Tuple<string, DataTable>("No Prescription Found!", new DataTable());
            }
        }

        #endregion

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {

            if (context.Patient == null)
            {
                MessageBox.Show("No patient selected!");
                return;
            }

            if ((context.PlanSumsInScope == null) & (context.PlansInScope == null) & (context.PlanSetup == null))
            {
                MessageBox.Show("No plans selected!");
                return;
            }

            // check if the plan is properly approved
            if ((context.PlanSetup.ApprovalStatus != PlanSetupApprovalStatus.PlanningApproved) & (context.PlanSetup.ApprovalStatus != PlanSetupApprovalStatus.TreatmentApproved))
            {
                MessageBox.Show("Warning!" + '\n' + context.PlanSetup.Id + " is " + context.PlanSetup.ApprovalStatus.ToString());
            }

            //MessageBox.Show(context.PlanSetup.ProtocolID.First().ToString() + '\n' + context.PlanSetup.ProtocolPhaseID.First().ToString());

            if (context.PlansInScope != null) { CheckPlan(context, context.PlanSetup, context.Image); }

            //if ((context.PlanSumsInScope != null))
            //{
            //    foreach (PlanSum ps in context.PlanSumsInScope)
            //    {
            //        foreach (PlanSetup planInCheck in ps.PlanSetups)
            //        {
            //            CheckPlan(context,planInCheck, context.Image);
            //        }

            //    }
            //}                

            WriteLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + context.CurrentUser.Name + "," + context.Patient.Id + "," + context.PlanSetup.Id + "," + context.PlanSetup.ApprovalStatus + "," + ErrorLog);
        }

        public void CheckPlan(ScriptContext context, PlanSetup plan, Image image)
        {

            System.Text.StringBuilder RxList = new System.Text.StringBuilder();
            DataTable FoundRx;

            var tuple = FindRx(context, "Approved");

            FoundRx = tuple.Item2;

            if (FoundRx == null) { MessageBox.Show("Checking Draft Rxs"); FoundRx = FindRx(context, "Draft").Item2; }
            if (FoundRx == null) { MessageBox.Show("still can't find Rx"); }
            try
            {
                var planE = context.PlanSetup.Beams.Where(p => p.IsSetupField == false).Select(o => o.EnergyModeDisplayName).Distinct().ToArray();
                if (FoundRx.Rows.Count != 0)
                {
                    foreach (DataRow r in FoundRx.Rows)
                    {
                        string[] foundE = r["Energies"].ToString().Split(',');
                        var diffE = planE.Union(foundE).Except(planE.Intersect(foundE)); // this finds the non-intersecting
                        string modeOut = "";
                        if (r["Energies"].ToString().Contains("E") && !r["Mode"].ToString().Contains("Electron")) { modeOut = "   ~~~~~ ALERT ~~~~~  Mode missing Electron!" + '\n'; }
                        if (r["Energies"].ToString().Contains("X") && !r["Mode"].ToString().Contains("Photon")) { modeOut = "   ~~~~~ ALERT ~~~~~  Mode missing Photon!" + '\n'; }
                        RxList.Append(r["PrescriptionName"].ToString() + '\n');
                        foreach (DataColumn c in FoundRx.Columns)
                        {
                            RxList.Append(c.ColumnName.ToString() + ": " + r[c.ColumnName].ToString());
                            if (c.ColumnName == "Energies" && diffE.Count() != 0) { RxList.Append('\n' + "~~~~  ALERT ~~~~" + '\n' + "CHECK ENERGIES: " + string.Join(", ", diffE)); ErrorLog = Add2Log(ErrorLog, "CHECK ENERGIES: " + string.Join(", ", diffE)); }
                            if (c.ColumnName == "Mode" && modeOut != "") { RxList.Append(modeOut); }
                            if (c.ColumnName == "CT Simulation" && r[c.ColumnName].ToString() == "NO") { RxList.Append("~~~~ ALERT ~~~~" + '\n' + "RX HAS WRONG SIMULATION!!!"); ErrorLog = Add2Log(ErrorLog, "RX HAS WRONG SIMULATION!!!"); }
                            // dose per fraction here
                            // number of fractions here
                            RxList.Append('\n');
                        }
                        RxList.Append('\n');
                    }
                    MessageBox.Show(tuple.Item1 + '\n' + RxList.ToString(), tuple.Item1);
                }
                else { MessageBox.Show("Could not find matching Rx."); }
            }
            catch
            {
                MessageBox.Show("Unable to find matching Rx!" + '\n' + '\n' + "Continuing plan check");
            }
            // declare local variables that reference the objects we need.
            StructureSet ss = plan.StructureSet;
            double[] data = new double[] { };
            //ImageProfile hor = new ImageProfile(image.UserOrigin, new VVector(1, 1, 1), data, "mm");
            //MessageBox.Show(hor.Count.ToString());
            //MessageBox.Show(data.ToString());

            double[] origin = new double[] { image.Origin.x, image.Origin.x, image.Origin.x };

            Structure body = ss.Structures.First(bdy => bdy.Id.Trim().ToLower().Contains("body"));
            string couchmax = "";
            if (ss.Structures.Where(x => x.Id == "CouchSurface").Count() != 0) { couchmax = FindMaxCouch(plan, ss.Structures.First(x => x.Id == "CouchSurface")); }

            //DataTable couchDT = AriaSQL()

            MessageBox.Show(FindCTCouch(image, plan) + '\n' + '\n' + "Contour Based:" + '\n' + FindPortraitCouch(plan, body) + '\n' + couchmax, "Couch Height");// + FindPortraitCouch(image, body));

            List<VVector> Isos = new List<VVector> { };

            #region Image
            // Here we test whether the planning CT is from our sim and whether the correct CT density table was used

            if (plan.TreatmentOrientation != image.ImagingOrientation)
            {
                MessageBox.Show("Treatment orientation NOT Imaging orientation!" + '\n' + "Tx: " + plan.TreatmentOrientation.ToString() + '\n' + "Im: " + image.ImagingOrientation.ToString());
                ErrorLog = Add2Log(ErrorLog, "Treatment orientation NOT Imaging orientation!");
            }

            #endregion

            #region BEAMS
            // OKAY TIME TO GO INTO THE ARIA DATABASE!!!
            DataTable fieldDT = AriaSQL("SELECT dbo.PlanSetup.PlanSetupId, dbo.Radiation.RadiationId, dbo.ExternalFieldCommon.TreatmentTime, Grat.GraphicAnnotationId FROM dbo.PlanSetup INNER JOIN dbo.Course ON dbo.Course.CourseSer = dbo.PlanSetup.CourseSer INNER JOIN dbo.Patient ON dbo.Patient.PatientSer = dbo.Course.PatientSer LEFT OUTER JOIN dbo.Radiation ON dbo.Radiation.PlanSetupSer = dbo.PlanSetup.PlanSetupSer LEFT OUTER JOIN dbo.ExternalFieldCommon ON dbo.ExternalFieldCommon.RadiationSer = dbo.Radiation.RadiationSer LEFT OUTER JOIN (SELECT dbo.GraphicAnnotation.GraphicAnnotationId, dbo.Image.ImageSer FROM dbo.Image LEFT OUTER JOIN dbo.GraphicAnnotation ON dbo.GraphicAnnotation.ImageSer = dbo.Image.ImageSer WHERE(dbo.GraphicAnnotation.GraphicAnnotationId = 'Graticule')) AS Grat ON Grat.ImageSer = dbo.Radiation.RefImageSer "
                + "WHERE(dbo.Patient.PatientId = '" + context.Patient.Id + "') AND(dbo.PlanSetup.PlanSetupId = '" + plan.Id + "') AND(dbo.Course.CourseId = '" + plan.Course.Id + "')");

            List<BeamOverview> BeamSum = new List<BeamOverview>();
            List<string> BeamSummary = new List<string>();

            // Check whether bolus is applied to any fields
            bool bolus_present = false;
            if (plan.Beams.Where(p => p.IsSetupField == false).Where(p => p.Boluses.Count() > 0).Count() > 0) { bolus_present = true; }

            // check if jaw tracking has been turned on for VMAT planning. Not done by default
            if (plan.Beams.Where(p => p.MLCPlanType.ToString()=="VMAT").Count() > 0)
            {
                //plan.OptimizationSetup.Parameters.Where(p=> OptimizationJawTrackingUsedParameter.Equals(p.GetType()))
                //if (!plan.OptimizationSetup.UseJawTracking) { MessageBox.Show("VMAT plan not using jaw tracking in optimization!"); }
            }

            //MessageBox.Show(string.Join(",",plan.Beams.Select(o=>o.Id).ToArray()));
            foreach (Beam b in plan.Beams)
            {
                string beamsum = "Is okay";

                if ((b.MLCPlanType.ToString()=="DoseDynamic") | (b.MLCPlanType.ToString() == "VMAT"))
                {
                    // Check only if there are more than 1 set of jaw positions AND more than 10 control points. Rejects FiF
                    if ((b.ControlPoints.GroupBy(x=>x.JawPositions).Distinct().Count() == 1) & (b.ControlPoints.Count() > 10))
                    {
                        beamsum = Add2BeamSum(beamsum, "Jaw tracking not enabled!");
                    }
                    //// For jaw tracking plans, check the minimum jaw width
                    //if (b.ControlPoints.GroupBy(x => x.JawPositions).Distinct().Count() > 1)
                    //{
                    //}
                }

                // Get the Aria info for the current beam
                var fld = from DataRow myRow in fieldDT.Rows
                          where (string)myRow["RadiationId"] == b.Id
                          select myRow;

                // Engage these only after the plan has been planning or Tx aprpoved
                if ((plan.ApprovalStatus == PlanSetupApprovalStatus.PlanningApproved) | (plan.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved))
                {
                    // Check the treatment time is 2x
                    if (!fld.First().ItemArray[2].Equals(DBNull.Value))
                    {
                        if ((Math.Abs(Convert.ToDouble(fld.First().ItemArray[2].ToString()) - b.Meterset.Value / b.DoseRate * 2) > 0.1) & (b.MLCPlanType.ToString() != "VMAT"))
                        {
                            beamsum = Add2BeamSum(beamsum, "Treatment Time multiplied by: " + Math.Round(Convert.ToDouble(fld.First().ItemArray[2].ToString()) / b.Meterset.Value * b.DoseRate, 1).ToString());
                        }
                    }
                }

                // Check if the field has a graticule
                if ((fld.First().ItemArray[3].ToString() != "Graticule")) { beamsum = Add2BeamSum(beamsum, "Field is missing Graticule converted to Contour!"); }

                // Check if the field # has already been used
                string fieldnum = "";
                if (b.IsSetupField == false) { fieldnum = CheckFieldNumber(context, plan, b.Id); }

                if (fieldnum != "") { beamsum = Add2BeamSum(beamsum, fieldnum); ErrorLog = Add2Log(ErrorLog, fieldnum); }

                // Check if there is an applicator, and if so that the tray information exists
                if (b.Applicator != null)
                {
                    foreach (Block bk in b.Blocks)
                    {
                        //MessageBox.Show("ID: " + bk.AddOnMaterial.Id + '\n' + "Diverging: " + bk.IsDiverging + '\n' + "Transmission: " + bk.TransmissionFactor + '\n');
                        if (b.Trays.Count() == 0)
                        {
                            beamsum = Add2BeamSum(beamsum, "Field is missing e- tray!  This must be fixed to treat!");
                            ErrorLog = Add2Log(ErrorLog, "Field is missing e- tray!");
                            //MessageBox.Show("Field: " + b.Id + " is missing e- tray!" + '\n' + "This must be fixed to treat!");
                        }
                        else {

                            //MessageBox.Show(b.Trays.First().Id + "   " + b.Trays.First().Name + "   " + b.Blocks.First().Tray.Comment);
                        }
                    }
                }

                // Check that calculated and planned SSDs match
                if ((plan.ApprovalStatus == PlanSetupApprovalStatus.TreatmentApproved) | (plan.ApprovalStatus == PlanSetupApprovalStatus.PlanningApproved))
                {
                    if ((Math.Round(b.SSD, 1) / 10).ToString("F1") != (Math.Round(b.PlannedSSD, 1) / 10).ToString("F1"))
                    {
                        beamsum = Add2BeamSum(beamsum, "SSDs do not match! Planned: " + (Math.Round(b.PlannedSSD, 1) / 10).ToString("F1") + ", Calculated: " + (Math.Round(b.SSD, 1) / 10).ToString("F1"));
                        ErrorLog = Add2Log(ErrorLog, "SSDs do not match!");
                        //MessageBox.Show("SSDs do not match" + '\n' + "Field Id: " + b.Id + '\n' + "Planned: " + (Math.Round(b.PlannedSSD,1)/10).ToString("F1") + '\n' + "Calculated: " + (Math.Round(b.SSD,1)/10).ToString("F1"));
                    }
                }

                // Check that the dose rate is 600
                if (b.IsSetupField == false)
                {
                    switch (b.EnergyModeDisplayName)
                    {
                        case "6X-FFF":
                            if (b.DoseRate != 1400)
                            {
                                beamsum = Add2BeamSum(beamsum, "Field does not have DR = 1400!");
                                ErrorLog = Add2Log(ErrorLog, "DR!=1400");
                            }
                            break;
                        case "10X-FFF":
                            if (b.DoseRate != 2400)
                            {
                                beamsum = Add2BeamSum(beamsum, "Field does not have DR = 2400!");
                                ErrorLog = Add2Log(ErrorLog, "DR!=2400");
                            }
                            break;
                        default:
                            if (b.DoseRate != 600)
                            {
                                beamsum = Add2BeamSum(beamsum, "Field does not have DR = 600!");
                                ErrorLog = Add2Log(ErrorLog, "DR!=600");
                            }
                            break;
                    }
                }

                // Check if bolus is applied to current field
                if (b.IsSetupField == false)
                {
                    if (bolus_present)
                    {
                        if (b.Boluses.Count() == 0)
                        {
                            beamsum = Add2BeamSum(beamsum, "Field is missing Bolus!"); ErrorLog = Add2Log(ErrorLog, "Bolus missing");
                        }
                    }
                }

                //if (b.IsSetupField == true)
                //{

                //}

                // BEAM NAME CHECK!!!
                // VMAT SPECIFIC
                string value = b.MLCPlanType.ToString();
                if (value == "VMAT")
                {
                    VMAT = true;
                    string gantry = Math.Round(b.ControlPoints[0].GantryAngle, 0) + "-" + Math.Round(b.ControlPoints[b.ControlPoints.Count - 1].GantryAngle, 0);
                    if (!b.Id.Replace(" ", "").Contains(gantry))
                    {
                        beamsum = Add2BeamSum(beamsum, "Field does not contain gantry start/stop in name!  " + '\n' + "Should contain: g" + gantry);
                        ErrorLog = Add2Log(ErrorLog, "No gantry start/stop in name.");
                        //MessageBox.Show("Field: " + b.Id + " does not contain gantry start/stop in name!  " + '\n' + "Needs to contain: g" + gantry);
                    }
                    if ((b.ControlPoints[0].CollimatorAngle == 0) | (b.ControlPoints[0].CollimatorAngle == 90))
                    {
                        beamsum = Add2BeamSum(beamsum, "Field has collimator angle=" + b.ControlPoints[0].CollimatorAngle.ToString() + "\n" + "Consider adjustment!!");
                        ErrorLog = Add2Log(ErrorLog, "VMAT collimator");
                        //MessageBox.Show("Field: " + b.Id + " has collimator angle=" + b.ControlPoints[0].CollimatorAngle.ToString() + "\n" + "Consider adjustment!!");
                    }
                }

                // IF NOT VMAT THEN RUN THE FOLLOWING CHECKS
                if (value != "VMAT")
                {
                    // CHECK DRRS

                    if (b.ReferenceImage == null)
                    {
                        beamsum = Add2BeamSum(beamsum, "Field does not have a reference image (DRR)!");
                        ErrorLog = Add2Log(ErrorLog, "No DRR");
                        //MessageBox.Show(b.Id + " does not have a reference image (DRR)!");
                    }

                    if (b.ReferenceImage != null)
                    {
                        if (b.Id.Length > 12)
                        {
                            try
                            {
                                int ref_len = b.Id.Length - 1;
                                if (b.ReferenceImage.Id.Substring(0, ref_len).Contains(b.Id.Substring(0, ref_len)) == false)
                                {
                                    beamsum = Add2BeamSum(beamsum, "DRR mislabeled!  " + b.ReferenceImage.Id); ErrorLog = Add2Log(ErrorLog, "DRR mislabeled");
                                }
                            }
                            catch
                            {
                                beamsum = Add2BeamSum(beamsum, "DRR mislabeled!  " + b.ReferenceImage.Id); ErrorLog = Add2Log(ErrorLog, "DRR mislabeled");
                            }
                        }
                        else if (b.ReferenceImage.Id.Contains(b.Id) == false)
                        {
                            beamsum = Add2BeamSum(beamsum, "DRR mislabeled!  " + b.ReferenceImage.Id); ErrorLog = Add2Log(ErrorLog, "DRR mislabeled");
                        }
                    }

                    // Check minimum MU for EDW
                    if (b.Wedges.Count() != 0)
                    {
                        if ((b.Meterset.Value < 21) & (b.Wedges.First().Id.Contains("EDW")))
                        {
                            beamsum = Add2BeamSum(beamsum, b.Wedges.First().Id + " has " + Math.Round(b.Meterset.Value, 1).ToString() + " " + b.Meterset.Unit.ToString() + ", which is less than the minimum 22 MU!"); ErrorLog = Add2Log(ErrorLog, "EDW MU");
                        }
                    }
                    // Check if the gantry is 180 for right sided
                    if ((Math.Round(b.ControlPoints[0].GantryAngle, 0) == 180) & (b.Technique.Id == "STATIC"))
                    {
                        VVector v = b.IsocenterPosition;

                        if ((v.x < -30) & ((plan.TreatmentOrientation == PatientOrientation.HeadFirstSupine) | (plan.TreatmentOrientation == PatientOrientation.FeetFirstProne)))
                        {
                            if (b.ControlPoints[0].GantryAngle < 180.1)
                            {
                                beamsum = Add2BeamSum(beamsum, "Gantry angle should be >= 180.1"); ErrorLog = Add2Log(ErrorLog, "G180.1");
                            }
                        }
                        if ((v.x > 30) & ((plan.TreatmentOrientation == PatientOrientation.HeadFirstSupine) | (plan.TreatmentOrientation == PatientOrientation.FeetFirstProne)))
                        {
                            if (b.ControlPoints[0].GantryAngle > 180.0)
                            {
                                beamsum = Add2BeamSum(beamsum, "Gantry angle should be <= 180.0"); ErrorLog = Add2Log(ErrorLog, "G<180.0");
                            }
                        }
                        if ((v.x > 30) & ((plan.TreatmentOrientation == PatientOrientation.HeadFirstProne) | (plan.TreatmentOrientation == PatientOrientation.FeetFirstSupine)))
                        {
                            if (b.ControlPoints[0].GantryAngle < 180.1)
                            {
                                beamsum = Add2BeamSum(beamsum, "Gantry angle should be >= 180.1"); ErrorLog = Add2Log(ErrorLog, "G>=180.1");
                            }
                        }
                        if ((v.x < -30) & ((plan.TreatmentOrientation == PatientOrientation.HeadFirstProne) | (plan.TreatmentOrientation == PatientOrientation.FeetFirstSupine)))
                        {
                            if (b.ControlPoints[0].GantryAngle > 180.0)
                            {
                                beamsum = Add2BeamSum(beamsum, "Gantry angle should be <= 180.0"); ErrorLog = Add2Log(ErrorLog, "G<=180.0");
                            }
                        }
                    }
                }

                Isos.Add(b.IsocenterPosition);
                // match static fields that contain "g", check gantry angle against name
                Regex regex = new Regex(@"g?\d+(\.\d+)?$"); //Regex(@"g[0-9]+");
                //Regex regex = new Regex(@"g[0-9]+");
                Match match;
                match = regex.Match(b.Id.ToLower());  // this searches if there is a gantry g[0-9] 

                if (!string.IsNullOrEmpty(match.Value))  // if the field ID CONTAINS g[0-9]
                {
                    if ((value != "VMAT") & (!b.IsSetupField))
                    {
                        if ((!match.Value.Contains("g" + b.ControlPoints[0].GantryAngle.ToString())) | (match.Value.Length != ("g" + b.ControlPoints[0].GantryAngle.ToString()).Length))
                        {
                            //MessageBox.Show("First Condition: " + (!match.Value.Contains("g" + b.ControlPoints[0].GantryAngle.ToString())) + '\n' + match.Value.ToString() + "   " + match.Value.Length.ToString() + '\n' + "Second Condition: " + (match.Value.Length != ("g" + b.ControlPoints[0].GantryAngle.ToString()).Length) + '\n' + "g" + b.ControlPoints[0].GantryAngle.ToString() + "   " + '\n' + ("g" + b.ControlPoints[0].GantryAngle.ToString()).Length.ToString());
                            
                            beamsum = Add2BeamSum(beamsum, "Field name should contain: g" + b.ControlPoints[0].GantryAngle.ToString()); ErrorLog = Add2Log(ErrorLog, "FieldName Gantry");
                            //MessageBox.Show("Field: " + b.Id + "\n" + "Needs to be: g" + b.ControlPoints[0].GantryAngle.ToString());
                            //if (b.Id.Contains("g"))
                            //b.ControlPoints[0].GantryAngle
                        }
                    }
                    if ((Math.Round(b.ControlPoints[0].PatientSupportAngle, 0) != 0) & (Math.Round(b.ControlPoints[0].PatientSupportAngle, 0) != 360))
                    {
                        if (b.Id.ToLower().Contains("g"))
                        {
                            if (!b.Id.ToLower().Contains("c"))
                            {
                                beamsum = Add2BeamSum(beamsum, "Field name has label g but no c for couch. Should contain: c" + Math.Round(360 - b.ControlPoints[0].PatientSupportAngle, 0)); ErrorLog = Add2Log(ErrorLog, "FieldName Couch");
                                //MessageBox.Show("Field: " + b.Id + " has label g but no c for couch." + '\n' + "Needs to be: c" + Math.Round(360 - b.ControlPoints[0].PatientSupportAngle, 0));
                            }
                            else if (!b.Id.ToLower().Contains("c" + Math.Round(360 - b.ControlPoints[0].PatientSupportAngle, 0)))
                            {
                                beamsum = Add2BeamSum(beamsum, "Field name shows incorrect couch angle. Should contain: c" + Math.Round(360 - b.ControlPoints[0].PatientSupportAngle, 0)); ErrorLog = Add2Log(ErrorLog, "FieldName Couch");
                                //MessageBox.Show("Field: " + b.Id + " shows incorrect couch angle." + '\n' + "Needs to be: c" + Math.Round(360 - b.ControlPoints[0].PatientSupportAngle, 0));
                            }

                        }
                    }

                    //// if there is a g180 field, check to see if it should be 180.1 or 180.0
                    //if (Math.Round(b.ControlPoints[0].GantryAngle, 0) == 180)
                    //{
                    //    VVector v = image.UserOrigin - b.IsocenterPosition;
                    //    if (v.x > 30)
                    //    {
                    //        if (b.ControlPoints[0].GantryAngle < 180.1)
                    //        {
                    //            beamsum = Add2BeamSum(beamsum, "Gantry angle should be >= 180.1");
                    //            //MessageBox.Show("Gantry angle for field: " + b.Id + " should be >= 180.1");
                    //        }
                    //    }
                    //    if (v.x < -30)
                    //    {
                    //        if (b.ControlPoints[0].GantryAngle > 180.0)
                    //        {
                    //            beamsum = Add2BeamSum(beamsum, "Gantry angle should be <= 180.0");
                    //            //MessageBox.Show("Gantry angle for field: " + b.Id + " should be <= 180.0");
                    //        }
                    //    }
                    //}
                }
                // CHECK ANATOMICALLY DEFINED FIELDS
                // If doesn't contain gXXX
                // check text, switch case, orientation, gantry angle, etc.
                List<string> Anatomical = new List<string> { "rl", "rtlat", "rlat", "rtla", "rightlat", "ll", "ltlat", "ltla", "llat", "leftlat", "rao", "rpo", "lpo", "lao", "ap", "pa" };
                if ((string.IsNullOrEmpty(match.Value)) & (Anatomical.Any(b.Id.Replace(" ", "").Replace("_", "").ToLower().Contains)))  // if field does NOT contain g[0-9]
                {
                    string AnatomicFld = CheckAnatomicFieldName(plan.TreatmentOrientation.ToString(), b.ControlPoints.First().GantryAngle, b.Id);
                    if (AnatomicFld != "")
                    {
                        beamsum = Add2BeamSum(beamsum, AnatomicFld); ErrorLog = Add2Log(ErrorLog, "FieldName anatomy");
                    }
                }



                //// check setup fields
                //if (b.IsSetupField)
                //{
                //    string setupmsg = CheckAnatomicFieldName(plan.TreatmentOrientation.ToString(), b.ControlPoints.First().GantryAngle, b.Id);
                //    if (!string.IsNullOrEmpty(setupmsg))
                //    {
                //        beamsum = Add2BeamSum(beamsum, setupmsg);
                //        //MessageBox.Show(setupmsg);
                //    }
                //    if (b.ReferenceImage == null)
                //    {
                //        beamsum = Add2BeamSum(beamsum, "Field does not have a reference image (DRR)!");
                //        //MessageBox.Show(b.Id + " does not have a reference image (DRR)!");
                //    }
                //}

                if (beamsum.Trim().Length > 7) { beamsum = beamsum.Trim().Replace("Is okay", "").Trim(); }
                BeamSummary.Add(beamsum);
                if (beamsum.Trim() != "Is okay")
                {
                    //MessageBox.Show("check step, is okay " + b.Id + '\n' + beamsum);
                    BeamOverview temp = new BeamOverview();
                    temp.Name = b.Id;
                    temp.Issues = beamsum;
                    BeamSum.Add(temp);
                }

            }
            string beam_output = "";
            foreach (BeamOverview b in BeamSum)
            {
                beam_output += "Field: " + b.Name + '\n' + b.Issues + '\n' + '\n';
            }
            if (BeamSum.Count == 0)
            {
                MessageBox.Show("No Field issues detected!", "Field Issues");
            }
            else
            {
                MessageBox.Show(beam_output, "Field Issues");
            }


            #endregion

            #region Plans, formerly isocenters and shifts

            // check machines
            string mch = "";
            foreach (Beam b in plan.Beams)
            {
                if (b.TreatmentUnit.Id != plan.Beams.First().TreatmentUnit.Id)
                {
                    mch = mch + "Field: " + b.Id + " uses machine: " + b.TreatmentUnit.Id + ", field:" + plan.Beams.First().Id + " uses machine: " + plan.Beams.First().TreatmentUnit.Id + '\n';
                }
            }
            mch = mch + CheckMachines(context, plan);
            if (mch != "") { MessageBox.Show(plan.Course.Id + " plans" + '\n' + mch, "Machine Mismatch - Current Machine: " + plan.Beams.First().TreatmentUnit.Id); ErrorLog = Add2Log(ErrorLog, "Machine Mismatch"); }

            // check machine schedule
            bool BID = false;
            DataTable mch_sched = AriaSQL("SELECT DISTINCT dbo.Machine.MachineId, dbo.Activity.ActivityCode, dbo.ScheduledActivity.ScheduledStartTime FROM dbo.Patient INNER JOIN dbo.ScheduledActivity ON dbo.Patient.PatientSer = dbo.ScheduledActivity.PatientSer INNER JOIN dbo.ActivityInstance ON dbo.ScheduledActivity.ActivityInstanceSer = dbo.ActivityInstance.ActivityInstanceSer INNER JOIN dbo.ResourceActivity ON dbo.ScheduledActivity.ScheduledActivitySer = dbo.ResourceActivity.ScheduledActivitySer INNER JOIN dbo.Machine ON dbo.ResourceActivity.ResourceSer = dbo.Machine.ResourceSer INNER JOIN dbo.Activity ON dbo.ActivityInstance.ActivitySer = dbo.Activity.ActivitySer WHERE(dbo.Patient.PatientId = '" + context.Patient.Id + "') AND(dbo.ScheduledActivity.ObjectStatus = 'Active') AND(dbo.ActivityInstance.ObjectStatus = 'Active') AND (dbo.Activity.ActivityCategorySer = 69) AND dbo.ScheduledActivity.ScheduledActivityCode = 'Open' AND ScheduledStartTime > GETDATE() ORDER BY ScheduledStartTime");
            string schedError = "";
            DateTime temp_date = DateTime.Now.AddMonths(5).Date;
            foreach (DataRow dr in mch_sched.Rows)
            {

                if (temp_date == DateTime.Now.AddMonths(5).Date) // if this is the first time being run, just grab the first date and move on
                    temp_date = Convert.ToDateTime(dr[2].ToString()).Date;
                else
                if (temp_date == Convert.ToDateTime(dr[2].ToString()).Date) { BID = true; } // check if date is the same

                if (dr[0].ToString() != context.PlanSetup.Beams.First().TreatmentUnit.Id)
                {
                    schedError += dr[0].ToString() + "  " + dr[1].ToString() + "  " + Convert.ToDateTime(dr[2].ToString()).ToString("M/d/yyyy hh:mm tt") + '\n';
                }
            }
            if (schedError != "") { MessageBox.Show("Patient is scheduled for different machine on following days: " + '\n' + schedError, "Machine Schedule Mismatch!"); ErrorLog = Add2Log(ErrorLog, "Machine Schedule Mismatch"); }


            // check shifts  // TODO !!!!!!!!!!! Include setupfields in the reporting, but IGNORE for multi iso.
            //var TxIsos = plan.Beams.Where(o => o.IsSetupField == false).Select(o => o.IsocenterPosition).Distinct();
            var TxIsos = plan.Beams.Select(o => o.IsocenterPosition).Distinct();
            string iso_shift = "";
            // string multi = "";
            // if (TxIsos.Count() > 1) { multi = " Mult. Isos.  Needs YELLOW sticker!"; }
            for (int i = 0; i < TxIsos.Count(); i++)
            {
                VVector c = TxIsos.ElementAt(i);
                if (Vec2Str(VecRound(c, 2)) != Vec2Str(VecRound(image.UserOrigin, 2)))
                {

                    // if (i == 0) { iso_shift += "Shift Exists!  Needs GREEN sticker!" + "   " + multi + '\n' + '\n'; } // This used to be for the stickers

                    iso_shift += "Iso Shift #: " + (i + 1).ToString() + '\n';
                    iso_shift += "Fields: " + string.Join(", ", plan.Beams.Where(o => o.IsocenterPosition.Equals(c)).Select(o => o.Id).ToArray()) + '\n';
                    iso_shift += "Shift from origin (cm): " + ReturnShiftDirections(plan.TreatmentOrientation.ToString(), VecRound((c - image.UserOrigin) / 10, 2)) + '\n';
                    if (i != 0) { iso_shift += "Shift from Iso #1 (cm): " + ReturnShiftDirections(plan.TreatmentOrientation.ToString(), VecRound((c - TxIsos.ElementAt(0)) / 10, 2)); }
                    iso_shift = iso_shift + " " + '\n' + '\n' + "Plan Comment:" + '\n' + plan.Comment;

                    if (plan.Comment == "") { iso_shift = iso_shift + '\n' + '\n' + "NEEDS A PLAN COMMENT!!!!"; ErrorLog = Add2Log(ErrorLog, "Plan Comment"); }
                }
            }
            if (iso_shift == "")
            {
                MessageBox.Show("No shifts from triangulation point", "Shifts");
            }
            else {
                MessageBox.Show(iso_shift, "Iso shift");
            }

            #endregion

            #region STRUCTURES

            string ssum = ""; // report string

            // check for couch in VMAT plan
            if (VMAT == true)
            {
                var couch = ss.Structures.Where(c => c.Id.Contains("Couch"));
                if (couch.Count() == 0) { ssum = Add2BeamSum(ssum, "Couch is MISSING in VMAT plan"); }
                couch = ss.Structures.Where(c => c.Id.Contains("CouchRail"));
                if (couch.Count() == 1) { ssum = Add2BeamSum(ssum, "Only 1 Couch Rail found! - " + couch.First().Id); }
                if (couch.Count() == 0) { ssum = Add2BeamSum(ssum, "Missing Couch Rails!"); }
            }

            if (VMAT == false)
            {
                var couch = ss.Structures.Where(c => c.Id.Contains("Couch"));
                if (couch.Count() == 0) { ssum = Add2BeamSum(ssum, "No Couch Structures were found!" + '\n' + '\n' + "Check immobilization is contoured." + '\n' + "No couch needed for inserts." + '\n' + "Check insert included in body structure."); }
            }

            if (ssum != "")
            {
                MessageBox.Show(ssum, "Couch");
            }
            ssum = "";

            List<string> emptyStructs = new List<string>();
            char[] sp = new char[] { ',', ' ' };
            List<RV_Struct> PlanStructs = new List<RV_Struct> { };

            // GET THE BODY CONTOUR for testing center.x
            //Structure body = ss.Structures.First(bdy => bdy.Id.ToLower().Contains("body"));

            // start structure loop
            foreach (Structure s in ss.Structures)
            {


                RV_Struct tempS = new RV_Struct();
                tempS.InitSlice();
                tempS.Name = s.Id;
                tempS.Target = false;
                tempS.HasLaterality = false;

                #region Laterality
                // Check laterality
                char[] lat_splitter = { ' ', '_' };
                double lat_thresh = 15;
                List<string> LT_LATER = new List<string> { "l", "lt", "left" };
                List<string> RT_LATER = new List<string> { "r", "rt", "right" };
                string[] laterality = s.Id.ToLower().Split(lat_splitter);
                //double shifty = s.CenterPoint.x;  // THESE CONTOURS ARE DEFINED BASED ON THE IMAGE DICOM!!!!  NOT ISO OR USERORIGIN
                double shifty = s.CenterPoint.x - body.CenterPoint.x; // This uses the X from the body contour
                if (LT_LATER.Any(laterality.Contains)) // THE LEFT SIDED CASE
                {
                    if ((shifty < lat_thresh)) // THIS NEXT PART IS UNNECESSARY, ECLIPSE HANDLES ORIENTATION & ((context.Image.ImagingOrientation == PatientOrientation.HeadFirstSupine) | (context.Image.ImagingOrientation == PatientOrientation.FeetFirstProne)))
                    {
                        // dataset.Where(i => !excluded.Any(e => i.Contains(e)));
                        //var setToRemove = new HashSet<Author>(authors);
                        //authorsList.RemoveAll(x => setToRemove.Contains(x));
                        tempS.HasLaterality = true;
                        tempS.Laterality = "Listed as Left, but should be Right  " + shifty.ToString("F1");

                    }
                }
                if (RT_LATER.Any(laterality.Contains))
                {
                    if ((shifty > -lat_thresh)) // & ((context.Image.ImagingOrientation == PatientOrientation.HeadFirstSupine) | (context.Image.ImagingOrientation == PatientOrientation.FeetFirstProne)))
                    {
                        string lat_test = "";
                        lat_test = string.Join(",", laterality.Where(i => !RT_LATER.Any(laterality.Contains)));
                        if (lat_test != "")
                        {
                            MessageBox.Show(string.Join(",", laterality.Where(i => !RT_LATER.Any(laterality.Contains))), "Laterality");
                            tempS.HasLaterality = true;
                            tempS.Laterality = "Listed as Right, but should be Left  " + shifty.ToString("F1");

                        }
                    }
                }
                #endregion

                if (s.IsEmpty) // check for empty structures
                {
                    tempS.Empty = true;
                    PlanStructs.Add(tempS);

                    emptyStructs.Add(s.Id);
                    ssum = Add2BeamSum(ssum, s.Id + " is empty!"); ErrorLog = Add2Log(ErrorLog, "Empty structure");
                    continue;
                }

                // define target volume
                if (s.Id == plan.TargetVolumeID)
                {
                    tempS.Target = true;
                    TargetStruct = s;
                }

                double t;
                s.GetAssignedHU(out t);
                tempS.Density = t;
                tempS.DensityOvrd = (s.GetAssignedHU(out t));



                //MessageBox.Show(s.MeshGeometry.Positions.ToString(), "MeshGeometry.Positions");
                //MessageBox.Show((s.CenterPoint.x - origin[0]).ToString() + " " + (s.CenterPoint.y - origin[1]).ToString() + " " + (s.CenterPoint.z - origin[2]).ToString(), "Center - origin");
                //MessageBox.Show(s.MeshGeometry.TextureCoordinates.ToString(), "MeshGeometry.TextureCoordinates");
                //MessageBox.Show("X: " + s.MeshGeometry.Bounds.X.ToString() + ", Y: " + s.MeshGeometry.Bounds.Y.ToString() + ", Z: " + s.MeshGeometry.Bounds.Z.ToString(), "MeshGeometry.Bounds");
                //MessageBox.Show(s.MeshGeometry.TriangleIndices.ToString(), "MeshGeometry.TriangleIndices");
                //MessageBox.Show(s.MeshGeometry.Bounds.Location.X.ToString() + " " + s.MeshGeometry.Bounds.Location.Y.ToString());

                #region Missing Slices
                int minSlice = ss.Image.ZSize; int maxSlice = 0;
                var list = new List<int> { };
                for (int i = 0; i < ss.Image.ZSize; i++)
                {
                    if (s.GetContoursOnImagePlane(i).Count() != 0)
                    {
                        list.Add(i);
                        if (i < minSlice) { minSlice = i; }
                        if (i > maxSlice) { maxSlice = i; }
                    }
                }

                tempS.MaxSlice = maxSlice;
                tempS.MinSlice = minSlice;
                tempS.SliceList = list;
                //var result = Enumerable.Range(minSlice, maxSlice).Except(list);
                tempS.MissingSliceList = Enumerable.Range(minSlice, maxSlice - minSlice).Except(list);
                if (tempS.MissingSliceList.Any())
                {
                    tempS.MissingSlices = true;
                }
                #endregion

                PlanStructs.Add(tempS);
            }

            #region Structure Reporting
            string rvs = "";
            if (PlanStructs.Where(c => c.MissingSlices == true).Count() > 0)
            {
                rvs = "MISSING SLICES" + '\n';
                foreach (RV_Struct rv in PlanStructs)
                {
                    if (rv.MissingSlices) { rvs += rv.Name + " missing slices at: " + string.Join(",", (rv.MissingSliceList.Select(x => Math.Round((double)x * context.Image.ZRes / 10 + (context.Image.Origin.z - context.Image.UserOrigin.z) / 10, 2))).ToArray()) + " cm" + '\n'; }
                }
                MessageBox.Show(rvs, "Structure Slices"); ErrorLog = Add2Log(ErrorLog, "Missing Slices");
            }

            rvs = "";
            if (PlanStructs.Where(c => c.HasLaterality == true).Count() > 0)
            {
                rvs = "LATERALITY ISSUES?" + '\n';
                foreach (RV_Struct rv in PlanStructs)
                {
                    if (rv.HasLaterality) { rvs += rv.Name + ": " + rv.Laterality + '\n'; }
                }
                MessageBox.Show(rvs, "Structure Laterality"); ErrorLog = Add2Log(ErrorLog, "Laterality");
            }
            //// compute the overlap of structures
            //string ssover = "";
            //foreach (Structure s in ss.Structures)
            //{
            //    float temp = CheckOverlap(s, TargetStruct);
            //    if (temp > 0)
            //    {
            //        ssover = ssover + s.Id + temp.ToString() + '\n';
            //    }
            //}
            //ssover = ssover.Substring(0, ssover.Length - 1);
            //MessageBox.Show("Contour Overlap with Target: " + TargetStruct.Id + '\n' + ssover);

            rvs = "";
            if (PlanStructs.Where(c => c.DensityOvrd == true).Count() > 0)
            {
                rvs = "DENSITY OVERRIDES" + '\n';
                foreach (RV_Struct rv in PlanStructs)
                {
                    if (rv.DensityOvrd) { rvs += rv.Name + "  HU: " + rv.Density.ToString() + '\n'; }
                }
                MessageBox.Show(rvs, "Structure Density");
            }
            #endregion

            #endregion

            #region DOSE

            // regex to find PTVs with numbers after them.  Use numbers after to define Rx dose
            // if only PTV exists (list contains only 1) then use the Rx dose as dose
            // cycle through and check V_Rx and V_95%Rx4



            //if (plan.Dose.DoseMax3DLocation
            //Boolean max_in_target = false;
            string doseCalc = "";
            string targDose = "";
            Structure targStruc;
            if (!string.IsNullOrEmpty(plan.TargetVolumeID.ToString()))
            {
                foreach (Structure st in ss.Structures)
                {
                    if (st.Id == plan.TargetVolumeID)
                    {
                        targStruc = st;
                        if ((targStruc.Volume < 25) & (plan.Dose.XRes > 2.0) & (plan.UniqueFractionation.PrescribedDosePerFraction.Dose > 400))
                        {
                            doseCalc = "WARNING!" + '\n' + "Target Structure: " + st.Id + " volume is " + Math.Round(st.Volume, 1).ToString() + " cc" + '\n' + "Dose Grid is: " + plan.Dose.XRes.ToString() + " mm.  " + "Dose/Fx = " + plan.UniqueFractionation.PrescribedDosePerFraction.Dose.ToString() + plan.UniqueFractionation.PrescribedDosePerFraction.UnitAsString + '\n' + '\n' + "Consider reducing dose grid resolution!";
                            ErrorLog = Add2Log(ErrorLog, "DoseGridSize");
                        }
                        targDose += st.Id + ": D_95%" + " = " + plan.GetDoseAtVolume(st, 95.00, VolumePresentation.Relative, DoseValuePresentation.Relative).ToString() + '\n';
                        targDose += st.Id + ": D_1cc" + " = " + plan.GetDoseAtVolume(st, 1.00, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative).ToString() + '\n';
                        targDose += st.Id + ": V_100%" + " = " + plan.GetVolumeAtDose(st, plan.TotalPrescribedDose, VolumePresentation.Relative).ToString("F1") + '\n';
                        targDose += st.Id + ": D_100%" + " = " + plan.GetDoseAtVolume(st, 100.00, VolumePresentation.Relative, DoseValuePresentation.Relative).ToString() + ", min dose" + '\n' + '\n';

                        targDose += body.Id + ": D_1cc" + " = " + plan.GetDoseAtVolume(body, 1.00, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative).ToString() + '\n' + '\n';
                        targDose += plan.PlanNormalizationMethod.ToString();
                        if (st.IsPointInsideSegment(plan.Dose.DoseMax3DLocation))
                        {
                            MessageBox.Show("Max dose is inside " + plan.TargetVolumeID + '\n' + targDose);
                        }
                        else
                        {
                            MessageBox.Show("Max dose is NOT inside " + plan.TargetVolumeID + '\n' + targDose);
                            ErrorLog = Add2Log(ErrorLog, "MaxDoseNotInTarget");
                        }
                    }

                }
            }

            if (plan.PhotonCalculationModel.Contains("AAA"))
            {
                string phetero = "";
                plan.PhotonCalculationOptions.TryGetValue("HeterogeneityCorrection", out phetero);
                if ((phetero != "ON") & (plan.Beams.Count(x => x.EnergyModeDisplayName.Contains("X")) > 0))
                {
                    doseCalc += '\n' + "WARNING!" + '\n' + "Heterogeneity Correction is OFF for photons!";
                    ErrorLog = Add2Log(ErrorLog, "DoseHeterogeneityOff");
                }
            }
            if (doseCalc != "") { MessageBox.Show(doseCalc, "Dose Calculation"); }

            // Use ARIA to check the total and session and daily dose limits
            string refpointsql = "SELECT dbo.RefPoint.RefPointId, dbo.DoseContribution.PrimaryFlag, dbo.RTPlan.NoFractions, dbo.DoseContribution.DosePerFraction, dbo.RefPoint.TotalDoseLimit, dbo.RefPoint.DailyDoseLimit, dbo.RefPoint.SessionDoseLimit FROM            dbo.RTPlan INNER JOIN dbo.DoseContribution ON dbo.RTPlan.RTPlanSer = dbo.DoseContribution.RTPlanSer INNER JOIN                            dbo.RefPoint ON dbo.DoseContribution.RefPointSer = dbo.RefPoint.RefPointSer INNER JOIN                            dbo.Patient ON dbo.RefPoint.PatientSer = dbo.Patient.PatientSer INNER JOIN                           dbo.PlanSetup ON dbo.RTPlan.PlanSetupSer = dbo.PlanSetup.PlanSetupSer INNER JOIN                           dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer AND dbo.PlanSetup.CourseSer = dbo.Course.CourseSer "
                + "WHERE(dbo.Patient.PatientId = '" + context.Patient.Id + "') AND(dbo.PlanSetup.PlanSetupId = '" + plan.Id + "') AND CourseId = '" + plan.Course.Id + "'";
            DataTable refpt = AriaSQL(refpointsql);

            refpt.PrimaryKey = new DataColumn[] { refpt.Columns["RefPointId"] };

            // Check if primary reference point is set to locationless point
            var PriRefPt = plan.Beams.Where(o => o.IsSetupField == false).First().FieldReferencePoints.Where(p => p.IsPrimaryReferencePoint == true).First();

            string PriRefPtMsg="";
            if (plan.IsDoseValid)
            {
                if (PriRefPt.RefPointLocation.x.ToString() != "NaN")
                {
                    PriRefPtMsg = ("Primary reference point not set to locationless point!" + '\n' + "Primary RefPt: " + PriRefPt.ReferencePoint.Id + '\n' + "RefPt Dose: ");
                    if (refpt.Rows.Find(PriRefPt.ReferencePoint.Id).Field<double>("DosePerFraction").ToString() != "")
                    {
                        PriRefPtMsg += (refpt.Rows.Find(PriRefPt.ReferencePoint.Id).Field<double>("DosePerFraction") * 100 * plan.UniqueFractionation.NumberOfFractions.Value).ToString() + " cGy";
                    }
                }
            }
            // Back to checking daily limits now

            if (refpt != null)
            {
                System.Text.StringBuilder RefPtList = new System.Text.StringBuilder();
                if (BID == true) { RefPtList.Append("Patient is scheduled BID!" + '\n' + '\n'); }

                foreach (DataRow r in refpt.Rows)
                {
                    try
                    {
                        if (r.Field<double>("DosePerFraction") > r.Field<double>("SessionDoseLimit")) { RefPtList.Append(r.Field<string>("RefPointId").ToString() + " is over SessionDoseLimit: " + (100.0 * r.Field<double>("DosePerFraction")).ToString("F3") + " > " + (100.0 * r.Field<double>("SessionDoseLimit")).ToString("F3") + '\n'); }
                        if (r.Field<double>("DosePerFraction") * 100.0 * Convert.ToDouble(r.Field<Int32>("NoFractions")) > r.Field<double>("TotalDoseLimit") * 100.0) { RefPtList.Append(r.Field<string>("RefPointId").ToString() + " is over TotalDoseLimit: " + (r.Field<double>("DosePerFraction") * 100.0 * Convert.ToDouble(r.Field<Int32>("NoFractions"))).ToString("F2") + " > " + (100.0 * r.Field<double>("TotalDoseLimit")).ToString("F2") + '\n'); }
                        if (BID)
                        {
                            if (r.Field<double>("DosePerFraction") * 2.0 > r.Field<double>("DailyDoseLimit")) { RefPtList.Append(r.Field<string>("RefPointId").ToString() + " is over DailyDoseLimit: " + (r.Field<double>("DosePerFraction") * 2.0).ToString("F3") + " > " + r.Field<double>("DailyDoseLimit").ToString("F3") + '\n'); }
                        }
                        else
                        {
                            if (r.Field<double>("DosePerFraction") > r.Field<double>("DailyDoseLimit")) { RefPtList.Append(r.Field<string>("RefPointId").ToString() + " is over DailyDoseLimit: " + (100.0 * r.Field<double>("DosePerFraction")).ToString("F3") + " > " + (100.0 * r.Field<double>("DailyDoseLimit")).ToString("F3") + '\n'); }
                        }
                    }
                    catch
                    {
                        RefPtList.Append(r.Field<string>("RefPointId").ToString() + " is missing some dose limits." + '\n');
                    }
                }
                if (RefPtList.ToString() != "" | PriRefPtMsg!="") { MessageBox.Show(RefPtList.ToString() + '\n' + '\n' + PriRefPtMsg, "Reference Points"); }
            }
            else { MessageBox.Show("No reference points have dose limits yet." + '\n' + "Ref point doses not checked." + '\n' + PriRefPtMsg, "Reference Points"); }



            //This is a nice way to get the keys and values from a dictionary!!!
            //MessageBox.Show(string.Join(",", plan.ElectronCalculationOptions.ToArray()));

            //DataTable RxDT = AriaSQL("SELECT DISTINCT dbo.Course.CourseId, dbo.Prescription.SimulationNeeded, dbo.Prescription.Notes, dbo.Prescription.Status, dbo.Prescription.NumberOfFractions, dbo.Prescription.Technique, dbo.Prescription.Site, dbo.Prescription.PrescriptionSer, dbo.PrescriptionAnatomy.AnatomyName, dbo.PrescriptionAnatomy.AnatomyRole, dbo.PrescriptionAnatomyItem.ItemType, dbo.PrescriptionAnatomyItem.ItemValue, dbo.PrescriptionAnatomyItem.ItemValueUnit FROM dbo.Patient INNER JOIN dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer INNER JOIN dbo.TreatmentPhase ON dbo.Course.CourseSer = dbo.TreatmentPhase.CourseSer INNER JOIN dbo.Prescription ON dbo.TreatmentPhase.TreatmentPhaseSer = dbo.Prescription.TreatmentPhaseSer INNER JOIN dbo.PrescriptionAnatomy ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionAnatomy.PrescriptionSer INNER JOIN dbo.PrescriptionAnatomyItem ON dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer INNER JOIN dbo.PrescriptionProperty ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionProperty.PrescriptionSer WHERE(dbo.Patient.PatientId = '" + context.Patient.Id + "')");
            //var rows = from row in RxDT.AsEnumerable()
            //           where row.Field<string>("ItemType") == "TOTAL DOSE"
            //           select row;
            //plan.Beams.Any(o=>o.EnergyModeDisplayName)
            //rows.
            //var RxEnergy = AriaSQL("SELECT DISTINCT dbo.PrescriptionProperty.PropertyValue FROM dbo.Patient INNER JOIN dbo.Course ON dbo.Patient.PatientSer = dbo.Course.PatientSer INNER JOIN dbo.TreatmentPhase ON dbo.Course.CourseSer = dbo.TreatmentPhase.CourseSer INNER JOIN dbo.Prescription ON dbo.TreatmentPhase.TreatmentPhaseSer = dbo.Prescription.TreatmentPhaseSer INNER JOIN dbo.PrescriptionAnatomy ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionAnatomy.PrescriptionSer INNER JOIN dbo.PrescriptionAnatomyItem ON dbo.PrescriptionAnatomy.PrescriptionAnatomySer = dbo.PrescriptionAnatomyItem.PrescriptionAnatomySer INNER JOIN dbo.PrescriptionProperty ON dbo.Prescription.PrescriptionSer = dbo.PrescriptionProperty.PrescriptionSer WHERE(dbo.Patient.PatientId = '5240781') AND(dbo.PrescriptionProperty.PropertyType = 1) AND dbo.Prescription.PrescriptionSer = 23837 ").AsEnumerable().ToList();


            #endregion
        }

        // 'listStructures' if of type IEnumerable<Structure>

        //// loop through structure list and find the PTV
        //Structure ptv = null;
        //foreach (Structure scan in listStructures)
        //{
        //    if (scan.Id == "PTV")
        //    {
        //        ptv = scan;
        //    }
        //}

        //// extract DVH data for PTV using bin width of 0.1.
        //DVHData dvh = plan.GetDVHCumulativeData(ptv, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);

        //string filename = @"c:\temp\keranen_dvh.csv";
        //System.IO.StreamWriter dvhFile = new System.IO.StreamWriter(filename);
        //// write a header
        //dvhFile.WriteLine("Dose,Volume");
        //// write all dvh points for the PTV.
        //foreach (DVHPoint pt in dvh.CurveData)
        //{
        //    string line = string.Format("{0},{1}", pt.DoseValue.Dose, pt.Volume);
        //    dvhFile.WriteLine(line);
        //}
        //dvhFile.Close();

        #region Logging
        void WriteLog(string strComments)
        {
            OpenFile();
            streamWriter.WriteLine(strComments);
            CloseFile();
        }

        void OpenFile()
        {
            DateTime dt = DateTime.Now;

            if (fi.Exists)
            {
                fileStream = new FileStream(fi.FullName, FileMode.Append, FileAccess.Write);
            }
            else {
                fileStream = new FileStream(fi.FullName, FileMode.Create, FileAccess.Write);
            }
            streamWriter = new StreamWriter(fileStream);
        }

        void CloseFile()
        {
            streamWriter.Close();
            fileStream.Close();
            GC.Collect();
        }
        #endregion
    }
}