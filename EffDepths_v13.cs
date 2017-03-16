using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Text.RegularExpressions;

namespace VMS.TPS

{
    public class Script
    {
        #region Variables and functions

        public Boolean VMAT;

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

        public string CheckReferencePoints(PlanSetup p)
        {
            String t = "";
            String refpoints = "";
            string[,] RParray = new string[p.Beams.Count(), p.Beams.First().FieldReferencePoints.Count()];
            //var refpts = p.Beams.First().FieldReferencePoints.Where(o => !double.IsNaN(o.EffectiveDepth));
            var refpts = p.Beams.Where(o=>o.IsSetupField==false).First().FieldReferencePoints.Where(o => o.RefPointLocation.x.ToString() != "NaN");
            //MessageBox.Show(refpts.Count().ToString() + '\n' + p.Beams.First().Id);
            //string[] rpts = p.Beams.Where(o=>o.IsSe tupField==false).First().FieldReferencePoints.Select(o => o.Id).ToArray();
            //MessageBox.Show(string.Join(",", refpts.ToList().ToString()));
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            Regex regex = new Regex(@"[0-9]+");

            char[] d = { ' ', '_' };
            
            foreach (var rp in refpts)
            {
                t = rp.ReferencePoint.Id.Replace("+","");
                //t = Regex.Replace(rp.ReferencePoint.Id, @"[^\w\d]", " ");
                //MessageBox.Show(t);
                refpoints += t;
                if (rp.IsPrimaryReferencePoint)
                {
                    refpoints += " (primary)";
                }
                refpoints += '\n';

                //foreach (Beam b in p.Beams.Where(o=>o.IsSetupField==false).OrderBy(s=>Convert.ToInt32(regex.Match(s.Id.Split(d).First()).Value.ToString())))
                foreach (Beam b in p.Beams.Where(o => o.IsSetupField == false).OrderBy(s => Convert.ToInt32(regex.Match(s.Id.Split(d).First()).Value.ToString())))
                    {
                    //MessageBox.Show(t);
                    //MessageBox.Show(b.Id.Split(d).First().ToString());
                    //MessageBox.Show(b.Id + '\n' + refpoints);
                    var frp = b.FieldReferencePoints.Where(q => q.ReferencePoint.Id == rp.ReferencePoint.Id).First();
                    //MessageBox.Show(rp.ReferencePoint.Id + '\n' + b.Id);
                    refpoints += "Beam: " + b.Id + ", CAX SSD: " + (b.SSD / 10).ToString("F1") + ", PSSD: " + (frp.SSD / 10).ToString("F1") + ", Eff. Depth: " + (Math.Round(frp.EffectiveDepth, 1) / 10).ToString("F1") + ", Dose: " + frp.FieldDose + '\n';
                }
                refpoints += '\n';
            }

            //MessageBox.Show(string.Join(",",refpts.ToString()));
            //int m, n;
            //m = 0; n = 0;
            
            //foreach (FieldReferencePoint rp in p.Beams.First().FieldReferencePoints)
            //{
            //    foreach (Beam b in p.Beams)
            //    {
            //        FieldReferencePoint[] temp_rp = b.FieldReferencePoints.AsEnumerable();
            //        sb.Append(rp.Id + temp_rp.)
            //    }
            //}

            //foreach (Beam b in p.Beams.OrderBy(s => s.Id))
            //{
            //    n = 0;
            //    foreach (var bb in b.FieldReferencePoints)
            //    {
            //        if (bb.RefPointLocation.x.ToString() != "NaN")
            //        {
            //            sb.Append((Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F2"));
            //            RParray[m, n] = (Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F2");
            //            //refpoints += "Id: " + bb.ReferencePoint.Id + ", IsPrimary: " + bb.IsPrimaryReferencePoint + ", EffDepth: " + (Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F1") + ", Dose: " + bb.FieldDose + '\n';
            //            //refpoints += "Beam: " + b.Id + ", Id: " + bb.ReferencePoint.Id + ", x: " + (bb.SSD/10).ToString() + ", EffDepth: " + (Math.Round(bb.EffectiveDepth, 1) / 10).ToString("F1") + ", Dose: " + bb.FieldDose + '\n';
            //            n = n + 1;
            //        }
            //    }
            //    m = m + 1;
            //}
            //string out_string = "BEAM_ID, ";
            //for (int i = 0; i < p.Beams.First().FieldReferencePoints.Count(); i++)
            //{
            //    out_string += p.Beams.First().FieldReferencePoints.ElementAt(i).ReferencePoint.Id + ", ";
            //}
            //out_string += '\n';
            //for (int i = 0; i < p.Beams.Count(); i++)
            //{
            //    out_string += p.Beams.ElementAt(i).Id + ": ";
            //    for (int j = 0; j < p.Beams.First().FieldReferencePoints.Count(); j++)
            //    {
            //        out_string += RParray[i, j] + ", ";
            //    }
            //    out_string += '\n';
            //}
            //MessageBox.Show(out_string);
            return refpoints;
            //return out_string;

            

        }

        #endregion

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            if (context.Patient == null)
            {
                MessageBox.Show("No patient selected!");
            }

            // declare local variables that reference the objects we need.
            PlanSetup plan = context.PlanSetup;

            MessageBox.Show(CheckReferencePoints(plan));

        }
    }
    //public class TableForm
    //{


    //}
}