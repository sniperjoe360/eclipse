using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Text;

namespace VMS.TPS


{
class Script
	{
		
		public void Execute(ScriptContext context /*, System.Windows.Window window*/)
    {
      if (context.Patient == null || context.PlanSetup == null || 
          context.PlanSetup.Dose == null || context.StructureSet == null)
        MessageBox.Show("Please load a patient, structure set, and a plan that has dose calculated!");

      // get reference to selected plan
      PlanSetup plan = context.PlanSetup;

      // linq query finds the first PTV structure
      Structure target = (from s in context.StructureSet.Structures where 
                          s.DicomType == "PTV" select s).FirstOrDefault();

      if (target == null) 
        throw new ApplicationException("Plan '"+plan.Id+"' has no PTV!");

      // build exported DVH filename, put it in the users temp directory
      string temp = System.Environment.GetEnvironmentVariable("TEMP");
      string dvhFilePath = temp + @"\dvh.csv";

      // export DVH for 'target' to 'dvhFilePath'
      exportDVH(plan, target, dvhFilePath);

      // 'Start' generated CSV file to launch Excel window
      System.Diagnostics.Process.Start(dvhFilePath);
      // Sleep for a few seconds to let Excel window start
      System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
    }
    private void exportDVH(PlanSetup plan, Structure target, string fileName)
    {
      // extract DVH data
      DVHData dvhData = plan.GetDVHCumulativeData(target,
                                    DoseValuePresentation.Relative,
                                    VolumePresentation.AbsoluteCm3,  0.1);

      if (dvhData == null)
        throw new ApplicationException("No DVH data for target '"+target.Id+"'.");

      // export DVH data as a CSV file
      using (System.IO.StreamWriter sw = 
                        new System.IO.StreamWriter(fileName, false, Encoding.ASCII))
      {
        // write the header, assume the first dvh point represents the others
        DVHPoint rep = dvhData.CurveData[0];
        sw.WriteLine(
          string.Format("Relative dose [{0}],Structure volume [{1}],", 
            rep.DoseValue.UnitAsString, rep.VolumeUnit)
        );
        // write each row of dose / volume data
        foreach (DVHPoint pt in dvhData.CurveData)
          sw.WriteLine(string.Format("{0:0.0},{1:0.00000}", 
                                pt.DoseValue.Dose, pt.Volume));
      }
    }
	}
}


// {
    // public class Execute(ScriptContext context /*, System.Windows.Window window*/)
    // {
      // if (context.Patient == null || context.PlanSetup == null || 
          // context.PlanSetup.Dose == null || context.StructureSet == null)
        // MessageBox.Show("Please load a patient, structure set, and a plan that has dose calculated!");

      // // get reference to selected plan
      // PlanSetup plan = context.PlanSetup;

      // // linq query finds the first PTV structure
      // Structure target = (from s in context.StructureSet.Structures where 
                          // s.DicomType == "PTV" select s).FirstOrDefault();

      // if (target == null) 
        // throw new ApplicationException("Plan '"+plan.Id+"' has no PTV!");

      // // build exported DVH filename, put it in the users temp directory
      // string temp = System.Environment.GetEnvironmentVariable("TEMP");
      // string dvhFilePath = temp + @"\dvh.csv";

      // // export DVH for 'target' to 'dvhFilePath'
      // exportDVH(plan, target, dvhFilePath);

      // // 'Start' generated CSV file to launch Excel window
      // System.Diagnostics.Process.Start(dvhFilePath);
      // // Sleep for a few seconds to let Excel window start
      // System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
    // }

       // public void Execute(ScriptContext context /*, System.Windows.Window window*/)
    // {
      // if (context.Patient == null || context.PlanSetup == null || 
          // context.PlanSetup.Dose == null || context.StructureSet == null)
        // MessageBox.Show("Please load a patient, structure set, and a plan that has dose calculated!");

      // // get reference to selected plan
      // PlanSetup plan = context.PlanSetup;

      // // linq query finds the first PTV structure
      // Structure target = (from s in context.StructureSet.Structures where 
                          // s.DicomType == "PTV" select s).FirstOrDefault();

      // if (target == null) 
        // throw new ApplicationException("Plan '"+plan.Id+"' has no PTV!");

      // // build exported DVH filename, put it in the users temp directory
      // string temp = System.Environment.GetEnvironmentVariable("TEMP");
      // string dvhFilePath = temp + @"\dvh.csv";

      // // export DVH for 'target' to 'dvhFilePath'
      // exportDVH(plan, target, dvhFilePath);

      // // 'Start' generated CSV file to launch Excel window
      // System.Diagnostics.Process.Start(dvhFilePath);
      // // Sleep for a few seconds to let Excel window start
      // System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
    // }
  // }
    // //public class TableForm