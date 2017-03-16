////////////////////////////////////////////////////////////////////////////////
// ExportBatchDVHs.cs
//
//  A ESAPI v11+ script that demonstrates Batch DVH export.
//
// Copyright (c) 2015 Varian Medical Systems, Inc.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy 
// of this software and associated documentation files (the "Software"), to deal 
// in the Software without restriction, including without limitation the rights 
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
// copies of the Software, and to permit persons to whom the Software is 
// furnished to do so, subject to the following conditions:
//
//  The above copyright notice and this permission notice shall be included in 
//  all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
// THE SOFTWARE.
////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS
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