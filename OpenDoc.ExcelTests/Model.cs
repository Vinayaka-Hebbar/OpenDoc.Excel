using System.Collections.Generic;

namespace OpenDoc.ExcelTests
{
    public class ParameterDataXBarRangeViewModel
    {
        public List<XBarRangeViewModel> XBarRangeDataList { get; set; }
        public ParameterComputeViewModel ParameterComputeData { get; set; }
    }

    //Plotting GRpah
    public class XBarRangeViewModel
    {
        public float GroupID { get; set; } //==> X Axisis
        public double Average { get; set; } //==> Y Axis of forst Chart
        public double Max { get; set; }
        public double Min { get; set; }
        public double Range { get; set; }//==> Y Axis of sec Chart
    }


    //to display as Table
    public class ParameterComputeViewModel
    {
        public double Average { get; set; }
        public double Range { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double UCL { get; set; }
        public double CL { get; set; }
        public double LCL { get; set; }
        public double RBar { get; set; }
        public double USL { get; set; }
        public double Target { get; set; }
        public double LSL { get; set; }
        public double CP { get; set; }
        public double CPU { get; set; }
        public double CPL { get; set; }
        public double CPK { get; set; }
        public double PP { get; set; }
        public double PPU { get; set; }
        public double PPL { get; set; }
        public double PPK { get; set; }
        //public double Average { get; set; }
    }
}
