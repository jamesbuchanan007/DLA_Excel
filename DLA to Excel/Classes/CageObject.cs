using Microsoft.Office.Interop.Excel;

namespace DLA_to_Excel
{
    public class CageObject
    {
        public Application App { get; set; }
        public Workbook xBook { get; set; }
        public Worksheet xCageSheet { get; set; }
        public Worksheet xAddressSheet { get; set; }
        public Worksheet xCaoSheet { get; set; }
        public Worksheet xOtherCodes { get; set; }
        public Worksheet xStrdSheet { get; set; }
        public Worksheet xReplSheet { get; set; }
        public Worksheet xFormerSheet { get; set; }
        public int cageRow { get; set; }
        public int addressRow { get; set; }
        public int caoRow { get; set; }
        public int otherCodesRow { get; set; }
        public int strdRow { get; set; }
        public int replRow { get; set; }
        public int formerRow { get; set; }
        public CAGECDS.CAGE_DATA cageM { get; set; }
        public CAGECDS.ADDRESS addressM { get; set; }
        public CAGECDS.CAO_ADP_POINT_CODES caoM { get; set; }
        public CAGECDS.OTHER_CODES otherM { get; set; }
        public CAGECDS.STANDARD_INDUSTRIAL_CLASSIFICATION strdM { get; set; }
        public CAGECDS.REPLACEMENT replM { get; set; }
        public CAGECDS.FORMER_DATA formerM { get; set; }
    }
}