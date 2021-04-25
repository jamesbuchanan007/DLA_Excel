namespace DLA_to_Excel
{
    public class FSC
    {
        public string FSG { get; set; }
        public string FSC_CODE { get; set; }
        public string FSC_TITLE { get; set; }
        public string FSC_NOTES { get; set; }
        public string FSC_INCLUDE { get; set; }
        public string FSC_EXCLUDE { get; set; }
        public string FSG_TITLE { get; set; }
        public string FSG_NOTES { get; set; }

        public void Clear() {
            FSG = "";
            FSC_CODE = "";
            FSC_TITLE = "";
            FSC_NOTES = "";
            FSC_INCLUDE = "";
            FSC_EXCLUDE = "";
            FSG_TITLE = "";
            FSG_NOTES = "";
        }
    }
}