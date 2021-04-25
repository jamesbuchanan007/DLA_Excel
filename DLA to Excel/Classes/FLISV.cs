namespace DLA_to_Excel
{
    public class FLISV
    {
        public string NIIN { get; set; }
        public string MRC { get; set; }
        public string MC { get; set; }
        public string CODE_CLEAR_REPLY { get; set; }

        public void Clear() {
            NIIN = "";
            MRC = "";
            MC = "";
            CODE_CLEAR_REPLY = "";
        }


    }
}