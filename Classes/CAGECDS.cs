namespace DLA_to_Excel
{
    public class CAGECDS
    {
        public class CAGE_DATA
        {
            public string CAGE_CD_9250 { get; set; }
            public string ADRS_NM_LN_NO_1087 { get; set; }
            public string ADRS_NM_C_TXT_1086 { get; set; }
        }

        public class ADDRESS
        {
            public string TEL_NBR_1356;
            public string CAGE_CD_9250 { get; set; }
            public string ST_ADRS_1_1082 { get; set; }
            public string ST_ADRS_2_1083 { get; set; }
            public string POBOX_1361 { get; set; }
            public string CITY_1084 { get; set; }
            public string ST_US_POSN_AB_0186 { get; set; }
            public string ZIP_CD_4400 { get; set; }
            public string CNTRY_1085 { get; set; }
        }

        public class CAO_ADP_POINT_CODES
        {
            public string CAGE_CD_9250 { get; set; }
            public string CAO_CD_8870 { get; set; }
            public string ADP_PNT_CD_8835 { get; set; }
        }

        public class OTHER_CODES
        {
            public string CAGE_CD_9250 { get; set; }
            public string CAGE_STAT_CD_2694 { get; set; }
            public string ASSOC_CD_CAGE_8855 { get; set; }
            public string CAGE_TYP_CD_4238 { get; set; }
            public string TYPE_CAGE_AFFL_0250 { get; set; }
            public string SZ_OF_BUS_CD_1364 { get; set; }
            public string PR_BUS_CAT_CD_1365 { get; set; }
            public string TYPE_BUS_CD_1366 { get; set; }
            public string WMN_OWND_BUS_1367 { get; set; }
        }

        public class STANDARD_INDUSTRIAL_CLASSIFICATION
        {
            public string CAGE_CD_9250 { get; set; }
            public string STD_IND_CL_CD_1368 { get; set; }
        }

        public class REPLACEMENT
        {
            public string CAGE_CD_9250 { get; set; }
            public string RPLM_CAGE_CD_3595 { get; set; }
        }

        public class FORMER_DATA
        {
            public string CAGE_CD_9250 { get; set; }
            public string ADRS_NM_LN_NO_1087 { get; set; }
            public string ADRS_NM_C_TXT_1086 { get; set; }
        }
    }
}