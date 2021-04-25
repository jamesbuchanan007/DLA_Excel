namespace DLA_to_Excel
{
    public class FCAGE
    {
        public string CAGE_CODE { get; set; }
        public string COMPANY_NAME_1 { get; set; }
        public string COMPANY_NAME_2 { get; set; }
        public string COMPANY_NAME_3 { get; set; }
        public string COMPANY_NAME_4 { get; set; }
        public string COMPANY_NAME_5 { get; set; }
        public string DOMESTIC_STREET_ADDRESS_1 { get; set; }
        public string DOMESTIC_STREET_ADDRESS_2 { get; set; }
        public string DOMESTIC_POST_OFFICE_BOX { get; set; }
        public string DOMESTIC_CITY { get; set; }
        public string DOMESTIC_STATE { get; set; }
        public string DOMESTIC_ZIP_CODE { get; set; }
        public string DOMESTIC_COUNTRY { get; set; }
        public string DOMESTIC_PHONE_FAX_NUMBER_1 { get; set; }
        public string DOMESTIC_PHONE_FAX_NUMBER_2 { get; set; }
        public string FOREIGN_STREET_ADDRESS_1 { get; set; }
        public string FOREIGN_STREET_ADDRESS_2 { get; set; }
        public string FOREIGN_POST_OFFICE_BOX { get; set; }
        public string FOREIGN_CITY { get; set; }
        public string FOREIGN_PROVINCE { get; set; }
        public string FOREIGN_COUNTRY { get; set; }
        public string FOREIGN_POSTAL_ZONE { get; set; }
        public string FOREIGN_PHONE_NUMBER { get; set; }
        public string FOREIGN_FAX_NUMBER { get; set; }
        public string CAO_CODE { get; set; }
        public string ADP_CODE { get; set; }
        public string STATUS_CODE { get; set; }
        public string ASSOCIATION_CODE { get; set; }
        public string TYPE_CODE { get; set; }
        public string AFFILIATION_CODE { get; set; }
        public string SIZE_OF_BUSINESS_CODE { get; set; }
        public string PRIMARY_BUSINESS_CATEGORY { get; set; }
        public string TYPE_OF_BUSINESS_CODE { get; set; }
        public string WOMAN_OWNED_BUSINESS { get; set; }
        public string STANDARD_INDUSTRIAL_1 { get; set; }
        public string STANDARD_INDUSTRIAL_2 { get; set; }
        public string STANDARD_INDUSTRIAL_3 { get; set; }
        public string STANDARD_INDUSTRIAL_4 { get; set; }
        public string REPLACEMENT_CAGE { get; set; }
        public string FORMER_NAME_1 { get; set; }
        public string FORMER_NAME_2 { get; set; }

        public void ClearForeign() {
            FOREIGN_STREET_ADDRESS_1 = "";
            FOREIGN_STREET_ADDRESS_2 = "";
            FOREIGN_POST_OFFICE_BOX = "";
            FOREIGN_CITY = "";
            FOREIGN_PROVINCE = "";
            FOREIGN_COUNTRY = "";
            FOREIGN_POSTAL_ZONE = "";
            FOREIGN_PHONE_NUMBER = "";
            FOREIGN_FAX_NUMBER = "";
        }

        public void ClearDomestic() {
            DOMESTIC_STREET_ADDRESS_1 = "";
            DOMESTIC_STREET_ADDRESS_2 = "";
            DOMESTIC_POST_OFFICE_BOX = "";
            DOMESTIC_CITY = "";
            DOMESTIC_STATE = "";
            DOMESTIC_ZIP_CODE = "";
            DOMESTIC_COUNTRY = "";
            DOMESTIC_PHONE_FAX_NUMBER_1 = "";
            DOMESTIC_PHONE_FAX_NUMBER_2 = "";
        }
    }
}