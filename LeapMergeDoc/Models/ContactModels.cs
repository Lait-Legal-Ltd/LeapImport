namespace LeapMergeDoc.Models
{
    public class ContactExcelData
    {
        public string? CardName { get; set; }  // Main name field
        public string? Email { get; set; }
        public DateTime? DateOfBirth { get; set; }
        public string? BuildingName { get; set; }
        public string? StreetLevel { get; set; }
        public string? Number { get; set; }
        public string? Street { get; set; }
        public string? TownCity { get; set; }
        public string? County { get; set; }
        public string? Postcode { get; set; }
        public string? Country { get; set; }
        public string? Phone { get; set; }
        public string? Home { get; set; }
        public string? Work { get; set; }
        public string? Mobile { get; set; }
        public string? Fax { get; set; }
        public string? POBoxInstructions { get; set; }
        public string? POBoxType { get; set; }
        public string? POBoxNumber { get; set; }
        public string? POBoxTownCity { get; set; }
        public string? POBoxCounty { get; set; }
        public string? POBoxPostcode { get; set; }
        public string? DxInstructions { get; set; }
        public string? DxNumber { get; set; }
        public string? Exchange { get; set; }
        public string? MktConsent { get; set; }
    }

    public class ProcessedContactData
    {
        public ContactExcelData? OriginalData { get; set; }
        public bool IsCompany { get; set; }
        public bool IsExistingClient { get; set; }  // Skip if true

        // Parsed name fields
        public string? GivenNames { get; set; }
        public string? LastName { get; set; }
        public string? CompanyName { get; set; }

        // Address combined
        public string? AddressLine1 { get; set; }
        public string? AddressLine2 { get; set; }

        // Contact type
        public int FkContactTypeId { get; set; }  // 1=Company, 2=Personal
    }
}
