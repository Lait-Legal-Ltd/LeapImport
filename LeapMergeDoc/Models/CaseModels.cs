namespace LeapMergeDoc.Models
{
    public class ExcelRowData
    {
        public string? Title { get; set; }
        public string? Initials { get; set; }
        public string? Forename { get; set; }  // Maps to GivenNames
        public string? Surname { get; set; }   // Maps to LastNames (Company name if no Forename)
        public string? GivenNames { get; set; }
        public string? LastNames { get; set; }
        public string? ShortName { get; set; }
        public string? ClientName { get; set; }
        public string? FirstEmailAddress { get; set; }
        public DateTime? DateOfBirth { get; set; }
        public string? House { get; set; }     // Maps to AddressLine1
        public string? Area { get; set; }      // Maps to AddressLine2
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
        public bool? MktConsent { get; set; }

        public string FullAddress => string.Join(", ", new[] { House, Area, BuildingName, Number, Street, StreetLevel, TownCity, County, Postcode, Country }
            .Where(s => !string.IsNullOrWhiteSpace(s)));

        public string PrimaryContactNumber => new[] { Mobile, Phone, Home, Work }
            .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";

        // For the new format: House = AddressLine1, Area = AddressLine2
        // Falls back to old format if House is empty
        public string AddressLine1 => !string.IsNullOrWhiteSpace(House) ? House : 
            string.Join(" ", new[] { BuildingName, Number, Street }.Where(s => !string.IsNullOrWhiteSpace(s)));

        public string AddressLine2 => !string.IsNullOrWhiteSpace(Area) ? Area : (StreetLevel ?? "");
    }

    public class CaseExcelData
    {
        public string? MatterNo { get; set; }  // The unique reference from Excel (used as CaseReferenceAuto)
        public int? CaseNumber { get; set; }   // Auto-generated, not from Excel
        public string? CaseName { get; set; }
        public DateTime? DateOpened { get; set; }
        public int? AreaOfPractice { get; set; }
        public int? CaseType { get; set; }
        public int? CaseSubType { get; set; }
        public int? PersonResponsible { get; set; }
        public int? PersonAssisting { get; set; }
        public int? PersonActing { get; set; }
        // Staff names from Excel (before lookup)
        public string? StaffRespName { get; set; }
        public string? StaffActName { get; set; }
        public string? StaffAssistName { get; set; }
        public string? CreditName { get; set; }  // Who brought the case
        public string? ClientFirstName { get; set; }
        public string? ClientLastName { get; set; }
        public string? ClientName { get; set; }
        public string? MatterType { get; set; }
        public string? MatterDescription { get; set; }
        public DateTime? ArchiveDate { get; set; }

        // New fields for LEAP format
        public string? ClientNo { get; set; }       // Client No column (e.g., SMI0001)
        public string? MatterNumber { get; set; }   // Matter column (e.g., 1, 2, 3)
        public string? FeeEarnerCode { get; set; }  // F/E column (fee earner code)
        public string? WorkId { get; set; }         // Work Id column (case type code)
        public string? Surname { get; set; }        // Surname for client lookup
        public string? Forename { get; set; }       // Forename for client lookup
    }

    public class ProcessedCaseData
    {
        public CaseExcelData? OriginalData { get; set; }
        public int FkBranchId { get; set; }
        public int? FkAreaOfPracticeId { get; set; }
        public int? FkCaseTypeId { get; set; }
        public int? FkCaseSubTypeId { get; set; }
        public string? CaseReferenceAuto { get; set; }
        public int? CaseNumber { get; set; }
        public string? CaseName { get; set; }
        public string? CaseNameWithClient { get; set; }
        public DateTime? DateOpened { get; set; }
        public int PersonOpened { get; set; }
        public int? PersonResponsible { get; set; }
        public int? PersonActing { get; set; }
        public int? PersonAssisting { get; set; }
        public int? CaseCredit { get; set; }  // Who brought the case
        public bool IsCaseActive { get; set; }
        public bool IsCaseArchived { get; set; }
        public bool IsCaseNotProceeding { get; set; }
        public bool MnlCheck { get; set; }
        public bool ConfSearch { get; set; }
        public int? LinkedClientId { get; set; }
        public string? ClientFullName { get; set; }
        
        // Contacts for company clients (not the main client)
        public List<CaseContactInfo> Contacts { get; set; } = new List<CaseContactInfo>();
    }

    public class ClientInfo
    {
        public int ClientId { get; set; }
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? Title { get; set; }
        public string? FullName { get; set; }
        public string? ClientType { get; set; }
        public string? CompanyName { get; set; }
    }

    /// <summary>
    /// Client master record from client CSV (Client No -> actual client info)
    /// Used to identify the real client for a case (vs contacts)
    /// </summary>
    public class ClientMasterRecord
    {
        public string? ClientNo { get; set; }    // e.g., SSN0001
        public string? Title { get; set; }
        public string? Initials { get; set; }
        public string? Forename { get; set; }    // If empty, Surname is company name
        public string? Surname { get; set; }     // Individual last name OR company name
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public bool IsCompany => string.IsNullOrEmpty(Forename);
        public string ClientName => IsCompany ? Surname ?? "" : $"{Forename} {Surname}".Trim();
    }

    /// <summary>
    /// Contact info for company clients (not the main client)
    /// </summary>
    public class CaseContactInfo
    {
        public string? Title { get; set; }
        public string? Forename { get; set; }
        public string? Surname { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
        public string FullName => $"{Forename} {Surname}".Trim();
    }

    public class ProcessedClientData
    {
        public ExcelRowData? OriginalData { get; set; }
        public string? ClientType { get; set; }
        public int? TitleId { get; set; }
        public int FkBranchId { get; set; }
        public int FkUserId { get; set; }
        public DateTime DateTimeCreated { get; set; }
        public bool IsArchived { get; set; }
        public bool IsActive { get; set; }
    }

    public class MatterTypeMatch
    {
        public int AreaOfPracticeId { get; set; }
        public int? CaseTypeId { get; set; }
        public string? AreaOfPracticeName { get; set; }
        public string? CaseTypeName { get; set; }
        public double ConfidenceScore { get; set; }
        public bool IsExactMatch { get; set; }
    }

    public class S3Configuration
    {
        public string BucketName { get; set; } = string.Empty;
        public string ProfilePath { get; set; } = @"C:\AWS_Profile\LAITLegal.txt";
    }

    public class FileUploadInfo
    {
        public string LocalFilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string S3Key { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public string Extension { get; set; } = string.Empty;
    }

    public class FolderFileGroup
    {
        public string FolderName { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public int? FolderId { get; set; }
        public List<FileUploadInfo> Files { get; set; } = new List<FileUploadInfo>();
    }

    #region User Import Models

    public class UserExcelRowData
    {
        public string? Title { get; set; }
        public string? FirstName { get; set; }
        public string? MiddleName { get; set; }
        public string? LastName { get; set; }
        public string? UserCode { get; set; }        // F/E column (Fee Earner code)
        public string? FeeEarnerDescription { get; set; } // Full name from Excel
        public string? Email { get; set; }
        public string? HomePhone { get; set; }
        public string? Mobile { get; set; }
        public string? Address { get; set; }
        public string? Qualifications { get; set; }
        public string? Designation { get; set; }     // Fee Earner Status (Consultant, etc.)
        public string? NiNumber { get; set; }
        public DateTime? DateOfBirth { get; set; }
        public bool? Sex { get; set; }              // true = Male, false = Female
        public string? Notes { get; set; }
        public bool? InUse { get; set; }            // In Use column -> maps to IsActive

        public string FullName => string.Join(" ", new[] { FirstName, MiddleName, LastName }
            .Where(s => !string.IsNullOrWhiteSpace(s)));
    }

    public class ProcessedUserData
    {
        public UserExcelRowData? OriginalData { get; set; }
        public int? TitleId { get; set; }
        public int? FkUserRoleId { get; set; }
        public int? FkBranchId { get; set; }
        public bool IsActive { get; set; } = true;
        public bool IsDeleted { get; set; } = false;
        public bool IsDuplicate { get; set; } = false;
        public string? DuplicateReason { get; set; }
        public int? ExistingUserId { get; set; }    // If duplicate found, store existing user ID
    }

    public class ExistingUserInfo
    {
        public int UserId { get; set; }
        public string? UserCode { get; set; }
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? Email { get; set; }
        public bool? IsActive { get; set; }
        public bool? IsDeleted { get; set; }
    }

    #endregion
}
