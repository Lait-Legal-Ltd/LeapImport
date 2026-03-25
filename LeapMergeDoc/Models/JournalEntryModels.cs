namespace LeapMergeDoc.Models
{
    /// <summary>
    /// Raw data from Excel file for journal entry import
    /// Columns: Matter, Client, Matter Description, Last Trans. Date, Amount
    /// </summary>
    public class JournalExcelData
    {
        public string? Matter { get; set; }           // Key to lookup case
        public string? Client { get; set; }           // Client name (informational)
        public string? MatterDescription { get; set; } // Description (informational)
        public DateTime? LastTransDate { get; set; }   // Last transaction date (informational)
        public decimal Amount { get; set; }            // Balance amount
    }

    /// <summary>
    /// Processed journal import data with case lookup results
    /// </summary>
    public class JournalImportData
    {
        public int? CaseId { get; set; }
        public int? LedgerCardId { get; set; }
        public int AccountId { get; set; }
        public string? CaseReference { get; set; }
        public string? ClientName { get; set; }
        public decimal Balance { get; set; }
        public string? Description { get; set; }
        public string AccountType { get; set; } = "case";  // "case" or "bank"
        public int LineNumber { get; set; }
        public bool IsFound { get; set; }
    }

    /// <summary>
    /// Summary statistics for preview display
    /// </summary>
    public class JournalImportSummary
    {
        public int TotalRecords { get; set; }
        public int FoundCases { get; set; }
        public int NotFoundCases { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal FoundAmount { get; set; }
        public decimal NotFoundAmount { get; set; }
    }
}
