namespace LeapMergeDoc.Models
{
    /// <summary>
    /// Raw data from CSV file for journal entry import
    /// Columns: Client (code), Matter, Client Name, Matter Description, F/E, W/T, Client (balance)
    /// </summary>
    public class JournalExcelData
    {
        public string? ClientCode { get; set; }        // Client code (e.g., "2DE0001")
        public string? Matter { get; set; }            // Matter number (e.g., "1")
        public string? ClientName { get; set; }        // Client name (e.g., "2 Degrees Limited")
        public string? MatterDescription { get; set; } // Description (e.g., "EMPLOYMENT")
        public string? FeeEarner { get; set; }         // F/E code
        public string? WorkType { get; set; }          // W/T code
        public decimal Amount { get; set; }            // Balance amount (last column "Client")
        
        // Combined reference for case lookup (format: ClientCode-Matter)
        public string CaseReference => $"{ClientCode}-{Matter}";
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
        public bool HasLedgerCard => LedgerCardId.HasValue;
    }

    /// <summary>
    /// Summary statistics for preview display
    /// </summary>
    public class JournalImportSummary
    {
        public int TotalRecords { get; set; }
        public int FoundCases { get; set; }
        public int NotFoundCases { get; set; }
        public int MissingLedgerCards { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal FoundAmount { get; set; }
        public decimal NotFoundAmount { get; set; }
    }
}
