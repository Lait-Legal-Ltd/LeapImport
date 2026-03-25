namespace LeapMergeDoc.Models
{
    /// <summary>
    /// Enum for different transaction types in the Excel sheet
    /// </summary>
    public enum AccountTransactionType
    {
        BankReceipt,
        BankPayment,
        ClientToOffice,
        Unknown
    }

    /// <summary>
    /// Raw data from Excel file for account update import
    /// </summary>
    public class AccountUpdateExcelData
    {
        public int RowNumber { get; set; }
        public string? TransactionType { get; set; }         // Type: Receipt, Payment, C2O
        public string? CaseReference { get; set; }           // Case reference to lookup
        public string? ClientName { get; set; }              // Client name (informational)
        public DateTime? TransactionDate { get; set; }       // Transaction date
        public decimal Amount { get; set; }                  // Transaction amount
        public string? Description { get; set; }             // Transaction description
        public string? PaymentReference { get; set; }        // Payment/Receipt reference
        public string? ReceivedFrom { get; set; }            // For receipts - who paid
        public string? PaidTo { get; set; }                  // For payments - who received
        public int PaymentTypeId { get; set; }               // 1=BACS, 2=Cash, 3=Cheque, etc.
        public string? Comments { get; set; }                // Additional comments
        public string? InvoiceNumber { get; set; }           // For C2O - invoice number
        public decimal? InvoiceAmount { get; set; }          // For C2O - invoice amount
        public string? BankAccountName { get; set; }         // Bank account name
        public int? ClientBankId { get; set; }               // Client bank ID
        public int? OfficeBankId { get; set; }               // Office bank ID
    }

    /// <summary>
    /// Processed account update data with case lookup results
    /// </summary>
    public class AccountUpdateImportData
    {
        public int RowNumber { get; set; }
        public AccountTransactionType TransactionType { get; set; }
        public int? CaseId { get; set; }
        public int? ClientId { get; set; }
        public int? LedgerCardId { get; set; }
        public string? CaseReference { get; set; }
        public string? ClientName { get; set; }
        public DateTime TransactionDate { get; set; }
        public decimal Amount { get; set; }
        public string? Description { get; set; }
        public string? PaymentReference { get; set; }
        public string? ReceivedFrom { get; set; }
        public string? PaidTo { get; set; }
        public int PaymentTypeId { get; set; }
        public string? Comments { get; set; }
        public string? InvoiceNumber { get; set; }
        public decimal InvoiceAmount { get; set; }
        public int ClientBankId { get; set; }
        public int OfficeBankId { get; set; }
        public bool IsFound { get; set; }
        public bool IsValid { get; set; }
        public string? ValidationError { get; set; }
        
        // For invoice creation (C2O)
        public int? InvoiceId { get; set; }
        public int? IncomeAccountId { get; set; }
    }

    /// <summary>
    /// Summary statistics for preview display
    /// </summary>
    public class AccountUpdateImportSummary
    {
        public int TotalRecords { get; set; }
        public int ValidRecords { get; set; }
        public int InvalidRecords { get; set; }
        
        public int ReceiptCount { get; set; }
        public decimal ReceiptTotal { get; set; }
        
        public int PaymentCount { get; set; }
        public decimal PaymentTotal { get; set; }
        
        public int ClientToOfficeCount { get; set; }
        public decimal ClientToOfficeTotal { get; set; }
        
        public int FoundCases { get; set; }
        public int NotFoundCases { get; set; }
        
        public List<string> Errors { get; set; } = new List<string>();
    }

    /// <summary>
    /// Result of import operation
    /// </summary>
    public class AccountUpdateImportResult
    {
        public int SuccessCount { get; set; }
        public int ErrorCount { get; set; }
        public int ReceiptsCreated { get; set; }
        public int PaymentsCreated { get; set; }
        public int InvoicesCreated { get; set; }
        public int ClientToOfficeCreated { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }

    /// <summary>
    /// Bank account information
    /// </summary>
    public class BankAccountInfo
    {
        public int BankId { get; set; }
        public string? BankName { get; set; }
        public string? AccountNumber { get; set; }
        public string? SortCode { get; set; }
        public string? Institution { get; set; }
        public bool IsClientBank { get; set; }
        public decimal OpeningBalance { get; set; }
        
        // Display name for dropdown
        public string DisplayName => $"{BankName} ({AccountNumber})";
    }

    /// <summary>
    /// Payment type information
    /// </summary>
    public class PaymentTypeInfo
    {
        public int PaymentTypeId { get; set; }
        public string? PaymentTypeName { get; set; }
    }
}
