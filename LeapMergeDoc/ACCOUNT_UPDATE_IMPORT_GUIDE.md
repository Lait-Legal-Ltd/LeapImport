# Account Update Import - Excel Template Guide

This document describes the expected Excel format for importing accounting transactions into LaitLegal.

## Overview

The Account Update Import feature allows you to import three types of transactions:
1. **Bank Receipts** - Money received into client bank accounts
2. **Bank Payments** - Money paid out from client bank accounts  
3. **Client to Office (C2O)** - Transfers from client bank to office bank (automatically creates an invoice)

## Excel Format: Bank Account Register

The system reads the "Bank Account Register" format with the following structure:

### Row Structure
- **Row 1**: Title row (ignored)
- **Row 2**: Header row
- **Row 3+**: Data rows

### Column Structure

| Column | Header Name | Description |
|--------|-------------|-------------|
| B | Transaction Date | Date of transaction (DD/MM/YYYY) |
| D | Transaction Details | Matter details (informational) |
| E | Reason | Transaction type: `Receipt` or `Payment` |
| H | Reason: Opening Balance | Transaction description/reason |
| J | Withdrawal | Payment amount (money out) |
| K | Deposit Balance | Receipt amount (money in) |
| M | BR | Bank reconciliation flag |
| O | Receipts /Payment | Type indicator |
| P | Client to office | C2O indicator (if present, creates invoice + transfer) |
| Q | Case Number | **KEY FIELD**: Maps to `case_reference_auto` in database |

### Case Number Field (Column Q)

This is the most important field. It must match the `case_reference_auto` field in `tbl_case_details_general`.

Examples from your data:
- `AS.CLT.LLT.20.25.Rav`
- `BT.IM.APL.9.Alkawari`
- `BT.CJ.GEN.20.19.Alka`
- `BT.IM.GEN.21.58.Alwa`

### Transaction Type Detection

The system determines transaction type based on:

1. **Receipt**: 
   - Column E (Reason) = "Receipt" OR
   - Column K (Deposit Balance) has a value > 0

2. **Payment**:
   - Column E (Reason) = "Payment" OR
   - Column J (Withdrawal) has a value > 0

3. **Client to Office (C2O)**:
   - Column P (Client to office) contains text

## Example Data (from your file)

```
Transaction Date | Transaction Details           | Reason  | Withdrawal | Deposit Balance | Receipts/Payment | Client to office | Case Number
19/04/2021       | Matter No. AS.CLT.LLT.20.25  | Receipt |            | 1,000.00        | Receipt          |                  | AS.CLT.LLT.20.25.Rav
19/04/2021       | Matter No. BT.IM.APL.9       | Payment | 500.00     |                 | Payment          |                  | BT.IM.APL.9.Alkawari
26/04/2021       | Matter No. BT.CJ.GEN.20.19   | Receipt |            | 500.00          | Receipt          |                  | BT.CJ.GEN.20.19.Alka
```

## Database Lookup Process

1. **Primary**: Exact match on `case_reference_auto`
2. **Fallback 1**: Exact match on `case_reference_manual`
3. **Fallback 2**: Partial match (LIKE) on either field

Once a case is found, the system retrieves:
- `case_id` - The unique case identifier
- `fk_client_id` - The client ID
- `ledger_card_id` - From `tbl_acc_ledger_cards` (for ledger transactions)

## Processing Order

The import processes transactions in the following order:
1. **Bank Receipts** - Processed first
2. **Bank Payments** - Processed second
3. **Client to Office** - Processed last (invoice is created automatically before the C2O transfer)

## Validation Rules

- CaseReference must match an existing case in the system
- Amount must be greater than 0
- TransactionDate must be a valid date
- For C2O transactions, an Office Bank must be selected
- Client Bank ID is always required (uses default if not specified)

## Database Tables Affected

### Bank Receipt
- `tbl_acc_transactions` - Main transaction record
- `tbl_acc_client_receipt` - Receipt record
- `tbl_acc_client_bank_transactions` - DR Client Bank (money in)
- `tbl_acc_ledger_card_transactions` - CR Client Ledger

### Bank Payment
- `tbl_acc_transactions` - Main transaction record
- `tbl_acc_client_payment` - Payment record
- `tbl_acc_client_bank_transactions` - CR Client Bank (money out)
- `tbl_acc_ledger_card_transactions` - DR Client Ledger

### Client to Office
**Invoice**:
- `tbl_acc_transactions` - Invoice transaction
- `tbl_acc_invoice` - Simple invoice (Professional Fee, NO VAT)
  - tax_amount = 0
  - total_amount = net_amount
  - balance_due = 0 (immediately paid)
  - status = 9 (fully paid)
- `tbl_acc_invoice_status` - Status set to Paid

**C2O Transfer**:
- `tbl_acc_transactions` - Transfer transaction
- `tbl_acc_client_to_office_transactions` - C2O record
- `tbl_acc_client_bank_transactions` - CR Client Bank (money out from client bank)
- `tbl_acc_ledger_card_transactions` - Client DR + Office CR (moves balance from client to office ledger)

**NOTE**: No office bank transaction is created. Money is not physically transferred to office bank account.

## Payment Type

All transactions use **Bank Transfer** (payment_type_id = 6) as the default payment type since original payment methods are unknown.

## Tips

1. Always preview data before importing to check for validation errors
2. Ensure all case references exist in the system before importing
3. Set default bank accounts before loading the Excel file
4. For C2O transactions, if no invoice number is provided, one will be auto-generated
5. Transactions are processed within database transactions - if one fails, it will be rolled back individually

## Transaction Logic Summary

| Transaction | Receipt Table | Payment Table | Invoice | Client Bank | Office Bank | Ledger (Client) | Ledger (Office) |
|-------------|--------------|---------------|---------|-------------|-------------|-----------------|-----------------|
| Bank Receipt | ✅ INSERT | - | - | ✅ DR | - | ✅ CR | - |
| Bank Payment | - | ✅ INSERT | - | ✅ CR | - | ✅ DR | - |
| Client to Office | - | - | ✅ INSERT (No VAT, Paid) | ✅ CR | ❌ | ✅ DR | ✅ CR |

## Troubleshooting

| Error | Solution |
|-------|----------|
| Case not found | Verify the case reference exists in the system |
| Invalid amount | Ensure amount is a positive number |
| Bank ID required | Select a default bank account or specify in Excel |
| Unknown transaction type | Use valid type: Receipt, Payment, or C2O |
