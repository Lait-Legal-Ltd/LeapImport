# Production Database Migration Guide

> ⚠️ **Run queries ONE BY ONE on production database**

---

## Step 1: Add case_credit Column (Run First)

```sql
ALTER TABLE tbl_case_details_general 
ADD COLUMN case_credit INT NULL DEFAULT NULL AFTER wh_status;
```

---

## Step 2: Clear Existing Data

### 2.1 Disable Foreign Key Checks
```sql
SET FOREIGN_KEY_CHECKS = 0;
```

### 2.2 Delete Case-Related Data (run one by one)
```sql
DELETE FROM tbl_case_client_greeting;
DELETE FROM tbl_case_clients;
DELETE FROM tbl_case_permissions;
DELETE FROM tbl_acc_ledger_cards;
DELETE FROM tbl_case_details_general;
```

### 2.3 Delete Client Data (run one by one)
```sql
DELETE FROM tbl_client_individual;
DELETE FROM tbl_client_company;
DELETE FROM tbl_client;
```

### 2.4 Delete Contact Data (if re-importing)
```sql
DELETE FROM tbl_contact_company;
DELETE FROM tbl_contact_personal;
DELETE FROM tbl_contact;
```

### 2.5 Reset Auto-Increment Counters
```sql
ALTER TABLE tbl_case_details_general AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_permissions AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_clients AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_ledger_cards AUTO_INCREMENT = 1;
ALTER TABLE tbl_client AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_individual AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_company AUTO_INCREMENT = 1;
ALTER TABLE tbl_contact AUTO_INCREMENT = 1;
ALTER TABLE tbl_contact_company AUTO_INCREMENT = 1;
ALTER TABLE tbl_contact_personal AUTO_INCREMENT = 1;
```

### 2.6 Re-enable Foreign Key Checks
```sql
SET FOREIGN_KEY_CHECKS = 1;
```

---

## Step 3: Export & Import Data

Use **mysqldump** or **HeidiSQL** to export from local and import to production.

**Import Order:**
1. Clients first
2. Cases second
3. Contacts last (will check against clients and skip duplicates)

---

## Step 4: Verify Import

```sql
SELECT 'tbl_client' AS table_name, COUNT(*) AS count FROM tbl_client
UNION ALL SELECT 'tbl_case_details_general', COUNT(*) FROM tbl_case_details_general
UNION ALL SELECT 'tbl_contact', COUNT(*) FROM tbl_contact
UNION ALL SELECT 'tbl_contact_company', COUNT(*) FROM tbl_contact_company
UNION ALL SELECT 'tbl_contact_personal', COUNT(*) FROM tbl_contact_personal;
```

---

## Checklist

- [ ] Step 1: Added case_credit column
- [ ] Step 2: Cleared existing data
- [ ] Step 3: Imported clients
- [ ] Step 3: Imported cases  
- [ ] Step 3: Imported contacts
- [ ] Step 4: Verified counts
- [ ] Step 5: Imported journal entries
- [ ] Step 6: Exported file upload data from local
- [ ] Step 6: Verified S3 files uploaded to production
- [ ] Step 6: Imported file upload data to production
- [ ] Step 6: Verified file upload import

---

## Step 5: Journal Entry Migration

### 5.1 Tables for Journal Entry (Import Order)

| # | Table | Description |
|---|-------|-------------|
| 1 | `tbl_acc_transactions` | Transaction records |
| 2 | `tbl_acc_journal_entry` | Journal entry headers |
| 3 | `tbl_acc_journal_entry_lines` | Journal entry line items |
| 4 | `tbl_acc_client_bank_transactions` | Bank transactions |
| 5 | `tbl_acc_ledger_card_transactions` | Ledger card transactions |

### 5.2 Clear Journal Data (if re-importing)

```sql
SET FOREIGN_KEY_CHECKS = 0;

DELETE FROM tbl_acc_ledger_card_transactions;
DELETE FROM tbl_acc_client_bank_transactions;
DELETE FROM tbl_acc_journal_entry_lines;
DELETE FROM tbl_acc_journal_entry;
DELETE FROM tbl_acc_transactions;

ALTER TABLE tbl_acc_transactions AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_journal_entry AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_journal_entry_lines AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_client_bank_transactions AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_ledger_card_transactions AUTO_INCREMENT = 1;

SET FOREIGN_KEY_CHECKS = 1;
```

### 5.3 Verify Journal Import

```sql
SELECT 'tbl_acc_transactions' AS table_name, COUNT(*) AS count FROM tbl_acc_transactions
UNION ALL SELECT 'tbl_acc_journal_entry', COUNT(*) FROM tbl_acc_journal_entry
UNION ALL SELECT 'tbl_acc_journal_entry_lines', COUNT(*) FROM tbl_acc_journal_entry_lines
UNION ALL SELECT 'tbl_acc_client_bank_transactions', COUNT(*) FROM tbl_acc_client_bank_transactions
UNION ALL SELECT 'tbl_acc_ledger_card_transactions', COUNT(*) FROM tbl_acc_ledger_card_transactions;
```

---

## Step 6: File Upload/Correspondence Migration

### 6.1 Tables for File Upload (Import Order)

| # | Table | Description |
|---|-------|-------------|
| 1 | `tbl_case_correspondence_folder` | Folder structure for case correspondence |
| 2 | `tbl_case_correspondence` | Correspondence records (one per folder) |
| 3 | `tbl_case_correspondence_document` | Document records (one per folder, title = folder name) |
| 4 | `tbl_case_correspondence_document_upload` | Upload records (one per folder) |
| 5 | `tbl_case_correspondence_document_upload_file` | Upload file records (one per file) |
| 6 | `tbl_case_correspondence_file` | File to folder mapping (one per file) |

### 6.2 Export Script from Local Database

**Option 1: Using mysqldump (Recommended)**

```bash
# Export all file upload tables for specific case IDs
mysqldump -u [username] -p [database_name] \
  tbl_case_correspondence_folder \
  tbl_case_correspondence \
  tbl_case_correspondence_document \
  tbl_case_correspondence_document_upload \
  tbl_case_correspondence_document_upload_file \
  tbl_case_correspondence_file \
  --where="fk_case_id IN (308,309,310)" \
  --no-create-info \
  --skip-triggers \
  > file_upload_export.sql
```

**Option 2: Using HeidiSQL**
1. Right-click on database → Export Database as SQL
2. Select only the 6 tables listed above
3. Filter by `fk_case_id` if needed
4. Export structure + data or data only

**Option 3: Export Specific Case IDs (SQL Script)**

```sql
-- Export folders
SELECT * FROM tbl_case_correspondence_folder 
WHERE fk_case_id IN (308,309,310) 
INTO OUTFILE '/tmp/tbl_case_correspondence_folder_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';

-- Export correspondence
SELECT * FROM tbl_case_correspondence 
WHERE fk_case_id IN (308,309,310) 
INTO OUTFILE '/tmp/tbl_case_correspondence_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';

-- Export documents
SELECT tccd.* FROM tbl_case_correspondence_document tccd
INNER JOIN tbl_case_correspondence tcc ON tccd.fk_correspondence_id = tcc.correspondence_id
WHERE tcc.fk_case_id IN (308,309,310)
INTO OUTFILE '/tmp/tbl_case_correspondence_document_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';

-- Export uploads
SELECT tccdu.* FROM tbl_case_correspondence_document_upload tccdu
INNER JOIN tbl_case_correspondence_document tccd ON tccdu.fk_case_document_id = tccd.case_document_id
INNER JOIN tbl_case_correspondence tcc ON tccd.fk_correspondence_id = tcc.correspondence_id
WHERE tcc.fk_case_id IN (308,309,310)
INTO OUTFILE '/tmp/tbl_case_correspondence_document_upload_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';

-- Export upload files
SELECT tccduf.* FROM tbl_case_correspondence_document_upload_file tccduf
INNER JOIN tbl_case_correspondence_document tccd ON tccduf.fk_case_document_id = tccd.case_document_id
INNER JOIN tbl_case_correspondence tcc ON tccd.fk_correspondence_id = tcc.correspondence_id
WHERE tcc.fk_case_id IN (308,309,310)
INTO OUTFILE '/tmp/tbl_case_correspondence_document_upload_file_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';

-- Export file mappings
SELECT * FROM tbl_case_correspondence_file 
WHERE fk_case_id IN (308,309,310)
INTO OUTFILE '/tmp/tbl_case_correspondence_file_export.sql'
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n';
```

### 6.3 Clear File Upload Data on Production (if re-importing)

**Option 1: Delete by Case ID (Selective)**

```sql
SET FOREIGN_KEY_CHECKS = 0;

-- Delete in reverse order (child to parent)
DELETE FROM tbl_case_correspondence_file WHERE fk_case_id IN (308,309,310);
DELETE FROM tbl_case_correspondence_document_upload_file 
WHERE fk_case_document_id IN (
    SELECT case_document_id FROM tbl_case_correspondence_document 
    WHERE fk_case_id IN (308,309,310)
);
DELETE FROM tbl_case_correspondence_document_upload 
WHERE fk_case_document_id IN (
    SELECT case_document_id FROM tbl_case_correspondence_document 
    WHERE fk_case_id IN (308,309,310)
);
DELETE FROM tbl_case_correspondence_document WHERE fk_case_id IN (308,309,310);
DELETE FROM tbl_case_correspondence WHERE fk_case_id IN (308,309,310);
DELETE FROM tbl_case_correspondence_folder WHERE fk_case_id IN (308,309,310);

SET FOREIGN_KEY_CHECKS = 1;
```

**Option 2: Truncate All Tables (Complete Reset)**

```sql
SET FOREIGN_KEY_CHECKS = 0;

-- Truncate in reverse order (child to parent)
TRUNCATE TABLE tbl_case_correspondence_file;
TRUNCATE TABLE tbl_case_correspondence_document_upload_file;
TRUNCATE TABLE tbl_case_correspondence_document_upload;
TRUNCATE TABLE tbl_case_correspondence_document;
TRUNCATE TABLE tbl_case_correspondence;
TRUNCATE TABLE tbl_case_correspondence_folder;

-- Reset auto-increment counters to 0
ALTER TABLE tbl_case_correspondence_folder AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_correspondence AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_correspondence_document AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_correspondence_document_upload AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_correspondence_document_upload_file AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_correspondence_file AUTO_INCREMENT = 1;

SET FOREIGN_KEY_CHECKS = 1;
```

### 6.4 Import to Production Database

**Using MySQL Command Line:**

```bash
mysql -u [username] -p [production_database] < file_upload_export.sql
```

**Using HeidiSQL:**
1. Open production database connection
2. File → Load SQL file
3. Select the exported SQL file
4. Execute

**Important Notes:**
- Ensure S3 bucket and files are already uploaded to production S3 bucket
- Verify `document_path` in `tbl_case_correspondence_document_upload_file` matches production S3 paths
- Check that `fk_user_id` values are correct (should be 1 for imported records)
- Verify folder structure matches between local and production

### 6.5 Verify File Upload Import

```sql
-- Verify counts by case
SELECT 
    fk_case_id,
    COUNT(DISTINCT folder_id) AS folder_count,
    COUNT(DISTINCT correspondence_id) AS correspondence_count,
    COUNT(DISTINCT case_document_id) AS document_count,
    COUNT(DISTINCT file_id) AS file_count
FROM (
    SELECT 
        tcc.fk_case_id,
        tccf.folder_id,
        tcc.correspondence_id,
        tccd.case_document_id,
        tccfile.file_id
    FROM tbl_case_correspondence tcc
    LEFT JOIN tbl_case_correspondence_folder tccf ON tccf.fk_case_id = tcc.fk_case_id
    LEFT JOIN tbl_case_correspondence_document tccd ON tccd.fk_correspondence_id = tcc.correspondence_id
    LEFT JOIN tbl_case_correspondence_file tccfile ON tccfile.fk_case_id = tcc.fk_case_id
    WHERE tcc.fk_case_id IN (308,309,310)
) AS combined
GROUP BY fk_case_id;

-- Verify folder structure
SELECT 
    fk_case_id,
    folder_name,
    parent_folder_id,
    folder_path,
    folder_order
FROM tbl_case_correspondence_folder
WHERE fk_case_id IN (308,309,310)
ORDER BY fk_case_id, parent_folder_id, folder_order;

-- Verify file mappings
SELECT 
    fk_case_id,
    fk_folder_id,
    COUNT(*) AS file_count
FROM tbl_case_correspondence_file
WHERE fk_case_id IN (308,309,310)
GROUP BY fk_case_id, fk_folder_id
ORDER BY fk_case_id, fk_folder_id;

-- Verify S3 paths
SELECT 
    fk_case_document_id,
    document_name,
    document_path
FROM tbl_case_correspondence_document_upload_file
WHERE fk_case_document_id IN (
    SELECT case_document_id FROM tbl_case_correspondence_document 
    WHERE fk_case_id IN (308,309,310)
)
LIMIT 10;
```

### 6.6 Checklist

- [ ] Step 6.1: Identified all tables for export
- [ ] Step 6.2: Exported data from local database
- [ ] Step 6.3: Cleared existing data on production (if re-importing)
- [ ] Step 6.4: Verified S3 files are uploaded to production bucket
- [ ] Step 6.4: Imported SQL script to production database
- [ ] Step 6.5: Verified folder counts match
- [ ] Step 6.5: Verified file mappings are correct
- [ ] Step 6.5: Verified S3 paths are accessible