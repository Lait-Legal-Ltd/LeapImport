# Clear Client and Case Data - Production SQL Commands

> ⚠️ **CRITICAL: Run queries ONE BY ONE on production database**
> 
> This will DELETE ALL client and case data. Cannot be undone!

---

## Step 1: Disable Foreign Key Checks

```sql
SET FOREIGN_KEY_CHECKS = 0;
```

---

## Step 2: Delete Case-Related Data (Run in this order)

```sql
-- 2.1 Delete case client greeting
DELETE FROM tbl_case_client_greeting;

-- 2.2 Delete case clients relationship
DELETE FROM tbl_case_clients;

-- 2.3 Delete case permissions
DELETE FROM tbl_case_permissions;

-- 2.4 Delete ledger cards (linked to cases)
DELETE FROM tbl_acc_ledger_cards;

-- 2.5 Delete case details (main case table)
DELETE FROM tbl_case_details_general;
```

---

## Step 3: Delete Client-Related Data (Run in this order)

```sql
-- 3.1 Delete individual client details
DELETE FROM tbl_client_individual;

-- 3.2 Delete company client details
DELETE FROM tbl_client_company;

-- 3.3 Delete main client records
DELETE FROM tbl_client;
```

---

## Step 4: Reset Auto-Increment Counters

```sql
-- Reset case tables
ALTER TABLE tbl_case_details_general AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_permissions AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_clients AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_client_greeting AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_ledger_cards AUTO_INCREMENT = 1;

-- Reset client tables
ALTER TABLE tbl_client AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_individual AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_company AUTO_INCREMENT = 1;
```

---

## Step 5: Re-enable Foreign Key Checks

```sql
SET FOREIGN_KEY_CHECKS = 1;
```

---

## Quick Copy - ALL IN ONE (Use with caution!)

```sql
-- ═══════════════════════════════════════════════════════════════════
-- CLEAR ALL CLIENT AND CASE DATA
-- ═══════════════════════════════════════════════════════════════════

SET FOREIGN_KEY_CHECKS = 0;

-- Case tables (child to parent order)
DELETE FROM tbl_case_client_greeting;
DELETE FROM tbl_case_clients;
DELETE FROM tbl_case_permissions;
DELETE FROM tbl_acc_ledger_cards;
DELETE FROM tbl_case_details_general;

-- Client tables (child to parent order)
DELETE FROM tbl_client_individual;
DELETE FROM tbl_client_company;
DELETE FROM tbl_client;

-- Reset auto-increment
ALTER TABLE tbl_case_details_general AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_permissions AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_clients AUTO_INCREMENT = 1;
ALTER TABLE tbl_case_client_greeting AUTO_INCREMENT = 1;
ALTER TABLE tbl_acc_ledger_cards AUTO_INCREMENT = 1;
ALTER TABLE tbl_client AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_individual AUTO_INCREMENT = 1;
ALTER TABLE tbl_client_company AUTO_INCREMENT = 1;

SET FOREIGN_KEY_CHECKS = 1;

-- ═══════════════════════════════════════════════════════════════════
-- COMPLETE
-- ═══════════════════════════════════════════════════════════════════
```

---

## Verification Queries

```sql
-- Check all tables are empty
SELECT 'tbl_client' AS table_name, COUNT(*) AS count FROM tbl_client
UNION ALL SELECT 'tbl_client_individual', COUNT(*) FROM tbl_client_individual
UNION ALL SELECT 'tbl_client_company', COUNT(*) FROM tbl_client_company
UNION ALL SELECT 'tbl_case_details_general', COUNT(*) FROM tbl_case_details_general
UNION ALL SELECT 'tbl_case_permissions', COUNT(*) FROM tbl_case_permissions
UNION ALL SELECT 'tbl_case_clients', COUNT(*) FROM tbl_case_clients
UNION ALL SELECT 'tbl_case_client_greeting', COUNT(*) FROM tbl_case_client_greeting
UNION ALL SELECT 'tbl_acc_ledger_cards', COUNT(*) FROM tbl_acc_ledger_cards;
```

**Expected Result:** All counts should be `0`

---

## Checklist

- [ ] Step 1: Disabled foreign key checks
- [ ] Step 2: Deleted all case data (5 tables)
- [ ] Step 3: Deleted all client data (3 tables)
- [ ] Step 4: Reset auto-increment counters
- [ ] Step 5: Re-enabled foreign key checks
- [ ] Verified all counts are 0
