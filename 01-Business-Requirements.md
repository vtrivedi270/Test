# Stationery Management System: Business Requirements

## 1. Purpose
Replace manual / email-based stationery requests with a controlled, auditable workflow: employees request items, managers approve, HRAD procures from vendors, and requestors confirm receipt. Management gets cost and consumption visibility.

## 2. Scope
In scope: request entry, multi-level approval, vendor ordering (PDF), delivery verification, store stock view, reports, master data administration, audit/exception/email logging.
Out of scope (phase 2): vendor invoicing/payments, barcode scanning.

## 3. Technology
- UI: SharePoint Framework (SPFx) + React + Fluent UI
- API: .NET Core Web API (Azure AD / Entra ID bearer auth)
- Database: SQL Server
- Email: Microsoft Graph (sendMail) or SMTP
- PDF: QuestPDF (server side)

## 4. Roles
| Role | Responsibility |
|---|---|
| Requestor | Create/edit draft, submit, track, verify delivery, close or partially close, raise query |
| Reporting Manager | Approve / Reject / Ask to Amend |
| Department Head | Approve / Reject / Ask to Amend |
| Division Head | Approve / Reject / Ask to Amend |
| HRAD | Validate request, adjust approved qty, select vendor, generate PDF/PO, send to vendor, record delivery, hand over to requestor |
| Admin | CRUD on all masters, view all data and logs |

Role comes from the Employee master (reporting manager, dept head and div head derived from the requestor's record).

## 4.1 Functional Requirements

### FR-1 Entry form
- Header: auto-filled requestor details (name, code, department, division, reporting manager), request date (auto), remarks.
- Line items: select item from Item master (unit and rate auto-filled), enter quantity, line total = qty x rate, grand total auto-calculated. Add/remove rows. No duplicate item in one request.
- Save as Draft or Submit. Mandatory: at least one line, qty > 0.

### FR-2 Approval workflow
1. Requestor submits (status: In Review, stage Reporting Manager)
2. Reporting Manager -> Dept Head -> Division Head, in sequence
3. At each stage: **Approve** (moves on), **Reject** (comment mandatory, request ends as Rejected), **Amend** (comment mandatory, returns to Requestor as Draft-Amend; resubmission restarts from Reporting Manager)
4. After Div Head approval: status **HR Approval**
5. Skip rules: if requestor is a manager/head, levels that equal the requestor or are empty are skipped automatically.

### FR-3 HRAD processing
- Review all details, may reduce approved quantity per line (with remarks).
- Check store availability: if stock covers the line, issue from store; otherwise order from vendor.
- Select vendor from Vendor master; generate Vendor PDF (PO) and email it; record PO number/date.
- Record delivery per line (received qty). Cross-check and hand over to requestor.

### FR-4 Requestor closure
- Verify received items per line. All received -> **Close**. Some missing -> **Partial Close** (shortfall remains visible to HRAD for follow-up) . Any issue -> raise query (email to HRAD logged against request).

### FR-5 Statuses
Draft, In Review, HR Approval, Ordered, Delivered (pending requestor), Closed, Partial Closed, Rejected, Amend Required.
(Ordered, Delivered, Rejected, Amend Required are added to make the flow trackable.)

### FR-6 Listing and history
- Listing with filter (status, date range, department, requestor, item), search, paging, sorting. Requestors see own; approvers see pending-for-me and team; HRAD/Admin see all.
- History tab per request: every action with actor, date/time, comment (approval timeline).

### FR-7 Dashboard
KPI cards: total requests, pending approvals, total cost (month), items below reorder level.
Charts: monthly cost/quantity trend; by department; by item (top N); store availability list.
Filters: month range, department, division, item, status. Export to Excel/CSV and PDF.

### FR-8 Admin
UI CRUD (add, edit, activate/deactivate; no hard delete when referenced) for Employee, Department, Division, Item (with category, unit, rate, stock qty, reorder level), Vendor.

### FR-9 Notifications and logs
- Email on: submit, pending approval, approve/reject/amend, HR sent to vendor, delivery ready, closure, query.
- Every email stored in EmailLog (status, error, retry count).
- AuditLog for create/update/status change; ExceptionLog for API errors.

## 5. Non-functional
- Authentication via Entra ID; authorization by role on API (never only in UI).
- Row-level rules: requestor sees own data only.
- Soft delete on masters; all tables have CreatedBy/At, ModifiedBy/At.
- Concurrency: RowVersion on RequestHeader.
- Performance: list pages < 2 s for 100k requests (server paging, indexes).
- Responsive for SharePoint pages and Teams tab.

## 6. Business Rules
1. Request number format: `STN-YYYY-00001`.
2. Rate is frozen into the line at submit time.
3. Approvers cannot approve own request.
4. Only Draft / Amend Required requests are editable by requestor.
5. HRAD issue from store reduces stock; vendor receipt increases then issue reduces stock (StockTransaction).
6. Delegation / absent approver: Admin can reassign pending approval (phase 2 if not needed now).

## 7. Acceptance criteria (samples)
- Totals recalculate on qty change.
- Reject without comment is blocked.
- Amended request restarts approval and shows full history.
- Partial close leaves outstanding qty visible in HRAD view.
- Vendor PDF contains header, vendor, line items, qty, rate, total.
