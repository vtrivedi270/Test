# SPFx + React UI (SPFx 1.18+, Fluent UI v8)

## Setup
```
yo @microsoft/sharepoint  ->  WebPart, React, name: StationeryApp
npm i @fluentui/react recharts dayjs
```
Add to `config/package-solution.json`:
```json
"webApiPermissionRequests": [{ "resource": "StationeryApi", "scope": "user_impersonation" }]
```
(Approve in SharePoint Admin Center -> API access. The API app registration must expose `user_impersonation` and pre-authorize SharePoint Online Client Extensibility.)

## UX Layout
Single web part, left nav (Pivot / Nav), role-aware:

| Screen | Roles | Content |
|---|---|---|
| Dashboard | All (HRAD/Admin see org-wide) | KPI cards, monthly chart, by department, by item, store availability, filters, Export Excel/PDF |
| My Requests / Listing | All | DetailsList with status chips, filters, search, paging; tabs: Mine / Pending for me / All |
| New Request (entry form) | Requestor | Header (auto) + editable line grid + auto totals; Save Draft, Submit |
| Request Detail | All involved | Summary, line items, history timeline, action panel (Approve / Reject / Amend with comment; HRAD vendor + PDF; Requestor receive/close/query) |
| Admin | Admin | Tabs: Employees, Departments, Divisions, Items, Vendors with add/edit/deactivate panels |

Status colors: Draft grey, InReview blue, AmendRequired orange, Rejected red, HRApproval purple, Ordered teal, Delivered indigo, Closed green, PartialClosed amber.

## Folder structure
```
src/webparts/stationeryApp/
  StationeryAppWebPart.ts
  components/
    App.tsx                     (nav + routing by state)
    Dashboard/Dashboard.tsx
    Requests/RequestList.tsx
    Requests/RequestForm.tsx
    Requests/RequestDetail.tsx
    Requests/ApprovalPanel.tsx
    Requests/HradPanel.tsx
    Requests/HistoryTimeline.tsx
    Admin/MasterGrid.tsx
    common/StatusChip.tsx
  services/ApiService.ts
  models/index.ts
```

## services/ApiService.ts
```ts
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const API_RESOURCE = 'api://<api-client-id>';
const API_BASE = 'https://<your-api-host>/api';

export class ApiService {
  private client!: AadHttpClient;
  constructor(private ctx: WebPartContext) {}

  private async init() {
    if (!this.client)
      this.client = await this.ctx.aadHttpClientFactory.getClient(API_RESOURCE);
  }
  private async handle<T>(r: HttpClientResponse): Promise<T> {
    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      throw new Error(e.error || `Request failed (${r.status})`);
    }
    return r.status === 204 ? (undefined as any) : r.json();
  }
  async get<T>(path: string): Promise<T> {
    await this.init();
    return this.handle<T>(await this.client.get(`${API_BASE}/${path}`, AadHttpClient.configurations.v1));
  }
  async post<T>(path: string, body?: any): Promise<T> {
    await this.init();
    return this.handle<T>(await this.client.post(`${API_BASE}/${path}`, AadHttpClient.configurations.v1, {
      headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body ?? {})
    }));
  }
  async put<T>(path: string, body: any): Promise<T> {
    await this.init();
    return this.handle<T>(await this.client.fetch(`${API_BASE}/${path}`, AadHttpClient.configurations.v1, {
      method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
    }));
  }
  async downloadBlob(path: string, fileName: string) {
    await this.init();
    const r = await this.client.get(`${API_BASE}/${path}`, AadHttpClient.configurations.v1);
    const blob = await r.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = fileName; a.click();
    URL.revokeObjectURL(a.href);
  }
}
```

## models/index.ts
```ts
export type Status = 'Draft'|'InReview'|'AmendRequired'|'Rejected'|'HRApproval'|'Ordered'|'Delivered'|'Closed'|'PartialClosed';
export interface Me { employeeId: number; fullName: string; email: string; role: string;
  departmentName: string; divisionName: string; reportingManagerName?: string; }
export interface Item { itemId: number; itemCode: string; itemName: string; unit: string; rate: number; stockQty: number; }
export interface RequestLine { requestItemId?: number; itemId: number; itemName?: string; unit: string;
  rate: number; quantity: number; approvedQty?: number; receivedQty?: number; remarks?: string; }
export interface RequestHeader { requestId: number; requestNo: string; requestDate: string; requestorName: string;
  departmentName: string; remarks?: string; totalAmount: number; status: Status; currentStage?: string;
  currentApproverName?: string; items: RequestLine[]; history: HistoryRow[]; }
export interface HistoryRow { stage: string; action: string; actorName: string; comments?: string; actionDate: string; }
```

## common/StatusChip.tsx
```tsx
import * as React from 'react';
const colors: Record<string, string> = {
  Draft:'#6b7280', InReview:'#2563eb', AmendRequired:'#ea580c', Rejected:'#dc2626', HRApproval:'#7c3aed',
  Ordered:'#0d9488', Delivered:'#4f46e5', Closed:'#16a34a', PartialClosed:'#d97706' };
export const StatusChip = ({ status }: { status: string }) => (
  <span style={{ background: colors[status] || '#6b7280', color: '#fff', padding: '2px 10px',
    borderRadius: 12, fontSize: 12, whiteSpace: 'nowrap' }}>{status.replace(/([A-Z])/g, ' $1').trim()}</span>
);
```

## Requests/RequestForm.tsx (entry form)
```tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { TextField, Dropdown, IDropdownOption, PrimaryButton, DefaultButton, IconButton,
         MessageBar, MessageBarType, Stack, Label } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';
import { Item, Me, RequestLine } from '../../models';

interface Props { api: ApiService; me: Me; requestId?: number; onDone: () => void; }

export const RequestForm: React.FC<Props> = ({ api, me, requestId, onDone }) => {
  const [items, setItems] = useState<Item[]>([]);
  const [lines, setLines] = useState<RequestLine[]>([]);
  const [remarks, setRemarks] = useState('');
  const [id, setId] = useState<number | undefined>(requestId);
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.get<Item[]>('masters/items?active=true').then(setItems);
    if (requestId) api.get<any>(`requests/${requestId}`).then(r => { setRemarks(r.remarks || ''); setLines(r.items); });
  }, []);

  const options: IDropdownOption[] = items.map(i => ({ key: i.itemId, text: `${i.itemCode} - ${i.itemName}` }));
  const grand = lines.reduce((s, l) => s + (l.quantity || 0) * l.rate, 0);

  const addRow = () => setLines([...lines, { itemId: 0, unit: '', rate: 0, quantity: 1 }]);
  const setItem = (idx: number, itemId: number) => {
    if (lines.some((l, i) => i !== idx && l.itemId === itemId)) { setError('Item already added'); return; }
    const it = items.find(x => x.itemId === itemId)!;
    setLines(lines.map((l, i) => i === idx ? { ...l, itemId, itemName: it.itemName, unit: it.unit, rate: it.rate } : l));
  };
  const setQty = (idx: number, q: string) =>
    setLines(lines.map((l, i) => i === idx ? { ...l, quantity: Math.max(0, Number(q) || 0) } : l));

  const save = async (submit: boolean) => {
    setError(undefined);
    if (lines.length === 0 || lines.some(l => !l.itemId || l.quantity <= 0)) { setError('Add items with quantity > 0'); return; }
    setBusy(true);
    try {
      const saved = await api.post<{ requestId: number }>('requests', {
        requestId: id, remarks, items: lines.map(l => ({ itemId: l.itemId, quantity: l.quantity, remarks: l.remarks })) });
      setId(saved.requestId);
      if (submit) await api.post(`requests/${saved.requestId}/submit`);
      onDone();
    } catch (e: any) { setError(e.message); } finally { setBusy(false); }
  };

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { maxWidth: 1000 } }}>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>{error}</MessageBar>}
      <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
        <TextField label="Requestor" value={me.fullName} readOnly />
        <TextField label="Department" value={me.departmentName} readOnly />
        <TextField label="Division" value={me.divisionName} readOnly />
        <TextField label="Reporting Manager" value={me.reportingManagerName || '-'} readOnly />
        <TextField label="Date" value={new Date().toLocaleDateString()} readOnly />
      </Stack>
      <TextField label="Remarks" multiline rows={2} value={remarks} onChange={(_, v) => setRemarks(v || '')} />

      <Label>Items</Label>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead><tr>{['#', 'Item', 'Unit', 'Rate', 'Qty', 'Total', ''].map(h =>
          <th key={h} style={{ textAlign: 'left', padding: 6, borderBottom: '1px solid #ddd' }}>{h}</th>)}</tr></thead>
        <tbody>
          {lines.map((l, i) => (
            <tr key={i}>
              <td style={{ padding: 6 }}>{i + 1}</td>
              <td style={{ padding: 6, minWidth: 260 }}>
                <Dropdown selectedKey={l.itemId || undefined} options={options} placeholder="Select item"
                  onChange={(_, o) => setItem(i, Number(o!.key))} />
              </td>
              <td style={{ padding: 6 }}>{l.unit}</td>
              <td style={{ padding: 6 }}>{l.rate.toFixed(2)}</td>
              <td style={{ padding: 6, width: 90 }}>
                <TextField type="number" value={String(l.quantity)} onChange={(_, v) => setQty(i, v || '0')} />
              </td>
              <td style={{ padding: 6 }}>{(l.quantity * l.rate).toFixed(2)}</td>
              <td><IconButton iconProps={{ iconName: 'Delete' }} onClick={() => setLines(lines.filter((_, x) => x !== i))} /></td>
            </tr>))}
        </tbody>
        <tfoot><tr><td colSpan={5} style={{ textAlign: 'right', padding: 6, fontWeight: 600 }}>Grand Total</td>
          <td style={{ padding: 6, fontWeight: 600 }}>{grand.toFixed(2)}</td><td /></tr></tfoot>
      </table>
      <DefaultButton iconProps={{ iconName: 'Add' }} text="Add item" onClick={addRow} />

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <DefaultButton text="Save as draft" disabled={busy} onClick={() => save(false)} />
        <PrimaryButton text="Submit for approval" disabled={busy} onClick={() => save(true)} />
        <DefaultButton text="Cancel" onClick={onDone} />
      </Stack>
    </Stack>
  );
};
```

## Requests/RequestList.tsx (listing)
```tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { DetailsList, IColumn, SelectionMode, Pivot, PivotItem, SearchBox, Dropdown, DatePicker,
         Stack, DefaultButton, PrimaryButton } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';
import { StatusChip } from '../common/StatusChip';

export const RequestList: React.FC<{ api: ApiService; onOpen: (id: number) => void; onNew: () => void; canCreate: boolean }> =
({ api, onOpen, onNew, canCreate }) => {
  const [view, setView] = useState('mine');
  const [status, setStatus] = useState<string>();
  const [search, setSearch] = useState('');
  const [from, setFrom] = useState<Date>(); const [to, setTo] = useState<Date>();
  const [page, setPage] = useState(1);
  const [data, setData] = useState<{ total: number; rows: any[] }>({ total: 0, rows: [] });

  const load = () => {
    const q = new URLSearchParams({ view, page: String(page), pageSize: '20' });
    if (status) q.set('status', status); if (search) q.set('search', search);
    if (from) q.set('from', from.toISOString()); if (to) q.set('to', to.toISOString());
    api.get<any>(`requests?${q}`).then(setData);
  };
  useEffect(load, [view, status, page]);

  const cols: IColumn[] = [
    { key: 'no', name: 'Request No', fieldName: 'requestNo', minWidth: 110 },
    { key: 'dt', name: 'Date', minWidth: 90, onRender: r => new Date(r.requestDate).toLocaleDateString() },
    { key: 'rq', name: 'Requestor', fieldName: 'requestorName', minWidth: 140 },
    { key: 'dp', name: 'Department', fieldName: 'departmentName', minWidth: 120 },
    { key: 'am', name: 'Amount', minWidth: 80, onRender: r => r.totalAmount.toFixed(2) },
    { key: 'st', name: 'Status', minWidth: 110, onRender: r => <StatusChip status={r.status} /> },
    { key: 'ap', name: 'With', fieldName: 'currentApproverName', minWidth: 130 },
  ];
  const statusOpts = ['Draft','InReview','AmendRequired','Rejected','HRApproval','Ordered','Delivered','Closed','PartialClosed']
    .map(s => ({ key: s, text: s }));

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Stack horizontal horizontalAlign="space-between">
        <Pivot selectedKey={view} onLinkClick={i => { setView(i!.props.itemKey!); setPage(1); }}>
          <PivotItem headerText="My requests" itemKey="mine" />
          <PivotItem headerText="Pending for me" itemKey="pending" />
          <PivotItem headerText="All" itemKey="all" />
        </Pivot>
        {canCreate && <PrimaryButton iconProps={{ iconName: 'Add' }} text="New request" onClick={onNew} />}
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
        <SearchBox placeholder="Search no / requestor" onSearch={v => { setSearch(v); load(); }} onClear={() => setSearch('')} />
        <Dropdown placeholder="Status" options={[{ key: '', text: 'All' }, ...statusOpts]} styles={{ root: { width: 160 } }}
          onChange={(_, o) => setStatus(o!.key ? String(o!.key) : undefined)} />
        <DatePicker placeholder="From" value={from} onSelectDate={d => setFrom(d || undefined)} />
        <DatePicker placeholder="To" value={to} onSelectDate={d => setTo(d || undefined)} />
        <DefaultButton text="Apply" onClick={load} />
      </Stack>
      <DetailsList items={data.rows} columns={cols} selectionMode={SelectionMode.none}
        onItemInvoked={r => onOpen(r.requestId)} onRenderRow={(p, d) => <div onClick={() => onOpen(p!.item.requestId)} style={{ cursor: 'pointer' }}>{d!(p!)}</div>} />
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
        <DefaultButton text="Prev" disabled={page === 1} onClick={() => setPage(page - 1)} />
        <span>Page {page} of {Math.max(1, Math.ceil(data.total / 20))}</span>
        <DefaultButton text="Next" disabled={page * 20 >= data.total} onClick={() => setPage(page + 1)} />
      </Stack>
    </Stack>
  );
};
```

## Requests/ApprovalPanel.tsx
```tsx
import * as React from 'react';
import { useState } from 'react';
import { TextField, PrimaryButton, DefaultButton, Stack, MessageBar, MessageBarType } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';

export const ApprovalPanel: React.FC<{ api: ApiService; requestId: number; onDone: () => void }> = ({ api, requestId, onDone }) => {
  const [comments, setComments] = useState(''); const [err, setErr] = useState<string>();
  const act = async (action: 'Approve' | 'Reject' | 'Amend') => {
    if (action !== 'Approve' && !comments.trim()) { setErr('Comment is required'); return; }
    try { await api.post(`requests/${requestId}/action`, { action, comments }); onDone(); }
    catch (e: any) { setErr(e.message); }
  };
  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      <TextField label="Comments" multiline rows={3} value={comments} onChange={(_, v) => setComments(v || '')} />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Approve" onClick={() => act('Approve')} />
        <DefaultButton text="Ask to amend" onClick={() => act('Amend')} />
        <DefaultButton text="Reject" styles={{ root: { color: '#dc2626' } }} onClick={() => act('Reject')} />
      </Stack>
    </Stack>
  );
};
```

## Requests/HradPanel.tsx (shown when status = HRApproval)
```tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { Dropdown, TextField, PrimaryButton, DefaultButton, Stack } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';

export const HradPanel: React.FC<{ api: ApiService; req: any; onDone: () => void }> = ({ api, req, onDone }) => {
  const [vendors, setVendors] = useState<any[]>([]);
  const [vendorId, setVendorId] = useState<number>();
  const [lines, setLines] = useState<any[]>(req.items.map((i: any) => ({
    requestItemId: i.requestItemId, itemName: i.itemName, quantity: i.quantity,
    stock: i.stockQty, approvedQty: i.quantity, issuedFromStockQty: Math.min(i.quantity, i.stockQty || 0) })));
  useEffect(() => { api.get<any[]>('masters/vendors?active=true').then(setVendors); }, []);
  const upd = (idx: number, k: string, v: number) => setLines(lines.map((l, i) => i === idx ? { ...l, [k]: v } : l));

  const send = async () => {
    await api.post(`hrad/${req.requestId}/process`, { vendorId, lines: lines.map(l => ({
      requestItemId: l.requestItemId, approvedQty: l.approvedQty, issuedFromStockQty: l.issuedFromStockQty })) });
    onDone();
  };
  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <table><thead><tr><th>Item</th><th>Req</th><th>Stock</th><th>Approved</th><th>From store</th><th>To order</th></tr></thead>
        <tbody>{lines.map((l, i) => (
          <tr key={i}><td>{l.itemName}</td><td>{l.quantity}</td><td>{l.stock}</td>
            <td><TextField type="number" value={String(l.approvedQty)} onChange={(_, v) => upd(i, 'approvedQty', Number(v))} /></td>
            <td><TextField type="number" value={String(l.issuedFromStockQty)} onChange={(_, v) => upd(i, 'issuedFromStockQty', Number(v))} /></td>
            <td>{Math.max(0, l.approvedQty - l.issuedFromStockQty)}</td></tr>))}</tbody></table>
      <Dropdown label="Vendor" options={vendors.map(v => ({ key: v.vendorId, text: v.vendorName }))}
        onChange={(_, o) => setVendorId(Number(o!.key))} />
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Send to vendor (PDF + email)" disabled={!vendorId} onClick={send} />
        <DefaultButton text="Preview PDF" onClick={() => api.downloadBlob(`hrad/${req.requestId}/pdf`, `PO-${req.requestNo}.pdf`)} />
      </Stack>
    </Stack>
  );
};
```
Other post-order actions: **Delivered** button (HRAD, calls `hrad/{id}/deliver`), then **Receive panel** for the Requestor: a grid of lines with editable `receivedQty`, buttons *Confirm & Close* (all received), *Partial Close*, and *Raise Query* (text -> `requests/{id}/query`).

## Dashboard/Dashboard.tsx
```tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer } from 'recharts';
import { DatePicker, Dropdown, DefaultButton, Stack } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';
import dayjs from 'dayjs';

const Card = ({ t, v }: { t: string; v: string | number }) => (
  <div style={{ background: '#fff', border: '1px solid #e5e7eb', borderRadius: 8, padding: 16, minWidth: 170 }}>
    <div style={{ color: '#6b7280', fontSize: 12 }}>{t}</div><div style={{ fontSize: 24, fontWeight: 600 }}>{v}</div></div>);

export const Dashboard: React.FC<{ api: ApiService }> = ({ api }) => {
  const [from, setFrom] = useState(dayjs().startOf('year').toDate());
  const [to, setTo] = useState(new Date());
  const [deptId, setDeptId] = useState<number>();
  const [depts, setDepts] = useState<any[]>([]);
  const [sum, setSum] = useState<any>({}); const [monthly, setMonthly] = useState<any[]>([]);
  const [byDept, setByDept] = useState<any[]>([]); const [byItem, setByItem] = useState<any[]>([]); const [store, setStore] = useState<any[]>([]);

  const load = async () => {
    const q = `from=${from.toISOString()}&to=${to.toISOString()}${deptId ? `&deptId=${deptId}` : ''}`;
    const [s, m, d, i, st] = await Promise.all([
      api.get<any>(`dashboard/summary?${q}`), api.get<any[]>(`dashboard/monthly?${q}`),
      api.get<any[]>(`dashboard/by-department?${q}`), api.get<any[]>(`dashboard/by-item?${q}`), api.get<any[]>('dashboard/store')]);
    setSum(s); setMonthly(m); setByDept(d); setByItem(i); setStore(st);
  };
  useEffect(() => { api.get<any[]>('masters/departments?active=true').then(setDepts); load(); }, []);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
        <DatePicker label="From" value={from} onSelectDate={d => d && setFrom(d)} />
        <DatePicker label="To" value={to} onSelectDate={d => d && setTo(d)} />
        <Dropdown label="Department" styles={{ root: { width: 200 } }}
          options={[{ key: 0, text: 'All' }, ...depts.map(d => ({ key: d.departmentId, text: d.departmentName }))]}
          onChange={(_, o) => setDeptId(Number(o!.key) || undefined)} />
        <DefaultButton text="Apply" onClick={load} />
        <DefaultButton iconProps={{ iconName: 'ExcelDocument' }} text="Export Excel"
          onClick={() => api.downloadBlob(`reports/export?from=${from.toISOString()}&to=${to.toISOString()}${deptId ? `&deptId=${deptId}` : ''}`, 'stationery-report.xlsx')} />
      </Stack>
      <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
        <Card t="Total requests" v={sum.totalRequests ?? 0} />
        <Card t="Pending approvals" v={sum.pending ?? 0} />
        <Card t="Total cost" v={(sum.totalCost ?? 0).toLocaleString()} />
        <Card t="Low-stock items" v={sum.lowStock ?? 0} />
      </Stack>
      <h3>Monthly cost</h3>
      <ResponsiveContainer width="100%" height={240}><LineChart data={monthly}><CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="month" /><YAxis /><Tooltip /><Line dataKey="totalCost" stroke="#2563eb" /></LineChart></ResponsiveContainer>
      <h3>By department</h3>
      <ResponsiveContainer width="100%" height={240}><BarChart data={byDept}><CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="departmentName" /><YAxis /><Tooltip /><Bar dataKey="totalCost" fill="#7c3aed" /></BarChart></ResponsiveContainer>
      <h3>By item (top 10)</h3>
      <ResponsiveContainer width="100%" height={240}><BarChart data={byItem}><CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="itemName" /><YAxis /><Tooltip /><Bar dataKey="totalCost" fill="#0d9488" /></BarChart></ResponsiveContainer>
      <h3>Store availability</h3>
      <table style={{ width: '100%' }}><thead><tr><th align="left">Item</th><th>Stock</th><th>Reorder</th></tr></thead>
        <tbody>{store.map(s => <tr key={s.itemId} style={{ color: s.isLow ? '#dc2626' : undefined }}>
          <td>{s.itemName}</td><td align="center">{s.stockQty} {s.unit}</td><td align="center">{s.reorderLevel}</td></tr>)}</tbody></table>
    </Stack>
  );
};
```

## Admin/MasterGrid.tsx (reusable CRUD)
```tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { DetailsList, IColumn, SelectionMode, Panel, TextField, Toggle, PrimaryButton, DefaultButton, Stack } from '@fluentui/react';
import { ApiService } from '../../services/ApiService';

export interface Field { key: string; label: string; type?: 'text' | 'number'; required?: boolean; }

export const MasterGrid: React.FC<{ api: ApiService; entity: string; idKey: string; fields: Field[] }> = ({ api, entity, idKey, fields }) => {
  const [rows, setRows] = useState<any[]>([]);
  const [edit, setEdit] = useState<any>();
  const [err, setErr] = useState<string>();
  const load = () => api.get<any[]>(`masters/${entity}`).then(setRows);
  useEffect(() => { load(); }, [entity]);

  const cols: IColumn[] = [
    ...fields.map(f => ({ key: f.key, name: f.label, fieldName: f.key, minWidth: 100 } as IColumn)),
    { key: 'active', name: 'Active', minWidth: 60, onRender: r => (r.isActive ? 'Yes' : 'No') },
    { key: 'edit', name: '', minWidth: 60, onRender: r => <DefaultButton text="Edit" onClick={() => setEdit({ ...r })} /> }];

  const save = async () => {
    const miss = fields.find(f => f.required && !edit[f.key]);
    if (miss) { setErr(`${miss.label} is required`); return; }
    try { edit[idKey] ? await api.put(`masters/${entity}/${edit[idKey]}`, edit) : await api.post(`masters/${entity}`, edit);
      setEdit(undefined); setErr(undefined); load(); } catch (e: any) { setErr(e.message); }
  };
  return (<>
    <PrimaryButton text="Add new" onClick={() => setEdit({ isActive: true })} styles={{ root: { marginBottom: 8 } }} />
    <DetailsList items={rows} columns={cols} selectionMode={SelectionMode.none} />
    <Panel isOpen={!!edit} onDismiss={() => setEdit(undefined)} headerText={edit?.[idKey] ? 'Edit' : 'Add'} isFooterAtBottom
      onRenderFooterContent={() => <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Save" onClick={save} /><DefaultButton text="Cancel" onClick={() => setEdit(undefined)} /></Stack>}>
      {edit && <Stack tokens={{ childrenGap: 10 }}>
        {err && <span style={{ color: '#dc2626' }}>{err}</span>}
        {fields.map(f => <TextField key={f.key} label={f.label} type={f.type || 'text'} required={f.required}
          value={edit[f.key] ?? ''} onChange={(_, v) => setEdit({ ...edit, [f.key]: f.type === 'number' ? Number(v) : v })} />)}
        <Toggle label="Active" checked={edit.isActive} onChange={(_, c) => setEdit({ ...edit, isActive: !!c })} />
      </Stack>}
    </Panel>
  </>);
};
```
Usage in Admin tab: `<MasterGrid api={api} entity="items" idKey="itemId" fields={[{key:'itemCode',label:'Code',required:true},{key:'itemName',label:'Name',required:true},{key:'unit',label:'Unit',required:true},{key:'rate',label:'Rate',type:'number'},{key:'stockQty',label:'Stock',type:'number'},{key:'reorderLevel',label:'Reorder level',type:'number'}]} />`. Repeat for employees, departments, divisions, vendors. Dropdown fields (department, manager, role) can be added as a field `type: 'select'`.

## App.tsx (role-aware shell)
```tsx
const [me, setMe] = useState<Me>();
const [page, setPage] = useState<'dash'|'list'|'new'|'detail'|'admin'>('dash');
const [selId, setSelId] = useState<number>();
useEffect(() => { api.get<Me>('me').then(setMe); }, []);   // add GET api/me in API (CurrentUser + names)
// Nav items: Dashboard, Requests, New Request (if any role), Admin (only me.role === 'Admin')
// RequestDetail renders:
//   - ApprovalPanel   when req.status === 'InReview' && req.currentApproverEmail === me.email
//   - HradPanel       when req.status === 'HRApproval' && me.role === 'HRAD'
//   - Deliver button  when req.status === 'Ordered' && me.role === 'HRAD'
//   - Receive panel   when req.status === 'Delivered' && me is requestor
//   - Edit form       when status in (Draft, AmendRequired) && me is requestor
//   - HistoryTimeline always
```

## Build and deploy
```
gulp bundle --ship && gulp package-solution --ship
```
Upload `.sppkg` to the App Catalog, approve API permission, add the web part to a full-width page.
