# .NET Core API (net8.0) - StationeryApi

## Structure
```
StationeryApi/
  Program.cs
  Data/AppDbContext.cs            (EF Core, DB-first scaffold from StationeryDB)
  Models/Enums.cs
  Dtos/Dtos.cs
  Services/WorkflowService.cs     (state machine)
  Services/RequestService.cs
  Services/EmailService.cs        (Graph sendMail + EmailLog)
  Services/PdfService.cs          (QuestPDF vendor PO)
  Middleware/ExceptionMiddleware.cs
  Controllers/ RequestsController, ApprovalsController, HradController,
               MastersController, DashboardController, ReportsController
```
Packages: `Microsoft.EntityFrameworkCore.SqlServer`, `Microsoft.Identity.Web`, `QuestPDF`, `ClosedXML`, `Microsoft.Graph`.
Scaffold: `dotnet ef dbcontext scaffold "<conn>" Microsoft.EntityFrameworkCore.SqlServer -o Models --context-dir Data -c AppDbContext`

## Program.cs
```csharp
using Microsoft.EntityFrameworkCore;
using Microsoft.Identity.Web;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<AppDbContext>(o =>
    o.UseSqlServer(builder.Configuration.GetConnectionString("Default")));
builder.Services.AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd");
builder.Services.AddAuthorization();
builder.Services.AddScoped<WorkflowService>();
builder.Services.AddScoped<RequestService>();
builder.Services.AddScoped<EmailService>();
builder.Services.AddScoped<PdfService>();
builder.Services.AddScoped<CurrentUser>();
builder.Services.AddHttpContextAccessor();
builder.Services.AddControllers().AddJsonOptions(o =>
    o.JsonSerializerOptions.ReferenceHandler = System.Text.Json.Serialization.ReferenceHandler.IgnoreCycles);
builder.Services.AddCors(o => o.AddPolicy("spfx", p => p
    .WithOrigins("https://<tenant>.sharepoint.com")
    .AllowAnyHeader().AllowAnyMethod()));
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();
app.UseMiddleware<ExceptionMiddleware>();
app.UseSwagger(); app.UseSwaggerUI();
app.UseHttpsRedirection();
app.UseCors("spfx");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.Run();
```

appsettings.json: `AzureAd` (Instance, TenantId, ClientId, Audience = `api://<client-id>`), `ConnectionStrings:Default`.

## CurrentUser (resolves employee + role from token)
```csharp
public class CurrentUser
{
    private readonly AppDbContext _db; private readonly IHttpContextAccessor _ctx;
    private EmployeeMaster? _emp;
    public CurrentUser(AppDbContext db, IHttpContextAccessor ctx) { _db = db; _ctx = ctx; }
    public string Email => _ctx.HttpContext!.User.FindFirst("preferred_username")?.Value
                        ?? _ctx.HttpContext!.User.FindFirst("upn")?.Value ?? "";
    public async Task<EmployeeMaster> GetAsync() =>
        _emp ??= await _db.EmployeeMasters.FirstOrDefaultAsync(e => e.Email == Email && e.IsActive)
                 ?? throw new UnauthorizedAccessException("User not found in Employee master");
}
```

## DTOs
```csharp
public record RequestItemDto(int ItemId, decimal Quantity, string? Remarks);
public record RequestSaveDto(int? RequestId, string? Remarks, List<RequestItemDto> Items);
public record ActionDto(string Action, string? Comments);          // Approve | Reject | Amend
public record HradProcessDto(int VendorId, List<HradLineDto> Lines, string? Comments);
public record HradLineDto(int RequestItemId, decimal ApprovedQty, decimal IssuedFromStockQty);
public record ReceiveDto(List<ReceiveLineDto> Lines, string? Comments);
public record ReceiveLineDto(int RequestItemId, decimal ReceivedQty);
public record ListFilter(string? Status, DateTime? From, DateTime? To, int? DeptId,
                         string? Search, string? View, int Page = 1, int PageSize = 20);
```

## WorkflowService (state machine)
```csharp
public class WorkflowService
{
    private readonly AppDbContext _db; private readonly EmailService _mail;
    public WorkflowService(AppDbContext db, EmailService mail) { _db = db; _mail = mail; }

    // Next approver after a stage, skipping empty / same-as-requestor levels
    public async Task<(string stage, int? approverId)> NextStageAsync(RequestHeader h, string? after)
    {
        var req = await _db.EmployeeMasters.FindAsync(h.RequestorId);
        var dept = await _db.DepartmentMasters.FindAsync(h.DepartmentId);
        var div  = await _db.DivisionMasters.FindAsync(h.DivisionId);
        var chain = new List<(string Stage, int? Id)> {
            ("ReportingManager", req!.ReportingManagerId),
            ("DeptHead",         dept!.DepartmentHeadId),
            ("DivHead",          div!.DivisionHeadId) };
        int start = after == null ? 0 : chain.FindIndex(c => c.Stage == after) + 1;
        for (int i = start; i < chain.Count; i++)
        {
            var (stage, id) = chain[i];
            if (id != null && id != h.RequestorId) return (stage, id);
        }
        return ("HRAD", null);                                  // all approvals done
    }

    public async Task SubmitAsync(RequestHeader h, EmployeeMaster actor)
    {
        if (h.Status is not ("Draft" or "AmendRequired")) throw new InvalidOperationException("Not editable");
        if (!h.RequestItems.Any()) throw new InvalidOperationException("Add at least one item");
        var (stage, id) = await NextStageAsync(h, null);
        h.SubmittedAt = DateTime.UtcNow;
        Move(h, stage, id);
        AddHistory(h, "Requestor", h.Status == "AmendRequired" ? "Resubmit" : "Submit", actor, null);
        await _db.SaveChangesAsync();
        await _mail.NotifyAsync(h, "Pending your approval");
    }

    public async Task ActAsync(RequestHeader h, EmployeeMaster actor, ActionDto dto)
    {
        if (h.Status != "InReview" || h.CurrentApproverId != actor.EmployeeId)
            throw new UnauthorizedAccessException("Not your approval");
        if (dto.Action != "Approve" && string.IsNullOrWhiteSpace(dto.Comments))
            throw new InvalidOperationException("Comment is mandatory");

        var stage = h.CurrentStage!;
        switch (dto.Action)
        {
            case "Approve":
                var (next, id) = await NextStageAsync(h, stage);
                Move(h, next, id);
                break;
            case "Reject": h.Status = "Rejected"; h.CurrentStage = null; h.CurrentApproverId = null; break;
            case "Amend":  h.Status = "AmendRequired"; h.CurrentStage = "Requestor"; h.CurrentApproverId = h.RequestorId; break;
            default: throw new InvalidOperationException("Invalid action");
        }
        AddHistory(h, stage, dto.Action, actor, dto.Comments);
        await _db.SaveChangesAsync();
        await _mail.NotifyAsync(h, dto.Action);
    }

    private static void Move(RequestHeader h, string stage, int? approver)
    {
        h.CurrentStage = stage; h.CurrentApproverId = approver;
        h.Status = stage == "HRAD" ? "HRApproval" : "InReview";
    }
    public void AddHistory(RequestHeader h, string stage, string action, EmployeeMaster a, string? c) =>
        _db.ApprovalHistories.Add(new ApprovalHistory {
            RequestId = h.RequestId, Stage = stage, Action = action, ActorId = a.EmployeeId, Comments = c });
}
```

## RequestService (save + totals + number)
```csharp
public async Task<RequestHeader> SaveAsync(RequestSaveDto dto, EmployeeMaster me)
{
    RequestHeader h;
    if (dto.RequestId is null)
    {
        var no = new SqlParameter("@RequestNo", System.Data.SqlDbType.NVarChar, 30) { Direction = System.Data.ParameterDirection.Output };
        await _db.Database.ExecuteSqlRawAsync("EXEC dbo.usp_NextRequestNo @RequestNo OUTPUT", no);
        h = new RequestHeader { RequestNo = (string)no.Value, RequestorId = me.EmployeeId,
            DepartmentId = me.DepartmentId, DivisionId = me.DivisionId, Status = "Draft",
            CreatedBy = me.Email };
        _db.RequestHeaders.Add(h);
    }
    else
    {
        h = await _db.RequestHeaders.Include(x => x.RequestItems).FirstAsync(x => x.RequestId == dto.RequestId);
        if (h.RequestorId != me.EmployeeId) throw new UnauthorizedAccessException();
        if (h.Status is not ("Draft" or "AmendRequired")) throw new InvalidOperationException("Not editable");
        _db.RequestItems.RemoveRange(h.RequestItems);
    }
    h.Remarks = dto.Remarks; h.ModifiedBy = me.Email; h.ModifiedAt = DateTime.UtcNow;

    var ids = dto.Items.Select(i => i.ItemId).ToList();
    if (ids.Count != ids.Distinct().Count()) throw new InvalidOperationException("Duplicate item");
    var items = await _db.ItemMasters.Where(i => ids.Contains(i.ItemId) && i.IsActive).ToDictionaryAsync(i => i.ItemId);
    h.RequestItems = dto.Items.Select(i => new RequestItem {
        ItemId = i.ItemId, Quantity = i.Quantity > 0 ? i.Quantity : throw new InvalidOperationException("Qty > 0"),
        Rate = items[i.ItemId].Rate, Unit = items[i.ItemId].Unit, Remarks = i.Remarks }).ToList();
    h.TotalAmount = h.RequestItems.Sum(x => x.Quantity * x.Rate);
    await _db.SaveChangesAsync();
    return h;
}
```

## Controllers
```csharp
[Authorize, ApiController, Route("api/requests")]
public class RequestsController : ControllerBase
{
    // GET api/requests?view=mine|pending|all&status=&from=&to=&deptId=&search=&page=&pageSize=
    [HttpGet] public async Task<IActionResult> List([FromQuery] ListFilter f) { /* role-aware query + paging */ }
    [HttpGet("{id}")] public async Task<IActionResult> Get(int id) { /* header + items + history */ }
    [HttpPost] public async Task<IActionResult> Save(RequestSaveDto dto) { /* RequestService.SaveAsync */ }
    [HttpPost("{id}/submit")] public async Task<IActionResult> Submit(int id) { /* Workflow.SubmitAsync */ }
    [HttpPost("{id}/action")] public async Task<IActionResult> Act(int id, ActionDto dto) { /* Workflow.ActAsync */ }
    [HttpPost("{id}/receive")] public async Task<IActionResult> Receive(int id, ReceiveDto dto) { /* requestor verify, close/partial */ }
    [HttpPost("{id}/query")] public async Task<IActionResult> Query(int id, [FromBody] string text) { /* email HRAD + log */ }
}

[Authorize, ApiController, Route("api/hrad")]
public class HradController : ControllerBase
{
    [HttpPost("{id}/process")]   // set approved qty, stock issue, vendor, PO no; generate PDF, email vendor
    public async Task<IActionResult> Process(int id, HradProcessDto dto)
    {
        var me = await _user.GetAsync();
        if (me.Role != "HRAD") return Forbid();
        var h = await _db.RequestHeaders.Include(x => x.RequestItems).FirstAsync(x => x.RequestId == id);
        if (h.Status != "HRApproval") return BadRequest("Invalid status");
        foreach (var l in dto.Lines)
        {
            var ri = h.RequestItems.First(x => x.RequestItemId == l.RequestItemId);
            ri.ApprovedQty = l.ApprovedQty; ri.IssuedFromStockQty = l.IssuedFromStockQty;
            var item = await _db.ItemMasters.FindAsync(ri.ItemId);
            if (item!.StockQty < l.IssuedFromStockQty) return BadRequest($"Insufficient stock: {item.ItemName}");
            item.StockQty -= l.IssuedFromStockQty;
            _db.StockTransactions.Add(new StockTransaction { ItemId = ri.ItemId, TxnType = "Issue",
                Quantity = l.IssuedFromStockQty, RequestId = id, CreatedBy = me.Email });
        }
        h.VendorId = dto.VendorId; h.PODate = DateTime.Today;
        h.PONumber = $"PO-{h.RequestNo}"; h.Status = "Ordered";
        _wf.AddHistory(h, "HRAD", "VendorSent", me, dto.Comments);
        await _db.SaveChangesAsync();
        var pdf = await _pdf.BuildVendorPoAsync(h.RequestId);
        await _mail.SendToVendorAsync(h, pdf);
        return Ok();
    }
    [HttpPost("{id}/deliver")]   // vendor delivered: stock Receipt not needed if issued direct; mark Delivered + notify requestor
    public async Task<IActionResult> Deliver(int id) { /* status Delivered, history Handover, email requestor */ }
    [HttpGet("{id}/pdf")] public async Task<IActionResult> Pdf(int id) =>
        File(await _pdf.BuildVendorPoAsync(id), "application/pdf", $"PO-{id}.pdf");
}

[Authorize, ApiController, Route("api/masters/{entity}")]   // employees|departments|divisions|items|vendors|categories
public class MastersController : ControllerBase
{
    // Admin-only generic CRUD. Soft delete = IsActive=false. Block deactivation if referenced by open requests.
    [HttpGet] public Task<IActionResult> List(string entity, bool? active) => throw new NotImplementedException();
    [HttpPost] public Task<IActionResult> Create(string entity, [FromBody] JsonElement body) => throw new NotImplementedException();
    [HttpPut("{id}")] public Task<IActionResult> Update(string entity, int id, [FromBody] JsonElement body) => throw new NotImplementedException();
}

[Authorize, ApiController, Route("api/dashboard")]
public class DashboardController : ControllerBase
{
    [HttpGet("summary")]  // cards: totals, pending, month cost, low stock count
    [HttpGet("monthly")]  // EXEC usp_Dash_Monthly
    [HttpGet("by-department")]
    [HttpGet("by-item")]
    [HttpGet("store")]    // vw_StoreAvailability
}

[Authorize, ApiController, Route("api/reports")]
public class ReportsController : ControllerBase
{
    [HttpGet("export")]   // same filter as list -> ClosedXML .xlsx
}
```
Implement master CRUD with explicit controllers per entity (Employees, Items, Vendors...) if you prefer strong typing over the generic route.

## PdfService (QuestPDF)
```csharp
public async Task<byte[]> BuildVendorPoAsync(int requestId)
{
    QuestPDF.Settings.License = LicenseType.Community;
    var h = await _db.RequestHeaders.Include(x => x.Vendor).Include(x => x.Requestor)
        .Include(x => x.RequestItems).ThenInclude(i => i.Item).FirstAsync(x => x.RequestId == requestId);
    return Document.Create(c => c.Page(p =>
    {
        p.Margin(30);
        p.Header().Column(col => {
            col.Item().Text("PURCHASE ORDER - STATIONERY").Bold().FontSize(16);
            col.Item().Text($"PO: {h.PONumber}   Date: {h.PODate:dd-MMM-yyyy}   Ref: {h.RequestNo}");
            col.Item().Text($"Vendor: {h.Vendor!.VendorName}, {h.Vendor.Address}");
        });
        p.Content().PaddingVertical(10).Table(t =>
        {
            t.ColumnsDefinition(c => { c.ConstantColumn(30); c.RelativeColumn(); c.ConstantColumn(60);
                                       c.ConstantColumn(60); c.ConstantColumn(70); c.ConstantColumn(80); });
            t.Header(hd => { foreach (var s in new[]{"#","Item","Unit","Qty","Rate","Amount"}) hd.Cell().Border(1).Padding(4).Text(s).Bold(); });
            int n = 1;
            foreach (var i in h.RequestItems.Where(x => (x.ApprovedQty ?? x.Quantity) - x.IssuedFromStockQty > 0))
            {
                var q = (i.ApprovedQty ?? i.Quantity) - i.IssuedFromStockQty;
                foreach (var s in new[]{ n++.ToString(), i.Item.ItemName, i.Unit, q.ToString("0.##"), i.Rate.ToString("N2"), (q*i.Rate).ToString("N2") })
                    t.Cell().Border(1).Padding(4).Text(s);
            }
        });
        p.Footer().Text($"Deliver to: HR & Admin, requested by {h.Requestor.FullName}");
    })).GeneratePdf();
}
```

## ExceptionMiddleware
```csharp
public class ExceptionMiddleware
{
    private readonly RequestDelegate _next;
    public ExceptionMiddleware(RequestDelegate next) => _next = next;
    public async Task Invoke(HttpContext ctx, AppDbContext db)
    {
        try { await _next(ctx); }
        catch (Exception ex)
        {
            db.ExceptionLogs.Add(new ExceptionLog { Source = ex.Source, Message = ex.Message,
                StackTrace = ex.ToString(), RequestPath = ctx.Request.Path,
                UserEmail = ctx.User.Identity?.Name });
            await db.SaveChangesAsync();
            ctx.Response.StatusCode = ex is UnauthorizedAccessException ? 403
                                    : ex is InvalidOperationException ? 400 : 500;
            await ctx.Response.WriteAsJsonAsync(new { error = ex is InvalidOperationException ? ex.Message : "Something went wrong" });
        }
    }
}
```

## EmailService (outline)
Insert row in `EmailLog` (Pending) -> send via Graph `/users/{sender}/sendMail` (app permission `Mail.Send`) -> update Status Sent / Failed + ErrorMessage. A small background job retries Failed rows (RetryCount < 3).
