using System.Reflection;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Data.SqlClient;
using DataTable = System.Data.DataTable;

namespace DataExtractAndImport;

internal class SecurityRole(int securityRoleId, string name)
{
    public int SecurityRoleId { get; } = securityRoleId;
    public string Name { get; set; } = name;

    public override bool Equals(object? obj) => Equals(obj as SecurityRole);

    public bool Equals(SecurityRole? obj) => obj != null && SecurityRoleId == obj.SecurityRoleId;

    public override int GetHashCode() => SecurityRoleId;

    public static List<SecurityRole> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading security roles.");
        const string sql = "select or2.OrganizationRoleID, or2.Name from OrganizationRole or2";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<SecurityRole>();

        while (rdr.Read())
            ret.Add(new SecurityRole(rdr.GetInt32(0), rdr.GetString(1)));

        rdr.Close();
        Console.WriteLine("Finished reading security roles.");
        return ret;
    }

    public static List<SecurityRole> Transform(List<SecurityRole> roles) => [.. roles.OrderBy(x => x.Name.ToUpper())];

    public static bool WriteAll() => throw new NotImplementedException();
}

internal class WorkQueue(int workQueueId, string name)
{
    public int WorkQueueId { get; set; } = workQueueId;
    public string Name { get; set; } = name;

    public static List<WorkQueue> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading work queues.");
        const string sql = "select WorkqueueID, Name from WorkQueue";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<WorkQueue>();

        while (rdr.Read()) ret.Add(new WorkQueue(rdr.GetInt32(0), rdr.GetString(1)));

        rdr.Close();
        Console.WriteLine("Finished reading work queues.");
        return ret;
    }

    public static List<WorkQueue> Transform(List<WorkQueue> q) => [.. q.OrderBy(x => x.Name.ToUpper())];

    public static bool WriteAll() => throw new NotImplementedException();
}

internal enum SubscriptionType
{
    User,
    Role
}

internal class WorkQueueSubscription
{
    public WorkQueue WorkQueue { get; set; }
    public User? User { get; set; }
    public SecurityRole? SecurityRole { get; set; }
    public SubscriptionType SubscriptionType { get; set; }

    public WorkQueueSubscription(WorkQueue workQueue, User? user)
    {
        WorkQueue = workQueue;
        User = user;
        SecurityRole = null;
        SubscriptionType = SubscriptionType.User;
    }

    public WorkQueueSubscription(WorkQueue workQueue, SecurityRole? securityRole)
    {
        WorkQueue = workQueue;
        User = null;
        SecurityRole = securityRole;
        SubscriptionType = SubscriptionType.Role;
    }

    public static List<WorkQueueSubscription> ReadAll(SqlConnection conn, List<WorkQueue> queues, List<User> users,
        List<SecurityRole> roles)
    {
        Console.WriteLine("Started reading queue subscriptions - user.");
        var sql = "select WorkQueueID, UserID from WorkQueueMember";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<WorkQueueSubscription>();

        while (rdr.Read())
        {
            var q = queues.FirstOrDefault(x => x.WorkQueueId == rdr.GetInt32(0));
            var u = users.FirstOrDefault(x => x.UserId == rdr.GetInt32(1));
            if (u == null || q == null)
            {
                Console.WriteLine($"Unable to find queue {rdr.GetInt32(0)} or user {rdr.GetInt32(1)}");
                continue;
            }

            ret.Add(new WorkQueueSubscription(q, u));
        }

        rdr.Close();
        Console.WriteLine("Finished reading queue subscriptions - user.");

        Console.WriteLine("Started reading queue subscriptions - role.");
        sql = "select WorkQueueID, RoleID from WorkQueueRole";
        using var cmd2 = new SqlCommand(sql, conn);
        rdr = cmd2.ExecuteReader();

        while (rdr.Read())
        {
            var q = queues.FirstOrDefault(x => x.WorkQueueId == rdr.GetInt32(0));
            var r = roles.FirstOrDefault(x => x.SecurityRoleId == rdr.GetInt32(1));
            if (r == null || q == null)
            {
                Console.WriteLine($"Unable to find queue {rdr.GetInt32(0)} or role {rdr.GetInt32(1)}");
                continue;
            }

            ret.Add(new WorkQueueSubscription(q, r));
        }

        rdr.Close();
        Console.WriteLine("Finished reading queue subscriptions - role.");
        return ret;
    }

    public static List<WorkQueueSubscription> Transform(List<WorkQueueSubscription> wqs) => wqs;
    public static void WriteAll() => throw new NotImplementedException();
}

internal class User(
    int userId,
    string userName,
    string firstName,
    string lastName,
    string emailAddress,
    string phone,
    bool externalAuthentication,
    int active,
    int supervisorId)
{
    public int UserId { get; set; } = userId;
    public string UserName { get; set; } = userName;
    public string FirstName { get; set; } = firstName;
    public string LastName { get; set; } = lastName;
    public string EmailAddress { get; set; } = emailAddress;
    public string Phone { get; set; } = phone;
    public bool ExternalAuthentication { get; set; } = externalAuthentication;
    public int Active { get; set; } = active;
    public int SupervisorId { get; set; } = supervisorId;
    public User? Supervisor { get; set; } = null;

    public static List<User> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading users.");
        const string sql =
            "select UserID, Username, FirstName, LastName, Email, MobilePhoneNumber, ExternalAuthorization, Status, Supervisor from users";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<User>();

        while (rdr.Read())
            ret.Add(new User(rdr.GetInt32(0), rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                rdr.IsDBNull(3) ? "" : rdr.GetString(3), rdr.IsDBNull(4) ? "" : rdr.GetString(4),
                rdr.IsDBNull(5) ? "" : rdr.GetString(5), rdr.GetBoolean(6), rdr.GetInt32(7),
                rdr.IsDBNull(8) ? -1 : rdr.GetInt32(8)));

        rdr.Close();

        foreach (var u in ret)
            u.Supervisor = u.SupervisorId == -1 ? u : ret.FirstOrDefault((x => x.UserId == u.SupervisorId), u);

        Console.WriteLine("Finished reading users.");
        return ret;
    }

    public static List<User> Transform(List<User> users) =>
        users.OrderBy(x => $"{x.FirstName.ToUpper()} {x.LastName.ToUpper()} {x.UserName.ToUpper()}").ToList();

    public static void WriteAll() => throw new NotImplementedException();
}

internal class Datalist(int datalistId, string name, string systemName)
{
    public int DatalistId { get; set; } = datalistId;
    public string Name { get; set; } = name;
    public string SystemName { get; set; } = systemName;

    public static List<Datalist> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading datalists.");
        const string sql = "select DataListID, Name, SystemName from DataList where Infrastructure = 0";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<Datalist>();

        while (rdr.Read())
            ret.Add(new Datalist(rdr.GetInt32(0), rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                rdr.IsDBNull(2) ? "" : rdr.GetString(2)));

        rdr.Close();
        Console.WriteLine("Finished reading datalists.");
        return ret;
    }

    public static List<Datalist> Transform(List<Datalist> dls) => dls.OrderBy(x => x.Name.ToUpper()).ToList();

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color, List<Datalist> dls)
    {
        // Get/create the worksheet
        if (!wb.TryGetWorksheet(sheetName, out var ws))
            ws = wb.AddWorksheet(sheetName);
        ws.TabColor = color;

        // Set sheet header name
        ws.Cell("A1").SetValue(sheetName).Style.Font.SetBold(true);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("System Name", typeof(string));

        foreach (var dl in dls)
            dt.Rows.Add(dl.DatalistId, dl.Name, dl.SystemName);

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        // Set up column with lookup values.
        ws.Range($"b2:b{dt.Rows.Count + 1}").FormulaArrayA1 = "=Datalists[Name]&\" (\"&Datalists[ID]&\")\"";

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(allowedElements: XLSheetProtectionElements.FormatEverything |
                                    XLSheetProtectionElements.SelectEverything | XLSheetProtectionElements.Sort |
                                    XLSheetProtectionElements.AutoFilter);
    }
}

internal class ListRelationship(Datalist parent, Datalist child)
{
    public Datalist Parent { get; set; } = parent;
    public Datalist Child { get; set; } = child;

    public static List<ListRelationship> ReadAll(SqlConnection conn, List<Datalist> dls)
    {
        Console.WriteLine("Started reading list relationships.");
        const string sql = "select ParentListID, ChildListID from ListRelationship";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<ListRelationship>();

        Datalist p, c;
        while (rdr.Read())
        {
            p = dls.Find(x => x.DatalistId == rdr.GetInt32(0));
            c = dls.Find(x => x.DatalistId == rdr.GetInt32(1));

            // Found a list relationship which is not relevant (e.g. for an infrastructure list).
            if (p == null || c == null)
                continue;

            ret.Add(new ListRelationship(p, c));
        }

        rdr.Close();
        Console.WriteLine("Finished reading list relationships.");
        return ret;
    }

    public static List<ListRelationship> Transform(List<ListRelationship> lrs) => lrs;

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color,
        List<ListRelationship> listRelationships, List<Datalist> dls)

    {
        // Get/create the worksheet
        if (!wb.TryGetWorksheet(sheetName, out var ws))
            ws = wb.AddWorksheet(sheetName);
        ws.TabColor = color;

        // Set sheet header name
        ws.Cell("A1").SetValue(sheetName).Style.Font.SetBold(true);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("Level 1", typeof(string));
        dt.Columns.Add("Level 2", typeof(string));
        dt.Columns.Add("Level 3", typeof(string));
        dt.Columns.Add("Level 4", typeof(string));
        dt.Columns.Add("Level 5", typeof(string));
        dt.Columns.Add("Level 6", typeof(string));
        dt.Columns.Add("Level 7", typeof(string));
        dt.Columns.Add("Level 8", typeof(string));
        dt.Columns.Add("Level 9", typeof(string));
        dt.Columns.Add("Level 10", typeof(string));

        // Set up the mapping of a DL to its children.  If a DL has no children, it is mapped to the empty list.
        var children = new Dictionary<Datalist, IEnumerable<Datalist>>();
        // Add the DLs to `children` where the DL has no child.
        foreach (var dl in dls.Where(x => listRelationships.All(lr => lr.Parent.DatalistId != x.DatalistId)))
            children[dl] = new List<Datalist>();
        // Add the DLs to `children` where the DL has 1+ children.
        foreach (var dl in dls.Where(x => listRelationships.Any(lr => lr.Parent.DatalistId == x.DatalistId)))
            children[dl] = listRelationships.Where(lr => lr.Parent.DatalistId == dl.DatalistId).Select(lr => lr.Child);

        // Find the roots aka Level 1 entries - these are workspaces that are never children in the list relationships
        var roots = dls.Where(dl => listRelationships.All(lr => lr.Child.DatalistId != dl.DatalistId));
        
        // ----------- 
        throw new Exception("Continue from here");

        foreach (var dl in dls)
            dt.Rows.Add(dl.DatalistId, dl.Name, dl.SystemName);

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        // Set up column with lookup values.
        ws.Range($"b2:b{dt.Rows.Count + 1}").FormulaArrayA1 = "=Datalists[Name]&\" (\"&Datalists[ID]&\")\"";

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(allowedElements: XLSheetProtectionElements.FormatEverything |
                                    XLSheetProtectionElements.SelectEverything | XLSheetProtectionElements.Sort |
                                    XLSheetProtectionElements.AutoFilter);
    }
}

internal class Permission
{
    private bool View { get; set; } = false;
    private bool Add { get; set; } = false;
    private bool Edit { get; set; } = false;
    private bool BulkEdit { get; set; } = false;
    private bool Delete { get; set; } = false;
    private bool ViewActivity { get; set; } = false;
    private bool Merge { get; set; } = false;
    private bool Move { get; set; } = false;
    private bool Administer { get; set; } = false;

    public Permission AddAdd()
    {
        View = true;
        Add = true;
        return this;
    }

    public Permission AddEdit()
    {
        View = true;
        Edit = true;
        return this;
    }

    public Permission AddBulkEdit()
    {
        View = true;
        Edit = true;
        BulkEdit = true;
        return this;
    }

    public Permission AddDelete()
    {
        View = true;
        Delete = true;
        return this;
    }

    public Permission AddViewActivity()
    {
        View = true;
        ViewActivity = true;
        return this;
    }

    public Permission AddMerge()
    {
        View = true;
        Edit = true;
        Delete = true;
        Merge = true;
        return this;
    }

    public Permission AddMove()
    {
        View = true;
        Move = true;
        return this;
    }

    public Permission AddView()
    {
        View = true;
        return this;
    }

    public Permission AddAdminister()
    {
        View = true;
        Add = true;
        Edit = true;
        BulkEdit = true;
        Delete = true;
        ViewActivity = true;
        Merge = true;
        Move = true;
        Administer = true;
        return this;
    }

    public Permission AddNoAccess()
    {
        View = false;
        Add = false;
        Edit = false;
        BulkEdit = false;
        Delete = false;
        ViewActivity = false;
        Merge = false;
        Move = false;
        Administer = false;
        return this;
    }

    private bool IsNoAccess() => !View && !Add && !Edit && !BulkEdit && !Delete && !ViewActivity && !Merge && !Move &&
                                 !Administer;

    private bool IsViewOnly() =>
        View && !Add && !Edit && !BulkEdit && !Delete && !ViewActivity && !Merge && !Move && !Administer;

    public override string ToString()
    {
        if (IsNoAccess()) return "N/A";
        if (Administer) return "Administer";
        if (IsViewOnly()) return "View Only";
        return (View ? "View" : "") + (Add ? ", Add" : "") + (Edit ? ", Edit" : "") + (BulkEdit ? ", Bulk Edit" : "") +
               (Delete ? ", Delete" : "") + (ViewActivity ? ", Activity Wall" : "") + (Merge ? ", Merge" : "") +
               (Move ? ", Move" : "");
    }

    public static Permission Parse(string text) => throw new NotImplementedException();
    public static void WriteAll() => throw new NotImplementedException();
}

internal class ListRole(Datalist datalist, SecurityRole securityRole, Permission permission)
{
    public Datalist Datalist { get; set; } = datalist;
    public SecurityRole SecurityRole { get; set; } = securityRole;
    public Permission Permission { get; set; } = permission;

    public static List<ListRole> ReadAll(SqlConnection conn, List<Datalist> dls, List<SecurityRole> roles)
    {
        Console.WriteLine("Started reading list roles.");
        const string sql =
            "select ListID, RoleID, AllowAddInd, AllowEditInd, AllowBulkEditInd, AllowDeletedInd, AllowActivityWallInd, AllowMergeInd, AllowMoveInd, AdministerInd from ListRole";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<ListRole>();

        while (rdr.Read())
        {
            var dl = dls.Find(x => x.DatalistId == rdr.GetInt32(0));

            // If DL is invalid (e.g. doesn't exist, is an infrastructure list, etc.)
            if (dl == null)
                continue;

            var role = roles.First(x => x.SecurityRoleId == rdr.GetInt32(1));
            var perm = (new Permission()).AddView();
            if (!rdr.IsDBNull(2) && rdr.GetBoolean(2)) perm.AddAdd();
            if (!rdr.IsDBNull(3) && rdr.GetBoolean(3)) perm.AddEdit();
            if (!rdr.IsDBNull(4) && rdr.GetBoolean(4)) perm.AddBulkEdit();
            if (!rdr.IsDBNull(5) && rdr.GetBoolean(5)) perm.AddDelete();
            if (!rdr.IsDBNull(6) && rdr.GetBoolean(6)) perm.AddViewActivity();
            if (!rdr.IsDBNull(7) && rdr.GetBoolean(7)) perm.AddMerge();
            if (!rdr.IsDBNull(8) && rdr.GetBoolean(8)) perm.AddMove();
            if (!rdr.IsDBNull(9) && rdr.GetBoolean(9)) perm.AddAdminister();
            ret.Add(new ListRole(dl, role, perm));
        }

        rdr.Close();
        Console.WriteLine("Finished reading list roles.");
        return ret;
    }

    public static List<ListRole> Transform(List<ListRole> lrs) => lrs;
    public static void WriteAll() => throw new NotImplementedException();
}

internal static class Export
{
    private const string DlSheetName = "Datalists";
    private static string LrSheetName = "Datalist Hierarchy";
    private static readonly XLColor DlSheetsColor = XLColor.LightBlue;

    private static string userSheetName = "Users-RO";
    private static string userHierarchySheetName = "UserHierarchy-RO";
    private static string roleSheetName = "SecurityRoles";
    private static string secMatrixSheetName = "SecurityMatrix";
    private static string queueSheetName = "WorkQueues";
    private static string workQueueMatrixSheetName = "WorkQueueMatrix";

    public static void Run(string connectionString, string filename)
    {
        // Excel workbook
        using var wb = new XLWorkbook();

        // Open a DB connection
        using var conn = new SqlConnection(connectionString);
        conn.Open();

        // Get and read the data
        var securityRoles = SecurityRole.Transform(SecurityRole.ReadAll(conn));
        var workQueues = WorkQueue.Transform(WorkQueue.ReadAll(conn));
        var users = User.Transform(User.ReadAll(conn));
        var workQueueSubscriptions =
            WorkQueueSubscription.Transform(WorkQueueSubscription.ReadAll(conn, workQueues, users, securityRoles));
        var dls = Datalist.Transform(Datalist.ReadAll(conn));
        var listRelationships = ListRelationship.Transform(ListRelationship.ReadAll(conn, dls));
        var listRoles = ListRole.Transform(ListRole.ReadAll(conn, dls, securityRoles));

        Datalist.WriteAll(wb, DlSheetName, DlSheetsColor, dls);
        ListRelationship.WriteAll(wb, LrSheetName, DlSheetsColor, listRelationships, dls);

        // Write out the Excel file.
        using var fs = File.Create(Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location) +
                                   Path.DirectorySeparatorChar.ToString() + filename);
        wb.SaveAs(fs);
    }
}