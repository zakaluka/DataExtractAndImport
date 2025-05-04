using System.Data;
using System.Reflection;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.VisualBasic.CompilerServices;
using DataTable = System.Data.DataTable;

namespace DataExtractAndImport;

internal class SecurityRole(int securityRoleId, string name)
{
    #region Fields

    public int SecurityRoleId { get; } = securityRoleId;
    public string Name { get; set; } = name;

    public override bool Equals(object? obj) => Equals(obj as SecurityRole);

    public bool Equals(SecurityRole? obj) => obj != null && SecurityRoleId == obj.SecurityRoleId;

    public override int GetHashCode() => SecurityRoleId;

    #endregion

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

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color, List<SecurityRole> srs)
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));

        for (var i = 0; i < srs.Count; i++)
        {
            dt.Rows.Add(srs[i].SecurityRoleId, srs[i].Name);
            ws.Cell($"B{i + 2}").SetValue(srs[i].ToFriendlyString());
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
                | XLSheetProtectionElements.AutoFilter
        );
    }

    public string ToFriendlyString() => $"{Name} ({SecurityRoleId})";
}

internal class WorkQueue(int workQueueId, string name) : IComparable<WorkQueue>
{
    #region Fields

    public int WorkQueueId { get; } = workQueueId;
    public string Name { get; set; } = name;

    public int CompareTo(WorkQueue? other) =>
        other == null
            ? 1
            : string.Compare(ToFriendlyString(), other.ToFriendlyString(), StringComparison.OrdinalIgnoreCase);

    public override bool Equals(object? obj) => Equals(obj as WorkQueue);

    public bool Equals(WorkQueue? obj) => obj != null && WorkQueueId == obj.WorkQueueId;

    public override int GetHashCode() => WorkQueueId;

    #endregion

    public static List<WorkQueue> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading work queues.");
        const string sql = "select WorkqueueID, Name from WorkQueue";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<WorkQueue>();

        while (rdr.Read())
            ret.Add(new WorkQueue(rdr.GetInt32(0), rdr.GetString(1)));

        rdr.Close();
        Console.WriteLine("Finished reading work queues.");
        return ret;
    }

    public static List<WorkQueue> Transform(List<WorkQueue> q) => [.. q.OrderBy(x => x.Name.ToUpper())];

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color, List<WorkQueue> wqs)
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));

        for (var i = 0; i < wqs.Count; i++)
        {
            dt.Rows.Add(wqs[i].WorkQueueId, wqs[i].Name);
            ws.Cell($"B{i + 2}").SetValue(wqs[i].ToFriendlyString());
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.AutoFilter
                | XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
        );
    }

    public string ToFriendlyString() => $"{Name} ({WorkQueueId})";
}

internal class WorkQueueSubscription
{
    #region Fields

    public WorkQueue WorkQueue { get; set; }
    public User? User { get; set; }
    public SecurityRole? SecurityRole { get; set; }
    public SubscriptionType Type { get; set; }

    public enum SubscriptionType
    {
        User,
        Role,
    }

    #endregion

    public WorkQueueSubscription(WorkQueue workQueue, User? user)
    {
        WorkQueue = workQueue;
        User = user;
        SecurityRole = null;
        Type = SubscriptionType.User;
    }

    public WorkQueueSubscription(WorkQueue workQueue, SecurityRole? securityRole)
    {
        WorkQueue = workQueue;
        User = null;
        SecurityRole = securityRole;
        Type = SubscriptionType.Role;
    }

    public static List<WorkQueueSubscription> ReadAll(
        SqlConnection conn,
        List<WorkQueue> queues,
        List<User> users,
        List<SecurityRole> roles
    )
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
    bool emailNotification,
    string phone,
    bool phoneNotification,
    bool externalAuthentication,
    int status,
    int supervisorId
) : IComparable<User>
{
    #region Fields

    public int UserId { get; } = userId;
    public string UserName { get; set; } = userName;
    public string FirstName { get; set; } = firstName;
    public string LastName { get; set; } = lastName;
    public string EmailAddress { get; set; } = emailAddress;
    public bool EmailNotification { get; set; } = emailNotification;
    public string Phone { get; set; } = phone;
    public bool PhoneNotification { get; set; } = phoneNotification;
    public bool ExternalAuthentication { get; set; } = externalAuthentication;
    public UserStatus Status { get; set; } = (UserStatus)status;
    public int SupervisorId { get; set; } = supervisorId;
    public string Supervisor { get; set; } = "";

    public int CompareTo(User? other) =>
        other == null
            ? 1
            : string.Compare(ToFriendlyString(), other.ToFriendlyString(), StringComparison.OrdinalIgnoreCase);

    public override bool Equals(object? obj) => Equals(obj as User);

    public bool Equals(User? obj) => obj != null && UserName.Equals(obj.UserName, StringComparison.OrdinalIgnoreCase);

    public override int GetHashCode() => UserId;

    public enum UserStatus
    {
        Active = 0,
        Draft = 1,
        Archived = 2,
        Deleted = 3,
    }

    #endregion

    public static List<User> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading users.");
        const string sql =
            "select UserID, Username, FirstName, LastName, Email, EMailNotifications, MobilePhoneNumber, "
            + "TextNotifications, ExternalAuthorization, Status, Supervisor from users";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<User>();

        while (rdr.Read())
            ret.Add(
                new User(
                    rdr.GetInt32(0),
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2),
                    rdr.IsDBNull(3) ? "" : rdr.GetString(3),
                    rdr.IsDBNull(4) ? "" : rdr.GetString(4),
                    !rdr.IsDBNull(5) && rdr.GetBoolean(5),
                    rdr.IsDBNull(6) ? "" : rdr.GetString(6),
                    !rdr.IsDBNull(7) && rdr.GetBoolean(7),
                    !rdr.IsDBNull(8) && rdr.GetBoolean(8),
                    rdr.IsDBNull(9) ? 3 : rdr.GetInt32(9),
                    rdr.IsDBNull(10) ? -1 : rdr.GetInt32(10)
                )
            );

        rdr.Close();

        Console.WriteLine("Finished reading users.");
        return ret;
    }

    public static List<User> Transform(List<User> users)
    {
        foreach (var u in users)
            if (u.SupervisorId == -1)
                // No supervisor in the DB
                u.Supervisor = "";
            else if (users.Any(r => r.UserId == u.SupervisorId))
                // The supervisor ID exists in the DB, so we can use it.
                u.Supervisor = users.First(r => r.UserId == u.SupervisorId).UserName;
            else
            {
                // The supervisor ID doesn't appear to exist in the system, so we clean it up.
                u.Supervisor = "";
                u.SupervisorId = -1;
            }

        return users.OrderBy(x => x.ToFriendlyString().ToUpper()).ToList();
    }

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color, List<User> users)
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("User Name", typeof(string));
        dt.Columns.Add("First Name", typeof(string));
        dt.Columns.Add("Last Name", typeof(string));
        dt.Columns.Add("Email", typeof(string));
        dt.Columns.Add("Email Notification?", typeof(bool));
        dt.Columns.Add("Phone", typeof(string));
        dt.Columns.Add("Phone Notification?", typeof(bool));
        dt.Columns.Add("External Authentication?", typeof(bool));
        dt.Columns.Add("Status", typeof(string));
        dt.Columns.Add("Supervisor", typeof(string));

        for (var i = 0; i < users.Count; i++)
        {
            var u = users[i];
            dt.Rows.Add(
                u.UserId,
                u.UserName,
                u.FirstName,
                u.LastName,
                u.EmailAddress,
                u.EmailNotification,
                u.Phone,
                u.PhoneNotification,
                u.ExternalAuthentication,
                u.Status,
                u.Supervisor
            );
            ws.Cell($"B{i + 2}").SetValue(u.ToFriendlyString());
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
                | XLSheetProtectionElements.AutoFilter
        );
    }

    public string ToFriendlyString() => $"{FirstName} {LastName} ({UserName})";

    public static void WriteHierarchy(XLWorkbook wb, string sheetName, XLColor color, List<User> users)
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

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
        dt.Columns.Add("Level 11", typeof(string));
        dt.Columns.Add("Level 12", typeof(string));
        dt.Columns.Add("Level 13", typeof(string));
        dt.Columns.Add("Level 14", typeof(string));
        dt.Columns.Add("Level 15", typeof(string));
        dt.Columns.Add("Level 16", typeof(string));
        dt.Columns.Add("Level 17", typeof(string));
        dt.Columns.Add("Level 18", typeof(string));
        dt.Columns.Add("Level 19", typeof(string));
        dt.Columns.Add("Level 20", typeof(string));

        // Set up the supervisory hierarchy, Supervisor => Supervisee.
        var hierarchy = new Dictionary<User, List<User>>();

        foreach (var u in users)
            hierarchy[u] = users.Where(x => x.SupervisorId == u.UserId && x.UserId != u.UserId).ToList();

        // Find the roots aka unsupervised users - those who don't have one or who are their own supervisor.
        var roots = users.Where(u => u.SupervisorId == -1 || u.UserId == u.SupervisorId);

        // Populate the data table's contents.
        foreach (var root in roots)
            RecursiveWriter(dt, 0, root, null, hierarchy);

        // Clean up any columns in the DataTable that are empty.
        var colsToRemove = dt
            .Columns.Cast<DataColumn>()
            .Where(column => dt.AsEnumerable().Select(row => row.Field<string>(column)).All(string.IsNullOrWhiteSpace))
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < colsToRemove.Count; i++)
            dt.Columns.Remove(colsToRemove[i]);

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
                | XLSheetProtectionElements.AutoFilter
        );
    }

    private static void RecursiveWriter(
        DataTable dt,
        int column,
        User u,
        DataRow? parentRow,
        Dictionary<User, List<User>> hierarchy
    )
    {
        // Write out the current DL first. Insert empty strings for columns that should be empty.
        List<object> dtRow = [];

        for (var i = 0; i < column; i++)
            dtRow.Add("");

        dtRow.Add(u.ToFriendlyString());

        for (var i = dtRow.Count; i < dt.Columns.Count; i++)
            dtRow.Add("");

        var row = dt.Rows.Add(dtRow.ToArray());
        row.SetParentRow(parentRow);

        var mySupervisees = hierarchy[u];

        // Write out each child.
        foreach (var s in mySupervisees)
            RecursiveWriter(dt, column + 1, s, row, hierarchy);
    }
}

internal class Datalist(int datalistId, string name, string systemName)
{
    #region Fields

    public int DatalistId { get; } = datalistId;
    public string Name { get; set; } = name;
    public string SystemName { get; set; } = systemName;

    public override bool Equals(object? obj) => Equals(obj as Datalist);

    public bool Equals(Datalist? obj) =>
        obj != null && SystemName.Equals(obj.SystemName, StringComparison.OrdinalIgnoreCase);

    public override int GetHashCode() => DatalistId;

    #endregion

    public static List<Datalist> ReadAll(SqlConnection conn)
    {
        Console.WriteLine("Started reading datalists.");
        const string sql = "select DataListID, Name, SystemName from DataList where Infrastructure = 0";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<Datalist>();

        while (rdr.Read())
            ret.Add(
                new Datalist(
                    rdr.GetInt32(0),
                    rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                    rdr.IsDBNull(2) ? "" : rdr.GetString(2)
                )
            );

        rdr.Close();
        Console.WriteLine("Finished reading datalists.");
        return ret;
    }

    public static List<Datalist> Transform(List<Datalist> dls) => dls.OrderBy(x => x.Name.ToUpper()).ToList();

    public static void WriteAll(XLWorkbook wb, string sheetName, XLColor color, List<Datalist> dls)
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

        // Populate data table
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("System Name", typeof(string));

        for (var i = 0; i < dls.Count; i++)
        {
            dt.Rows.Add(dls[i].DatalistId, dls[i].Name, dls[i].SystemName);
            ws.Cell($"B{i + 2}").SetValue(dls[i].ToFriendlyString());
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
                | XLSheetProtectionElements.AutoFilter
        );
    }

    public string ToFriendlyString() => $"{Name} ({DatalistId})";
}

internal class ListRelationship(Datalist parent, Datalist child)
{
    #region Fields

    public Datalist Parent { get; } = parent;
    public Datalist Child { get; } = child;

    public override bool Equals(object? obj) => Equals(obj as ListRelationship);

    public bool Equals(ListRelationship? obj) => obj != null && Parent.Equals(obj.Parent) && Child.Equals(obj.Child);

    public override int GetHashCode() => int.Parse(Parent.GetHashCode() + "000" + Child.GetHashCode());

    #endregion

    public static List<ListRelationship> ReadAll(SqlConnection conn, List<Datalist> dls)
    {
        Console.WriteLine("Started reading list relationships.");
        const string sql = "select ParentListID, ChildListID from ListRelationship";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<ListRelationship>();

        Datalist p,
            c;
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

    public static void WriteAll(
        XLWorkbook wb,
        string sheetName,
        XLColor color,
        List<ListRelationship> listRelationships,
        List<Datalist> dls
    )
    {
        var ws = Utility.GetStandardWorkSheet(wb, sheetName, color);

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
        dt.Columns.Add("Level 11", typeof(string));
        dt.Columns.Add("Level 12", typeof(string));
        dt.Columns.Add("Level 13", typeof(string));
        dt.Columns.Add("Level 14", typeof(string));
        dt.Columns.Add("Level 15", typeof(string));
        dt.Columns.Add("Level 16", typeof(string));
        dt.Columns.Add("Level 17", typeof(string));
        dt.Columns.Add("Level 18", typeof(string));
        dt.Columns.Add("Level 19", typeof(string));
        dt.Columns.Add("Level 20", typeof(string));

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

        // Populate the data table's contents.
        foreach (var root in roots)
            RecursiveWriter(dt, 0, root, children);

        // Clean up any columns in the DataTable that are empty.
        var colsToRemove = dt
            .Columns.Cast<DataColumn>()
            .Where(column => dt.AsEnumerable().Select(row => row.Field<string>(column)).All(string.IsNullOrWhiteSpace))
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < colsToRemove.Count; i++)
            dt.Columns.Remove(colsToRemove[i]);

        ws.Cell("c1").InsertTable(dt, sheetName, true);

        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);
        ws.Protect(
            allowedElements: XLSheetProtectionElements.FormatEverything
                | XLSheetProtectionElements.SelectEverything
                | XLSheetProtectionElements.Sort
                | XLSheetProtectionElements.AutoFilter
        );
    }

    // Writes a DL and its children in a depth first manner.
    private static void RecursiveWriter(
        DataTable dt,
        int column,
        in Datalist dl,
        in Dictionary<Datalist, IEnumerable<Datalist>> children
    )
    {
        // Write out the current DL first.
        List<object> dtRow = [];

        for (var i = 0; i < column; i++)
            dtRow.Add("");

        dtRow.Add(dl.ToFriendlyString());

        for (var i = dtRow.Count; i < dt.Columns.Count; i++)
            dtRow.Add("");

        dt.Rows.Add(dtRow.ToArray());

        var myChildren = children[dl];

        // Write out each child.
        foreach (var c in myChildren)
            RecursiveWriter(dt, column + 1, c, children);
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

    private bool IsNoAccess() =>
        !View && !Add && !Edit && !BulkEdit && !Delete && !ViewActivity && !Merge && !Move && !Administer;

    private bool IsViewOnly() =>
        View && !Add && !Edit && !BulkEdit && !Delete && !ViewActivity && !Merge && !Move && !Administer;

    public override string ToString()
    {
        if (IsNoAccess())
            return "";
        if (Administer)
            return "Administer";
        if (IsViewOnly())
            return "View Only";
        return (View ? "View" : "")
            + (Add ? ", Add" : "")
            + (Edit ? ", Edit" : "")
            + (BulkEdit ? ", Bulk Edit" : "")
            + (Delete ? ", Delete" : "")
            + (ViewActivity ? ", Activity Wall" : "")
            + (Merge ? ", Merge" : "")
            + (Move ? ", Move" : "");
    }

    public static Permission Parse(string text) => throw new NotImplementedException();

    public void Write() => throw new NotImplementedException();
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
            if (!rdr.IsDBNull(2) && rdr.GetBoolean(2))
                perm.AddAdd();
            if (!rdr.IsDBNull(3) && rdr.GetBoolean(3))
                perm.AddEdit();
            if (!rdr.IsDBNull(4) && rdr.GetBoolean(4))
                perm.AddBulkEdit();
            if (!rdr.IsDBNull(5) && rdr.GetBoolean(5))
                perm.AddDelete();
            if (!rdr.IsDBNull(6) && rdr.GetBoolean(6))
                perm.AddViewActivity();
            if (!rdr.IsDBNull(7) && rdr.GetBoolean(7))
                perm.AddMerge();
            if (!rdr.IsDBNull(8) && rdr.GetBoolean(8))
                perm.AddMove();
            if (!rdr.IsDBNull(9) && rdr.GetBoolean(9))
                perm.AddAdminister();
            ret.Add(new ListRole(dl, role, perm));
        }

        rdr.Close();
        Console.WriteLine("Finished reading list roles.");
        return ret;
    }

    public static List<ListRole> Transform(List<ListRole> lrs) => lrs;

    public static void Write() => throw new NotImplementedException();
}

internal static class Utility
{
    public static IXLWorksheet GetStandardWorkSheet(XLWorkbook wb, string sheetName, XLColor color)
    {
        // Get/create the worksheet
        if (!wb.TryGetWorksheet(sheetName, out var ws))
            ws = wb.AddWorksheet(sheetName);
        ws.TabColor = color;

        // Set sheet header name
        ws.Cell("A1").SetValue(sheetName).Style.Font.SetBold(true);

        return ws;
    }
}

internal static class Matrix
{
    public static void WriteWorkQueueSubscriptions(
        XLWorkbook wb,
        string UserXWorkQueueSheetName,
        string WorkQueueXUserSheetName,
        string RoleXWorkQueueSheetName,
        string WorkQueueXRoleSheetName,
        XLColor color,
        List<WorkQueueSubscription> subscriptions
    )
    {
        // Matrix
        WriteUserXWorkQueueSubscription(
            Utility.GetStandardWorkSheet(wb, UserXWorkQueueSheetName, color),
            UserXWorkQueueSheetName,
            subscriptions
        );

        WriteWorkQueueXUserSubscription(
            Utility.GetStandardWorkSheet(wb, WorkQueueXUserSheetName, color),
            UserXWorkQueueSheetName,
            subscriptions
        );
    }

    private static void WriteUserXWorkQueueSubscription(
        IXLWorksheet ws,
        string sheetName,
        List<WorkQueueSubscription> subscriptions
    )
    {
        var dt = new DataTable();
        dt.TableName = sheetName;
        dt.Columns.Add("User", typeof(string));
        foreach (var wq in subscriptions.Select(x => x.WorkQueue).Distinct().Order())
            dt.Columns.Add(wq.ToFriendlyString(), typeof(string));

        // Populate the table.
        foreach (
            var u in subscriptions
                .Where(x =>
                    x.Type == WorkQueueSubscription.SubscriptionType.User && x.User!.Status == User.UserStatus.Active
                )
                .Select(x => x.User!)
                .Distinct()
                .Order()
        )
        {
            var r = dt.Rows.Add();
            r.SetField("User", u.ToFriendlyString());
            foreach (var subs in subscriptions.Where(x => u.Equals(x.User)))
                r.SetField(subs.WorkQueue.ToFriendlyString(), "x");
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);
        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);

        // Clean up any columns in the DataTable that are empty.
        var colsToRemove = dt
            .Columns.Cast<DataColumn>()
            .Where(column => dt.AsEnumerable().Select(row => row.Field<string>(column)).All(string.IsNullOrWhiteSpace))
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < colsToRemove.Count; i++)
            dt.Columns.Remove(colsToRemove[i]);

        // Clean up any rows in the DataTable that are empty.
        var rowsToRemove = dt
            .Rows.Cast<DataRow>()
            .Where(r =>
                r.ItemArray.Count(val =>
                    val != null && val.GetType() != typeof(DBNull) && !string.IsNullOrWhiteSpace((string)val)
                ) == 1
            )
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < rowsToRemove.Count; i++)
            dt.Rows.Remove(rowsToRemove[i]);

        ws.Range(2, 4, ws.RangeUsed()!.LastRowUsed().RowNumber(), ws.RangeUsed()!.LastColumnUsed().ColumnNumber())
            .AddConditionalFormat()
            .WhenNotBlank()
            .Fill.SetBackgroundColor(XLColor.Black)
            .Font.SetFontColor(XLColor.Black);
    }

    private static void WriteWorkQueueXUserSubscription(
        IXLWorksheet ws,
        string sheetName,
        List<WorkQueueSubscription> subscriptions
    )
    {
        var dt = new DataTable();

        // Add the table' columns
        dt.TableName = sheetName;
        dt.Columns.Add("Work Queue", typeof(string));
        foreach (
            var u in subscriptions
                .Where(x =>
                    x.Type == WorkQueueSubscription.SubscriptionType.User && x.User!.Status == User.UserStatus.Active
                )
                .Select(x => x.User!)
                .Distinct()
                .Order()
        )
            dt.Columns.Add(u.ToFriendlyString(), typeof(string));

        // Populate the table's rows.
        foreach (
            var wq in subscriptions
                .Where(x =>
                    x.Type == WorkQueueSubscription.SubscriptionType.User && x.User!.Status == User.UserStatus.Active
                )
                .Select(x => x.WorkQueue)
                .Distinct()
                .Order()
        )
        {
            var r = dt.Rows.Add();
            r.SetField("Work Queue", wq.ToFriendlyString());
            foreach (
                var subs in subscriptions
                    .Where(x =>
                        x.Type == WorkQueueSubscription.SubscriptionType.User
                        && x.User!.Status == User.UserStatus.Active
                    )
                    .Where(x => wq.Equals(x.WorkQueue))
            )
                r.SetField(subs.User!.ToFriendlyString(), "x");
        }

        ws.Cell("c1").InsertTable(dt, sheetName, true);
        ws.ColumnsUsed().AdjustToContents();
        ws.Column("b").Hide();
        ws.SheetView.FreezeRows(1);

        // Clean up any columns in the DataTable that are empty.
        var colsToRemove = dt
            .Columns.Cast<DataColumn>()
            .Where(column => dt.AsEnumerable().Select(row => row.Field<string>(column)).All(string.IsNullOrWhiteSpace))
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < colsToRemove.Count; i++)
            dt.Columns.Remove(colsToRemove[i]);

        // Clean up any rows in the DataTable that are empty.
        var rowsToRemove = dt
            .Rows.Cast<DataRow>()
            .Where(r =>
                r.ItemArray.Count(val =>
                    val != null && val.GetType() != typeof(DBNull) && !string.IsNullOrWhiteSpace((string)val)
                ) == 1
            )
            .ToList();
        // NOTE: Cannot convert to foreach as both colsToRemove and dt.Columns are referring to the same DataColumn
        // entries behind the scenes. This means that when we remove something from dt.Columns, we also indirectly
        // modify the entries in colsToRemove.
        // ReSharper disable once ForCanBeConvertedToForeach
        for (var i = 0; i < rowsToRemove.Count; i++)
            dt.Rows.Remove(rowsToRemove[i]);

        ws.Range(2, 4, ws.RangeUsed()!.LastRowUsed().RowNumber(), ws.RangeUsed()!.LastColumnUsed().ColumnNumber())
            .AddConditionalFormat()
            .WhenNotBlank()
            .Fill.SetBackgroundColor(XLColor.Black)
            .Font.SetFontColor(XLColor.Black);
    }
}

internal static class Export
{
    #region Fields

    private const string DlSheetName = "Datalists";
    private const string LrSheetName = "Datalist Hierarchy";
    private static readonly XLColor DlSheetsColor = XLColor.LightBlue;

    private const string UserSheetName = "Users";
    private const string UserHierarchySheetName = "User Hierarchy";
    private static readonly XLColor UserSheetsColor = XLColor.LightGreen;

    private const string RoleSheetName = "Security Roles";
    private static readonly XLColor SecuritySheetsColor = XLColor.LightYellow;

    private const string QueueSheetName = "Work Queues";
    private static readonly XLColor WorkQueueSheetsColor = XLColor.LightPink;

    // Naming convention: Row x Column.
    private const string UserXRoleMatrixSheetName = "User-Role Matrix";
    private const string DlXRoleMatrixSheetName = "Datalist-Role Matrix";
    private const string DlXWorkQueueMatrixSheetName = "Datalist-Queue Matrix";
    private const string UserXWorkQueueMatrixSheetName = "User-Queue Matrix";
    private const string RoleXWorkQueueMatrixSheetName = "Role-Queue Matrix";
    private const string WorkQueueXUserMatrixSheetName = "Queue-User Matrix";
    private const string WorkQueueXRoleMatrixSheetName = "Queue-Role Matrix";
    private static readonly XLColor MatrixColor = XLColor.LightApricot;

    #endregion

    public static void Run(string connectionString, string filename)
    {
        Console.WriteLine("Program starts.");

        // Excel workbook
        using var wb = new XLWorkbook();

        // Open a DB connection
        using var conn = new SqlConnection(connectionString);
        conn.Open();

        // Get and read the data
        var securityRoles = SecurityRole.Transform(SecurityRole.ReadAll(conn));
        var workQueues = WorkQueue.Transform(WorkQueue.ReadAll(conn));
        var users = User.Transform(User.ReadAll(conn));
        var workQueueSubscriptions = WorkQueueSubscription.Transform(
            WorkQueueSubscription.ReadAll(conn, workQueues, users, securityRoles)
        );
        var dls = Datalist.Transform(Datalist.ReadAll(conn));
        var listRelationships = ListRelationship.Transform(ListRelationship.ReadAll(conn, dls));
        var listRoles = ListRole.Transform(ListRole.ReadAll(conn, dls, securityRoles));

        // Write out the data in Excel sheets.
        Datalist.WriteAll(wb, DlSheetName, DlSheetsColor, dls);
        ListRelationship.WriteAll(wb, LrSheetName, DlSheetsColor, listRelationships, dls);
        User.WriteAll(wb, UserSheetName, UserSheetsColor, users);
        User.WriteHierarchy(wb, UserHierarchySheetName, UserSheetsColor, users);
        SecurityRole.WriteAll(wb, RoleSheetName, SecuritySheetsColor, securityRoles);
        WorkQueue.WriteAll(wb, QueueSheetName, WorkQueueSheetsColor, workQueues);
        Matrix.WriteWorkQueueSubscriptions(
            wb,
            UserXWorkQueueMatrixSheetName,
            WorkQueueXUserMatrixSheetName,
            RoleXWorkQueueMatrixSheetName,
            WorkQueueXRoleMatrixSheetName,
            MatrixColor,
            workQueueSubscriptions
        );

        // Write out the Excel file.
        using var fs = File.Create(
            Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location)
                + Path.DirectorySeparatorChar.ToString()
                + filename
        );
        wb.SaveAs(fs);

        Console.WriteLine("Program complete.");
    }
}
