using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace DataExtractAndImport;

internal class SecurityRole(int securityRoleId, string name)
{
    public int SecurityRoleId { get; } = securityRoleId;
    public string Name { get; set; } = name;

    public override bool Equals(object? obj) => Equals(obj as SecurityRole);

    public bool Equals(SecurityRole? obj) => obj != null && SecurityRoleId == obj.SecurityRoleId;

    public override int GetHashCode() => SecurityRoleId;

    public static List<SecurityRole> Read(SqlConnection conn)
    {
        Console.WriteLine("Started reading security roles.");
        const string sql = "select or2.OrganizationRoleID, or2.Name from OrganizationRole or2";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<SecurityRole>();

        while (rdr.Read()) ret.Add(new SecurityRole(rdr.GetInt32(0), rdr.GetString(1)));

        rdr.Close();
        Console.WriteLine("Finished reading security roles.");
        return ret;
    }

    public static List<SecurityRole> Transform(List<SecurityRole> roles) => [.. roles.OrderBy(x => x.Name.ToUpper())];

    public static bool Write() => throw new NotImplementedException();
}

internal class WorkQueue(int workQueueId, string name)
{
    public int WorkQueueId { get; set; } = workQueueId;
    public string Name { get; set; } = name;

    public static List<WorkQueue> Read(SqlConnection conn)
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

    public static bool Write() => throw new NotImplementedException();
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

    public static List<WorkQueueSubscription> Read(SqlConnection conn, List<WorkQueue> queues, List<User> users,
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
    public static void Write() => throw new NotImplementedException();
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

    public static List<User> Read(SqlConnection conn)
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

    public static void Write() => throw new NotImplementedException();
}

internal class Datalist(int datalistId, string name, string systemName)
{
    public int DatalistId { get; set; } = datalistId;
    public string Name { get; set; } = name;
    public string SystemName { get; set; } = systemName;

    public static List<Datalist> Read(SqlConnection conn)
    {
        Console.WriteLine("Started reading datalists.");
        const string sql = "select DataListID, Name, SystemName from DataList";
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
    public static void Write() => throw new NotImplementedException();
}

internal class ListRelationship(Datalist parent, Datalist child)
{
    public Datalist Parent { get; set; } = parent;
    public Datalist Child { get; set; } = child;

    public static List<ListRelationship> Read(SqlConnection conn, List<Datalist> dls)
    {
        Console.WriteLine("Started reading list relationships.");
        const string sql = "select ParentListID, ChildListID from ListRelationship";
        using var cmd = new SqlCommand(sql, conn);
        var rdr = cmd.ExecuteReader();
        var ret = new List<ListRelationship>();

        while (rdr.Read())
        {
            var p = dls.First(x => x.DatalistId == rdr.GetInt32(0));
            var c = dls.First(x => x.DatalistId == rdr.GetInt32(1));
            ret.Add(new ListRelationship(p, c));
        }

        rdr.Close();
        Console.WriteLine("Finished reading list relationships.");
        return ret;
    }

    public static List<ListRelationship> Transform(List<ListRelationship> lrs) => lrs;
    public static void Write() => throw new NotImplementedException();
}

internal class Permission
{
    private bool View { get; set; }
    private bool Add { get; set; }
    private bool Edit { get; set; }
    private bool BulkEdit { get; set; }
    private bool Delete { get; set; }
    private bool ViewActivity { get; set; }
    private bool Merge { get; set; }
    private bool Move { get; set; }
    private bool Administer { get; set; }

    public Permission()
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
    }

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

    public Permission AddNoAccess() => new();

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

    public static void Read() => throw new NotImplementedException();
    public static void Transform() => throw new NotImplementedException();
    public static void Write() => throw new NotImplementedException();
}

internal class ListRole
{
    public Datalist Datalist { get; set; }
    public SecurityRole SecurityRole { get; set; }
    public Permission Permission { get; set; }

    public ListRole(Datalist datalist, SecurityRole securityRole, Permission permission)
    {
        Datalist = datalist;
        SecurityRole = securityRole;
        Permission = permission;
    }

    public static void Read() => throw new NotImplementedException();
    public static void Transform() => throw new NotImplementedException();
    public static void Write() => throw new NotImplementedException();
}

internal class Export
{
    private string userSheetName = "Users-RO";
    private string userHierarchySheetName = "UserHierarchy-RO";
    private string dlSheetName = "Datalists-RO";
    private string dlHierarchySheetName = "DLHierarchy-RO";
    private string roleSheetName = "SecurityRoles";
    private string secMatrixSheetName = "SecurityMatrix";
    private string queueSheetName = "WorkQueues";
    private string workQueueMatrixSheetName = "WorkQueueMatrix";


    public void Run(String connectionString, string filename)
    {
        // Excel workbook
        var wb = new XLWorkbook();

        // Open a DB connection
        var conn = new SqlConnection(connectionString);
        conn.Open();

        // Get and read the data


        Console.WriteLine(connectionString);
        Console.WriteLine(filename);
    }
}