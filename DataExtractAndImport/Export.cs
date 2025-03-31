using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Data.SqlClient;

namespace DataExtractAndImport;

internal class SecurityRole
{
    public int SecurityRoleId { get; }
    public string Name { get; set; }

    public SecurityRole(int securityRoleId, string name)
    {
        SecurityRoleId = securityRoleId;
        Name = name;
    }

    public override bool Equals(object? obj) => Equals(obj as SecurityRole);

    private bool Equals(SecurityRole? obj) => obj != null && SecurityRoleId == obj.SecurityRoleId;

    public override int GetHashCode() => SecurityRoleId;

    /// <summary>
    /// Reads the security roles from the mCase database.
    /// </summary>
    /// <param name="conn">Connection to SQL database</param>
    /// <returns>List of roles from the database.</returns>
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

    public static List<SecurityRole> Transform(List<SecurityRole> roles)
    {
        roles.Sort((x, y) => String.Compare(x.Name, y.Name, StringComparison.OrdinalIgnoreCase));
        return roles;
    }

    public static bool Write()
    {
        throw new NotImplementedException();
    }
}

internal class WorkQueue
{
    public int WorkQueueId { get; set; }
    public string Name { get; set; }

    public WorkQueue(int workQueueId, string name)
    {
        WorkQueueId = workQueueId;
        Name = name;
    }
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
    public SecurityRoles? SecurityRole { get; set; }
    public SubscriptionType SubscriptionType { get; set; }

    public WorkQueueSubscription(WorkQueue workQueue, User? user, SubscriptionType subscriptionType)
    {
        WorkQueue = workQueue;
        User = user;
        SubscriptionType = subscriptionType;
    }

    public WorkQueueSubscription(WorkQueue workQueue, SecurityRoles? securityRole, SubscriptionType subscriptionType)
    {
        WorkQueue = workQueue;
        SecurityRole = securityRole;
        SubscriptionType = subscriptionType;
    }
}

internal class User
{
    public int UserId { get; set; }
    public string UserName { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Email { get; set; }
    public string EmailAddress { get; set; }
    public string Phone { get; set; }
    public bool ExternalAuthentication { get; set; }
    public int Active { get; set; }
    public int Supervisor { get; set; }

    public User(int userId, string userName, string firstName, string lastName, string email, string emailAddress,
        string phone, bool externalAuthentication, int active, int supervisor)
    {
        UserId = userId;
        UserName = userName;
        FirstName = firstName;
        LastName = lastName;
        Email = email;
        EmailAddress = emailAddress;
        Phone = phone;
        ExternalAuthentication = externalAuthentication;
        Active = active;
        Supervisor = supervisor;
    }
}

internal class Datalist
{
    public int DatalistId { get; set; }
    public string Name { get; set; }
    public string SystemName { get; set; }

    public Datalist(int datalistId, string name, string systemName)
    {
        DatalistId = datalistId;
        Name = name;
        SystemName = systemName;
    }
}

internal class ListRelationship
{
    public Datalist Parent { get; set; }
    public Datalist Child { get; set; }

    public ListRelationship(Datalist parent, Datalist child)
    {
        Parent = parent;
        Child = child;
    }
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
}

internal class ListRole
{
    public Datalist Datalist { get; set; }
    public SecurityRoles SecurityRoles { get; set; }
    public Permission Permission { get; set; }

    public ListRole(Datalist datalist, SecurityRoles securityRoles, Permission permission)
    {
        Datalist = datalist;
        SecurityRoles = securityRoles;
        Permission = permission;
    }
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


    public void Run(String ConnectionString, string Filename)
    {
        // Excel workbook
        var wb = new XLWorkbook();

        // Open a DB connection
        var conn = new SqlConnection(ConnectionString);
        conn.Open();


        Console.WriteLine(ConnectionString);
        Console.WriteLine(Filename);
    }
}