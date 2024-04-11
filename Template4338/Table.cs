using System.Data.Entity;

namespace Template4338
{
    public class MyDbContext : DbContext
    {
        public DbSet<Table> Tables { get; set; }
        public DbSet<TableJSON> TablesJSON { get; set; }

    }

    public class Table
    {
        public int Id { get; set; }
        public string ExcelId { get; set; }
        public string CodeOrder { get; set; }
        public string CreateDate { get; set; }
        public string CreateTime { get; set; }
        public string CodeClient { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }
        public string ClosedDate { get; set; }
        public string ProkatTime { get; set; }
    }

    public class TableJSON
    {
        public int Id { get; set; }
        public string CodeOrder { get; set; }
        public string CreateDate { get; set; }
        public string CreateTime { get; set; }
        public string CodeClient { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }
        public string ClosedDate { get; set; }
        public string ProkatTime { get; set; }      

        public TableJSON()
        {

        }
        public TableJSON(int id, string codeorder, string createdate, string createtime, string codeclient, string services, string status, string closeddate, string prokattime)
        {
            Id = id;
            CodeOrder = codeorder;
            CreateDate = createdate;
            CreateTime = createtime;
            CodeClient = codeclient;
            Services = services;           
            Status = status;
            ClosedDate = closeddate;
            ProkatTime = prokattime;

        }
    }
    public class ProkatTimeTable
    {
        public string ProkatTime { get; set; }
    }
}
