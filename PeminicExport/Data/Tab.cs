using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;

namespace PeminicExportApp.Data
{
    [TableName("peminictabs"),PrimaryKey("nmTabId")]
    class Tab
    {
        [Column("nmTabId")]
        public int TabId { get; set; }
        [Column("stTabName")]
        public string Name { get; set; }
        [Column("stDbTabName")]
        public string DatabaseName { get; set; }
        [Column("nmOrder")]
        public int Order { get; set; }


        public List<Field> Fields { get; set; }
    }
}
