using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;

namespace PeminicExportApp.DataAccess
{
    class Database : NPoco.Database
    {
        public string SelectedDB {get;set;}
        
        public Database(string connStringName, string dbname) : base(connStringName) {
            this.Execute("use `"+dbname+"`;");
            this.SelectedDB = dbname;
        }

        public static IDatabase getDatabase(string dbname) {
            IDatabase db = new Database("MySQLPeminic",dbname);
            return db;
        }
    }
}
