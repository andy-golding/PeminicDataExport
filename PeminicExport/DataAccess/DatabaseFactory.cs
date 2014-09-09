using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;
namespace PeminicExportApp.DataAccess
{
    class DatabaseFactory
    {
        //Class properties
        public string DBName { get; set; }
        
        //Factory bits
        private static DatabaseFactory factoryInstance;

        public static DatabaseFactory Factory
        {
            get
            {
                if (factoryInstance == null)
                {
                    factoryInstance = new DatabaseFactory();
                }
                return factoryInstance;
            }
        }

        public static void SelectDB(string dbname) {
            Factory.DBName = dbname;
        }

        public static IDatabase getDatabase() {
            return Database.getDatabase(Factory.DBName);
        }
    }
}
