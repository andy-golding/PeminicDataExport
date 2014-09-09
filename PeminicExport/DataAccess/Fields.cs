using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeminicExportApp.Data;
using NPoco;

namespace PeminicExportApp.DataAccess
{
    class Fields
    {
        public static Field getField(int recid)
        {
            IDatabase db = DatabaseFactory.getDatabase();
            return db.SingleById<Field>(recid);
        }

    }
}
