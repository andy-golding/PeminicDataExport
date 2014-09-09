using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeminicExportApp.Data;
using NPoco;

namespace PeminicExportApp.DataAccess
{
    class CaseFactory
    {
        private int uniqueIdField;

        private bool GenerateRandom { get; set; }

        private static CaseFactory caseFactoryInstance;

        private Case getCase(int caseid)
        {
            Case pcase = new Case();
            IDatabase db = DatabaseFactory.getDatabase();
            Sql sql = new Sql("select d.recid, d.nmseqnr, d.nmcaseid, f.nmparenttabid, d.nmfieldid, cast(d.stvalue as char) stvalue, f.nmdatatype from peminicdata d inner join peminicfields f on d.nmfieldid=f.recid where nmcaseid=@0", new[] { caseid });
            pcase.DataEntries = db.Fetch<PeminicExportApp.Data.Data>(sql);
            if (GenerateRandom)
            {
                pcase.DataEntries.ForEach(data => data.GenerateRandom = true);
            }

            if (uniqueIdField > 0)
            {
                PeminicExportApp.Data.Data uniquerec = pcase.DataEntries.Find(delegate(PeminicExportApp.Data.Data data)
                {
                    return data.FieldId == uniqueIdField;
                });
                if (uniquerec != null)
                {
                    pcase.UniqueID = uniquerec.Value;
                }
                else
                {
                    pcase.UniqueID = "Missing Unique ID (case " + caseid.ToString() + ")";                    
                }
            }
            return pcase;
        }

        private List<int> getCaseIdsForChoice(Choice choice)
        {            
            IDatabase db = DatabaseFactory.getDatabase();
            Sql sql = new Sql("select distinct d.nmcaseid from peminicdata d where d.nmfieldid=@0 and d.stvalue=@1", new object[] { choice.FieldID, '|'+choice.RecId.ToString()+'|' });
            var results = db.Fetch<int>(sql);
            return results;
        }

        private List<int> getCaseIds()
        {
            IDatabase db = DatabaseFactory.getDatabase();
            Sql sql = new Sql("select distinct d.nmcaseid from peminicdata d");
            var results = db.Fetch<int>(sql);
            return results;
        }

        private void setUniqueField(int fieldid)
        {
            uniqueIdField = fieldid;
        }

        public static CaseFactory Factory
        {
            get
            {
                if (caseFactoryInstance == null)
                {
                    caseFactoryInstance = new CaseFactory();
                }
                return caseFactoryInstance;
            }
        }

        public static void SetUniqueField(int fieldid)
        {
            Factory.setUniqueField(fieldid);
        }

        public static Case GetCase(int CaseId)
        {
            return Factory.getCase(CaseId);
        }

        public static List<int> GetCaseIDsForChoice(Choice choice)
        {
            return Factory.getCaseIdsForChoice(choice);
        }

        public static List<int> GetCaseIDs()
        {
            return Factory.getCaseIds();
        }

        public static void GenerateRadomData() {
            Factory.GenerateRandom = true;
        }

    }
}
