using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeminicExportApp.Data;
using NPoco;

namespace PeminicExportApp.DataAccess
{
    /**
     * SurveyStructure.cs
     * Defines the layout of a particular survey based on a choice value passed in, used to select fields and tabs
     */
    
    class SurveyStructure
    {

        public static Survey LoadSurveyStructure()
        {
            Survey surveyStruct = new Survey();                                   
            IDatabase db = DatabaseFactory.getDatabase();
            Sql sql = new Sql("select distinct pt.nmTabId, pt.stTabName, pt.stDbTabName, pt.nmOrder, f.recid, f.nmtabid, f.stfieldname, f.nmparenttabid, f.nmorder, f.nmdatatype " +
                                    "from peminicfields f " +
                                    "inner join " +
                                    "peminictabs pt on pt.nmTabId = f.nmParentTabId " +
                                    "inner join ( " +
                                    "select distinct d.nmfieldid from peminicdata ) foo on f.recid=foo.nmfieldid" +
                                    "order by pt.nmorder, f.nmorder");
            var fields = db.FetchOneToMany<Tab, Field>(x => x.TabId, sql);
            surveyStruct.Tabs = db.FetchOneToMany<Tab, Field>(x => x.TabId, sql);
            return surveyStruct;
        }


        public static Survey LoadSurveyStructure(Choice discrimiatorChoiceValue)
        {
            Survey surveyStruct = new Survey();
            Field discriminatorField = Fields.getField(discrimiatorChoiceValue.FieldID);
            string ChoiceSearch = "";
            ChoiceSearch = "|" + discrimiatorChoiceValue.RecId.ToString() + "|";
            IDatabase db = DatabaseFactory.getDatabase();
            object[] sqlparams = {discrimiatorChoiceValue.FieldID, ChoiceSearch};
            
            Sql sql = new Sql("select distinct pt.nmTabId, pt.stTabName, pt.stDbTabName, pt.nmOrder, f.recid, f.nmtabid, f.stfieldname, f.nmparenttabid, f.nmorder, f.nmdatatype " +
                                    "from peminicfields f " +
                                    "inner join " +
                                    "peminictabs pt on pt.nmTabId = f.nmParentTabId " +
                                    "inner join ( "+
                                    "select distinct d.nmfieldid from peminicdata c inner join peminicdata d on c.nmcaseid=d.nmcaseid "+
                                    "where c.nmfieldid = @0 " +
                                    "and c.stvalue = @1 ) foo on f.recid=foo.nmfieldid " +
                                    "order by pt.nmorder, f.nmorder", sqlparams);
            var fields = db.FetchOneToMany<Tab, Field>(x => x.TabId, sql);
            surveyStruct.Tabs = db.FetchOneToMany<Tab, Field>(x => x.TabId, sql);
            return surveyStruct;
        }
    }
}
