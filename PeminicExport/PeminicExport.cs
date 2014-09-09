using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PeminicExportApp.DataAccess;
using PeminicExportApp.Data;
using PeminicExportApp.Transform;

namespace PeminicExportApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string dbstr = ConfigurationManager.AppSettings["dbnames"];
            string[] dbs = dbstr.Split(',');
            string idmapstr = ConfigurationManager.AppSettings["idfieldmap"];
            Dictionary<string, int> idmap = new Dictionary<string, int>();
            foreach (string pairstr in idmapstr.Split(','))
            {
                string[] pair = pairstr.Split(':');
                idmap.Add(pair[0], int.Parse(pair[1]));
            }
            Dictionary<string, int> choicemap = new Dictionary<string, int>();
            foreach (string pairstr in ConfigurationManager.AppSettings["choicemap"].Split(','))
            {
                string[] pair = pairstr.Split(':');
                choicemap.Add(pair[0], int.Parse(pair[1]));
            }

            bool random = false;
            if (bool.TryParse(ConfigurationManager.AppSettings["randomtestdata"], out random) && random) 
            {
                CaseFactory.GenerateRadomData();
            }
            foreach (string dbname in dbs)
            {
                DatabaseFactory.SelectDB(dbname);
                if (choicemap.ContainsKey(dbname))
                {
                    foreach (Choice choice in Choices.GetChoices(choicemap[dbname]))
                    {
                        List<int> caseIds = CaseFactory.GetCaseIDsForChoice(choice);
                        if (caseIds.Count > 0)
                        {
                            Survey surveyStruct = SurveyStructure.LoadSurveyStructure(choice);
                            OpenSpreadsheet os = new OpenSpreadsheet(surveyStruct, @"C:\Users\Public\Documents\" + choice.ChoiceValue.Replace(@"\", "").Replace("/", "") + ".xlsx");
                            CaseFactory.SetUniqueField(idmap[dbname]);
                            caseIds.ForEach(id => os.AddCase(CaseFactory.GetCase(id)));
                            os.Close();
                        }
                    }
                }
                else
                {
                         Survey surveyStruct = SurveyStructure.LoadSurveyStructure();
                        OpenSpreadsheet os = new OpenSpreadsheet(surveyStruct, @"C:\Users\Public\Documents\" + dbname + ".xlsx");
                        CaseFactory.SetUniqueField(idmap[dbname]);
                        List<int> caseIds = CaseFactory.GetCaseIDs();
                        caseIds.ForEach(id => os.AddCase(CaseFactory.GetCase(id)));
                        os.Close();

                }
            }
        }
    }
}
