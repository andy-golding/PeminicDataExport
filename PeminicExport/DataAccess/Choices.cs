using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeminicExportApp.Data;
using NPoco;

namespace PeminicExportApp.DataAccess
{
    class Choices
    {
        private static Choices instance;
        private Dictionary<int, Choice> choicerepo = new Dictionary<int, Choice>();

        private static Choices Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Choices();
                }
                return instance;
            }
        }

        private Choice get(int RecId) {
            if (choicerepo.ContainsKey(RecId)) {
                return choicerepo[RecId];
            } else {
            IDatabase db = DatabaseFactory.getDatabase();
            Choice choice = db.SingleOrDefaultById<Choice>(RecId);
            choicerepo.Add(RecId,choice);
            return choice;
            }
        }

        public static List<Choice> GetChoicesFromString(string str) {
            List<Choice> choices = new List<Choice>();
            if (str=="-1") {
                return choices;
            }
            foreach(string item in str.Split('|')) {
                if (!item.Trim().Equals(string.Empty)){
                    int choiceid;
                    if (int.TryParse(item, out choiceid))
                    {
                        Choice ch = Instance.get(choiceid);
                        if (ch != null)
                        {
                            choices.Add(ch);
                        }
                    }
                }
            }
            return choices;
        }

        public static Choice getChoice(int RecId) {
            return Instance.get(RecId);
        }

        public static List<Choice> GetChoices(int fieldid)
        {
            IDatabase db = DatabaseFactory.getDatabase();
            return db.Fetch<Choice>("Select RecId, stChoiceValue, nmFieldId from peminicchoices where nmFieldId=@0", new[] { fieldid });
        }

    }
}
