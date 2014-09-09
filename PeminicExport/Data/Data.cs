using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;

namespace PeminicExportApp.Data
{
    [TableName("peminicdata"),PrimaryKey("RecId")]
    class Data
    {
        /*
         * <option value='3'>Short Text</option>
         * <option value='4'>Long Text</option>
         * <option value='2'>Yes/No</option><
         * option value='5'>Money</option>
         * <option value='28'>Money - LAYER total for reference field</option>
         * <option value='29'>Money - DATABASE total for reference field</option>
         * <option value='30'>Money - CALCULATION FORMULA</option>
         * <option value='21'>Time (only)</option>
         * <option value='8'>Date & Time</option>
         * <option value='20'>Date (only)</option>
         * <option value='22'>DOB & AGE (calculation)</option>
         * <option value='25'>Week Day</option>
         * <option value='9' selected>Choice</option>
         * <option value='26'>Reference List (from multi layers)</option>
         * <option value='24'>Auto Counter</option>
         * <option value='1'>Integer (whole numbers)</option>
         * <option value='31'>NUMBER</option>
         * <option value='12'>People Involved</option>
         * <option value='13'>Upload</option>
         * <option value='14'>Template</option>
         * <option value='15'>EMail</option>
         * <option value='16'>Fields to complete</option>
         * <option value='17'>Mail Merge</option>
         * <option value='18'>Send E-Mail</option>
         * <option value='19'>IVR Voice Recording</option>
         * <option value='27'>No data label only</option>
         * <option value='32'>Long Text / Append only</option>
         * <option value='33'>Reporting Formula Field</option>
         * <option value='34'>LAYER total for reference field</option>
         */

        private static string LoremIpsum(int numchars)
        {

            var words = new[]{"lorem", "ipsum", "dolor", "sit", "amet", "consectetuer",
        "adipiscing", "elit", "sed", "diam", "nonummy", "nibh", "euismod",
        "tincidunt", "ut", "laoreet", "dolore", "magna", "aliquam", "erat"};

            var rand = new Random();

            string result = string.Empty;

            while (result.Length < numchars) {                         
                        result += " "+words[rand.Next(words.Length)];
             }
             result += ".";

            return result.Substring(1);
        }

        public const int ChoiceType = 9;
        public const int Text = 3;
        public const int LongText = 4;
        public const int YesNo = 2;
        public const int Money = 5;
        public const int MoneyLayerTotal = 28;
        public const int MoneyDBTotal = 29;
        public const int MoneyCalc = 30;
        public const int Time = 21;
        public const int DateTime = 8;
        public const int Date = 20;
        public const int DOB = 22;
        public const int Weekday = 25;
        public const int ReferenceList = 26;
        public const int AutoCounter = 24;
        public const int Number = 31;
        public const int PeopleInvolved = 12;
        public const int Upload = 13;
        public const int Template = 14;
        public const int Email = 15;
        public const int FieldsToComplete = 16;
        public const int MailMerge = 17;
        public const int SendEmail = 18;
        public const int IVRVoiceRecording = 19;
        public const int LabelOnly = 27;
        public const int LongTextAppend = 32;
        public const int ReportingFormula = 33;
        public const int LayerTotal = 34;

        [Column("RecId")]
        public int RecId { get; set; }

        [Column("nmseqnr")]
        public int LayerId { get; set; }

        [Column("nmCaseId")]
        public int CaseId { get; set; }

        [Column("nmParentTabId")]
        public int TabId { get; set; }

        [Column("nmfieldid")]
        public int FieldId { get; set; }

        private string stValue;
        [Column("stValue")]
        public string Value { 
            get 
            {
                if (this.GenerateRandom && this.Type != ChoiceType)
                {
                    return Data.LoremIpsum(this.stValue.Length);
                }
                else
                {
                    return this.stValue;
                }
            }
             set  { stValue = value;}
        }

        [Column("nmdatatype")]
        public int Type { get; set; }

        [Ignore]
        public Choice ChoiceValue { get; set; }

        [Ignore]
        public int[] ChoiceIds
        {
            get{
                if(Type!=9) {
                    string[] parts = Value.Split('|');
                    int[] choiceids = new int[parts.Length];
                    int idx=0;
                    foreach (string part in parts)
                    {
                        int ipart;
                        int.TryParse(part,out ipart);
                        choiceids[idx]=ipart;
                        idx++;
                    }
                    return choiceids;
                }
                else {
                    return null;
                }
            }
        }
        [Ignore]
        public bool GenerateRandom { get; set; }
    }
}
