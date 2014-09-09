using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;

namespace PeminicExportApp.Data
{
    [TableName("peminicchoices"),PrimaryKey("recid")]
    class Choice
    {
        [Column("recid")]
        public int RecId { get; set; }

        [Column("stChoiceValue")]
        public string ChoiceValue { get; set; }
        
        [Column("nmFieldId")]
        public int FieldID { get; set; }
    }
}
