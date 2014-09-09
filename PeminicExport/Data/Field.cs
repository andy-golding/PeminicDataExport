using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPoco;

namespace PeminicExportApp.Data
{
    [TableName("peminicfields"),PrimaryKey("recid")]
    class Field
    {
        [Column("recid")]
        public int RecID { get; set; }

        [Column("nmTabId")]
        public int FieldSetId { get; set; }

        [Column("stFieldName")]
        public string Name { get; set; }

        [Column("nmParentTabId")]
        public int TabId { get; set; }
        
        [Column("nmOrder")]
        public int Order { get; set; }

        [Column("nmDatatype")]
        public int Type { get; set; }



        [Ignore]
        public Tab ParentTab { get; set; }

        
        public override string ToString()
        {
            return "Field "+Name + "("+RecID.ToString()+")";
        }
    }
}
