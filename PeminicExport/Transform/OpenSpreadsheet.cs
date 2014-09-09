using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PeminicExportApp.Data;
using PeminicExportApp.DataAccess;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PeminicExportApp.Transform
{
    class FieldMapEntry
    {
        public int TabId { get; set; }
        public int FieldId { get; set; }
        public uint ColNum {get;set;}
        public string ColRef { get; set; }

        public override string ToString()
        {
            return TabId.ToString() + " "+FieldId.ToString() + " "+ ColRef;
        }
    }

    class DataTypeDescription {
        public string Description { get; set; }
        public CellValues CellType { get; set; }
    }

    class DataDecorator {
        
        public const CellValues Default = CellValues.String;
        public static Dictionary<int, DataTypeDescription> TypeDescriptions = new Dictionary<int, DataTypeDescription>()
        {
            {PeminicExportApp.Data.Data.AutoCounter, new DataTypeDescription {Description="Auto-incremented Field",CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.ChoiceType, new DataTypeDescription {Description="Choice",CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Text,  new DataTypeDescription {Description="Text", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.LongText,  new DataTypeDescription {Description="Long Text", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Money,  new DataTypeDescription {Description="Money", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Date,  new DataTypeDescription {Description="Date Only", CellType=CellValues.Date}},
            {PeminicExportApp.Data.Data.DateTime,  new DataTypeDescription {Description="Date/Time", CellType=CellValues.Date}},
            {PeminicExportApp.Data.Data.DOB,  new DataTypeDescription {Description="Date Of Birth", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Email,  new DataTypeDescription {Description="Email", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.FieldsToComplete,  new DataTypeDescription {Description="Fields to Complete", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.IVRVoiceRecording,  new DataTypeDescription {Description="Voice Recording", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.LabelOnly,  new DataTypeDescription {Description="Label", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.LayerTotal,  new DataTypeDescription {Description="Layer Total", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.LongTextAppend,  new DataTypeDescription {Description="Long Text (Appended)", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.MailMerge,  new DataTypeDescription {Description="Mail Merge", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.MoneyCalc,  new DataTypeDescription {Description="Calculated Money Value", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.MoneyDBTotal,  new DataTypeDescription {Description="Calculated DB Money Value", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.MoneyLayerTotal,  new DataTypeDescription {Description="Calculated Money Layer Total", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Number,  new DataTypeDescription {Description="Number", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.PeopleInvolved,  new DataTypeDescription {Description="People Involed", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.ReferenceList,  new DataTypeDescription {Description="Reference List", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.ReportingFormula,  new DataTypeDescription {Description="Reporting Formula", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.SendEmail,  new DataTypeDescription {Description="Send Email", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Template,  new DataTypeDescription {Description="Template", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Time,  new DataTypeDescription {Description="Time", CellType=CellValues.Date}},
            {PeminicExportApp.Data.Data.Upload,  new DataTypeDescription {Description="Upload", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.Weekday,  new DataTypeDescription {Description="Weekday", CellType=DataDecorator.Default}},
            {PeminicExportApp.Data.Data.YesNo,  new DataTypeDescription {Description="Yes/No", CellType=DataDecorator.Default}}
        };

        public static Cell GetCellValue(PeminicExportApp.Data.Data dataEntry) {
            if (dataEntry == null) { return new Cell() { CellValue = new CellValue(string.Empty), DataType = new EnumValue<CellValues>(DataDecorator.Default) }; }

            string Value = string.Empty;
            EnumValue<CellValues> type = null;

            if(DataDecorator.TypeDescriptions.ContainsKey(dataEntry.Type)) {
                type=new EnumValue<CellValues>(TypeDescriptions[dataEntry.Type].CellType);
            } else {
                type=new EnumValue<CellValues> (DataDecorator.Default);
            }

            if(dataEntry.Type==PeminicExportApp.Data.Data.ChoiceType) {
                List<Choice> choices =Choices.GetChoicesFromString(dataEntry.Value);
                foreach (Choice ch in choices) {
                    Value +=", "+ch.ChoiceValue;

                }
                if (Value != string.Empty)
                {
                    Value = Value.Substring(1);
                }
            }
            else
            {
                Value = dataEntry.Value;
            }
            return new Cell() { CellValue = new CellValue(Value), DataType = DataDecorator.Default};            
        }
    }

    class DataRow {
        private Dictionary<FieldMapEntry,PeminicExportApp.Data.Data> rowData = new Dictionary<FieldMapEntry,PeminicExportApp.Data.Data>();

        public void Add(FieldMapEntry fieldmap, PeminicExportApp.Data.Data value){
            rowData.Add(fieldmap,value);
        }

        public void Update(FieldMapEntry fieldmap, PeminicExportApp.Data.Data value)
        {
            rowData[fieldmap] = value;
        }

        public void WriteRow(TabsheetData tabsheetdata, string uniqueid, int layerid, uint rowidx)
        {
            Row row = new Row(){RowIndex=rowidx};
            Cell uniquecell = new Cell() { CellReference = "A" + rowidx.ToString(), CellValue = new CellValue(uniqueid), DataType = new EnumValue<CellValues>(DataDecorator.Default) };
            tabsheetdata.UpdateMaxWidth(0, uniqueid);
            Cell layercell = new Cell() { CellReference = "B" + rowidx.ToString(), CellValue = new CellValue(layerid.ToString()), DataType = new EnumValue<CellValues>(DataDecorator.Default) };
            tabsheetdata.UpdateMaxWidth(1, layerid.ToString());
            row.Append(uniquecell);
            row.Append(layercell);
            List<FieldMapEntry> cellrefs = rowData.Keys.ToList<FieldMapEntry>();
            cellrefs.Sort((ent1, ent2) => ent1.ColNum.CompareTo(ent2.ColNum));
            uint idx=2;
            foreach (var fieldMapEntry in cellrefs)
            {
                Cell cell = DataDecorator.GetCellValue(rowData[fieldMapEntry]);
                tabsheetdata.UpdateMaxWidth(idx,cell.CellValue.Text);
                cell.CellReference = fieldMapEntry.ColRef + rowidx.ToString();
                row.Append(cell);
                idx++;
            }
            SheetData data = tabsheetdata.WorksheetPart.Worksheet.GetFirstChild<SheetData>();
            data.Append(row);
        }
    }

    class DataRowSet
    {
        private Dictionary<int, DataRow> rowSet = new Dictionary<int, DataRow>();
        private string uniqueId;
        private List<FieldMapEntry> rowtemplate;

        public DataRowSet(string uid, List<FieldMapEntry> template){
            uniqueId=uid;
            this.rowtemplate = template;
        }

        private DataRow EnsureRow(int layerid){
            if (!rowSet.ContainsKey(layerid))
            {
                rowSet.Add(layerid,new DataRow());
                foreach (FieldMapEntry entry in rowtemplate)
                {
                    rowSet[layerid].Add(entry, null);
                }

            }
            return rowSet[layerid];
        }

        public void Add(FieldMapEntry mapEntry, PeminicExportApp.Data.Data dataEntry)
        {
            this.EnsureRow(dataEntry.LayerId);
            rowSet[dataEntry.LayerId].Update(mapEntry, dataEntry);
        }

        public void WriteRowSet(TabsheetData data, uint lastRowIndex)
        {
            uint idx = lastRowIndex;
            foreach (KeyValuePair<int, DataRow> pair in rowSet)
            {
                pair.Value.WriteRow(data, uniqueId, pair.Key, idx);
                idx++;
            }
            //part.Worksheet.Save();
        }
    }

    class TabsheetData
    {
        Dictionary<uint,int> columnwidths = new Dictionary<uint,int>();
        public WorksheetPart WorksheetPart { get; set; }
        const int MaxWidth = 100;
        const int MinWidth = 10;
        public void UpdateMaxWidth(uint colnum, string value)
        {
            if (!columnwidths.ContainsKey(colnum)) {
                columnwidths.Add(colnum,MinWidth);
            }
            if (columnwidths[colnum] < value.Length)
            {
                columnwidths[colnum] = value.Length > MaxWidth ? MaxWidth : value.Length;
            }
        }

        public void WriteWidths()
        {
            Columns cols = new Columns();
            List<uint> keys = columnwidths.Keys.ToList<uint>();
            keys.Sort();
            foreach (uint colnum in keys) 
            {
                cols.Append(new Column() { Max = colnum+1, Min = colnum+1, BestFit=true, CustomWidth = true, Width = 1.5 * (double)columnwidths[colnum] });
            }
            this.WorksheetPart.Worksheet.InsertBefore<Columns>(cols, this.WorksheetPart.Worksheet.FirstChild);
        }
    }

    class OpenSpreadsheet
    {
        private Survey template;
        private SpreadsheetDocument document;
        private Dictionary<int, TabsheetData> tabsheets = new Dictionary<int, TabsheetData>();
        private Dictionary<int, FieldMapEntry> fieldmap = new Dictionary<int, FieldMapEntry>();
        private uint SheetId = 1;

        private SharedStringTablePart sharedStringTablePart
        {
            get
            {
                SharedStringTablePart shareStringPart;
                if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    shareStringPart.SharedStringTable = new SharedStringTable();
                }
                return shareStringPart;
            }
        }

        private int AddSharedString(string value)
        {
            sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
            sharedStringTablePart.SharedStringTable.Save();
            return sharedStringTablePart.SharedStringTable.ChildElements.Count - 1;
        }
        

        private void AddField(Row row, PeminicExportApp.Data.Field field, uint idx)
        {
            string colletter = this.ColNumToLetter(idx);
            string cellref = colletter + "1";
            Cell cell;

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellref).Count() > 0)
            {
                cell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellref).First();
            }
            else
            {
                cell = new Cell() { CellReference = cellref };
                row.Append(cell);
            }
            //CellValue cv = new CellValue(AddSharedString(field.Name).ToString());
            CellValue cv = new CellValue(field.Name);
            cell.CellValue = cv;
            cell.DataType = new EnumValue<CellValues>(DataDecorator.Default);
            fieldmap.Add(field.RecID, new FieldMapEntry(){ TabId = field.TabId, FieldId = field.RecID, ColNum = idx, ColRef = colletter});
        }

        private void writeHeader(TabsheetData tabsheetdata, Tab tab)
        {
            WorksheetPart worksheetPart = tabsheetdata.WorksheetPart;
            SheetData sheet = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            Row row;
            if (sheet.Elements<Row>().Where( r => r.RowIndex == 1).Count() > 0)
            {
                row =  sheet.Elements<Row>().Where( r => r.RowIndex == 1).First();
            }
            else
            {
                row = new Row() { RowIndex = 1 };                
                Cell cell = new Cell() { CellReference = "A1" };               
                CellValue cv = new CellValue("Unique ID");
                cell.CellValue = cv;
                cell.DataType = new EnumValue<CellValues>(DataDecorator.Default);
                row.Append(cell);

                cell = new Cell() { CellReference = "B1" };
                cv = new CellValue("Layer ID");
                cell.CellValue = cv;
                cell.DataType = new EnumValue<CellValues>(DataDecorator.Default);
                row.Append(cell);

                sheet.Append(row);
            }

            uint idx = 2;
            foreach (PeminicExportApp.Data.Field field in tab.Fields)
            {
                AddField(row, field, idx);
                tabsheetdata.UpdateMaxWidth(idx, field.Name);
                idx++;
            }
            worksheetPart.Worksheet.Save();
        }

        private int ColLetterToNum(string str)
        {
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int outs = 0;
            int factor = 0;
            foreach (char let in str.Reverse())
            {
                int pos = alphabet.IndexOf(let);
                outs += (pos + 1) * this.pow(26, factor);
                factor++;
            }
            return outs - 1;
        }

        private string ColNumToLetter(uint num)
        {
            string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string outs = "";
            int number = (int)num;
            if (number == 0) return alphabet[0].ToString();
            while (number >= 0)
            {
                int idx = number % 26;
                outs = alphabet[idx].ToString() + outs;
                number /= 26;
                number--;
            }
            return outs;
        }

        private int pow(int x, int y)
        {
            if (y == 0) return 1;
            if (y == 1) return x;
            int tmp = x;
            int n = 31;
            while ((y <<= 1) >= 0) n--;
            while (--n > 0)
                tmp = tmp * tmp *
                     (((y <<= 1) < 0) ? x : 1);
            return tmp;
        }

        private WorksheetPart AddSheet(WorkbookPart workbookPart, string sheetName)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.Save();
            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);


            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = this.SheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();
            this.SheetId++;
            return worksheetPart;
        }

        private void createSpreadSheet(string documentname) {
            document = SpreadsheetDocument.Create(documentname, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            workbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            foreach (Tab tab in template.Tabs)
            {
                WorksheetPart worksheetPart = this.AddSheet(workbookPart, tab.Name);
                TabsheetData tabsheetdata = new TabsheetData { WorksheetPart = worksheetPart };
                tabsheets.Add(tab.TabId, tabsheetdata );
                writeHeader(tabsheetdata, tab);
            }
        }

        private void writeDataDictionary()
        {
            WorksheetPart worksheetPart = this.AddSheet(document.WorkbookPart, "Data Dictionary");
            Columns cols = new Columns();
            worksheetPart.Worksheet.InsertBefore<Columns>(cols, worksheetPart.Worksheet.FirstChild);
            cols.Append(new Column { CustomWidth = true, BestFit = true, Width = 40, Min = 1, Max = 1 });
            cols.Append(new Column { CustomWidth = true, BestFit = true, Width = 64, Min = 2, Max = 2 });
            cols.Append(new Column { CustomWidth = true, BestFit = true, Width = 18, Min = 3, Max = 3 });
            cols.Append(new Column { CustomWidth = true, BestFit = true, Width = 64, Min = 4, Max = 4 });
            SheetData sheet = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            uint ridx = 2;
            Row row = new Row() { RowIndex = 1 };
            row.Append(new Cell { CellReference = "A1", DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue("Worksheet Name") });
            row.Append(new Cell { CellReference = "B1", DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue("Field Name") });
            row.Append(new Cell { CellReference = "C1", DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue("Worksheet Column") });
            row.Append(new Cell { CellReference = "D1", DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue("Field Data Type") });
            sheet.Append(row);
            foreach (var tab in template.Tabs)
            {                
                row = new Row() { RowIndex = ridx };
                row.Append(new Cell { CellReference = "A" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(tab.Name) });
                row.Append(new Cell { CellReference = "B" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(string.Empty) });
                row.Append(new Cell { CellReference = "C" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(string.Empty) });
                row.Append(new Cell { CellReference = "D" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(string.Empty) });
                ridx++;
                sheet.Append(row);
                foreach (var field in tab.Fields)
                {
                    row = new Row() { RowIndex = ridx };
                    DataTypeDescription typedesc = DataDecorator.TypeDescriptions[field.Type];
                    row.Append(new Cell { CellReference = "A" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(string.Empty) });
                    row.Append(new Cell { CellReference = "B" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(field.Name) });
                    row.Append(new Cell { CellReference = "C" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(fieldmap[field.RecID].ColRef) });
                    row.Append(new Cell { CellReference = "D" + ridx.ToString(), DataType = new EnumValue<CellValues>(CellValues.String), CellValue = new CellValue(typedesc.Description) });
                    ridx++;
                    sheet.Append(row);
                }
            }
        }

        public OpenSpreadsheet(Survey template, string DocumentName)
        {
            this.template = template;
            createSpreadSheet(DocumentName);
        }

        public void Close(){
            foreach (KeyValuePair<int, TabsheetData> pair in this.tabsheets)
            {
                pair.Value.WriteWidths();
                pair.Value.WorksheetPart.Worksheet.Save();
            }
            this.writeDataDictionary();
            this.document.WorkbookPart.Workbook.Save();
            this.document.Close();
        }

        public void AddCase(Case pcase)
        {
            Dictionary<int, DataRowSet> rowsets = new Dictionary<int, DataRowSet>();
            foreach (PeminicExportApp.Data.Data dataEntry in pcase.DataEntries)
            {
                if (!rowsets.ContainsKey(dataEntry.TabId))
                {
                    List<FieldMapEntry> template = fieldmap.Values.Where(kv => kv.TabId == dataEntry.TabId).ToList();
                    rowsets.Add(dataEntry.TabId, new DataRowSet(pcase.UniqueID, template));
                }                
                rowsets[dataEntry.TabId].Add(this.fieldmap[dataEntry.FieldId],dataEntry);
            }

            

            foreach (KeyValuePair<int, DataRowSet> pair in rowsets)
            {
                uint rowidx=(uint)this.tabsheets[pair.Key].WorksheetPart.Worksheet.GetFirstChild<SheetData>().Count()+1;
                pair.Value.WriteRowSet(this.tabsheets[pair.Key], rowidx);                
            }
        }

    }
}
