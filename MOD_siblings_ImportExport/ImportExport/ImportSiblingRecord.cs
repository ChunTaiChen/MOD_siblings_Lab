using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;
using Campus.Import;
using MOD_siblings_ImportExport.UDT;
using FISCA.UDT;

namespace MOD_siblings_ImportExport.ImportExport
{
    public class ImportSiblingRecord : ImportWizard
    {
        // 設定檔
        private ImportOption _Option;

        // 新增
        private List<SiblingRecord> _InsertRecList;
        // 更新
        private List<SiblingRecord> _UpdateRecList;

        public override ImportAction GetSupportActions()
        {
            //新增或更新
            return ImportAction.InsertOrUpdate;
        }


        public override string GetValidateRule()
        {
            // 資料驗證規格 XML 
            return Properties.Resources.ImportSiblingRecordDataVal;
        }

        public override string Import(List<Campus.DocumentValidator.IRowStream> Rows)
        {
            // 實作匯入資料
            // 比對待匯入Excel檔案內資料轉換取得 學生ID List
            List<int> studIdList = new List<int>();
            foreach (IRowStream ir in Rows)
            {

                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    string key = ir.GetValue("學號") + "_" + ir.GetValue("狀態");
                    if (Global._AllStudentNumberStatusIDTemp.ContainsKey(key))
                        if (!studIdList.Contains(Global._AllStudentNumberStatusIDTemp[key]))
                            studIdList.Add(Global._AllStudentNumberStatusIDTemp[key]);
                }
            }


            AccessHelper accessHelper = new AccessHelper();
            string qry = "ref_student_id in(" + string.Join(",", studIdList.ToArray()) + ")";
            // 透過 StudentID 取得兄弟姊妹資料
            List<SiblingRecord> SiblingRecordList = accessHelper.Select<SiblingRecord>(qry);

            // 建立已有資料修改索引
            Dictionary<string, SiblingRecord> SiblingRecordDict = new Dictionary<string, SiblingRecord>();
            foreach (SiblingRecord rec in SiblingRecordList)
            {
                // key: StudentID+兄弟姊妹姓名+稱謂
                string key = rec.StudnetID + "_" + rec.SiblingName + "_" + rec.SiblingTitle;
                if (!SiblingRecordDict.ContainsKey(key))
                    SiblingRecordDict.Add(key, rec);
            }

            int TotalCount = 0;

            // 需要新增資料
            List<SiblingRecord> InsertRecList = new List<SiblingRecord>();

            // 需要修改資料
            List<SiblingRecord> UpdateRecList = new List<SiblingRecord>();

            // 資料整理
            foreach (IRowStream ir in Rows)
            {
                TotalCount++;
                this.ImportProgress = TotalCount;
                if (ir.Contains("學號") && ir.Contains("狀態") && ir.Contains("兄弟姊妹姓名") && ir.Contains("稱謂"))
                {
                    // 判斷需要新增或修該
                    string key = ir.GetValue("學號") + "_" + ir.GetValue("狀態");
                    int sid = 0;

                    // 比對解析 StudentID 
                    if (Global._AllStudentNumberStatusIDTemp.ContainsKey(key))
                    {
                        sid = Global._AllStudentNumberStatusIDTemp[key];
                    }

                    string key1 = sid + "_" + ir.GetValue("兄弟姊妹姓名") + "_" + ir.GetValue("稱謂");

                    SiblingRecord siblingRecord = null;
                    bool isInsert = false;
                    if (!SiblingRecordDict.ContainsKey(key1))
                    {
                        siblingRecord = new SiblingRecord();
                        // 有比對到學生 ID
                        if (sid > 0)
                        {
                            isInsert = true;
                            // 填入預設值
                            siblingRecord.StudnetID = sid;
                            siblingRecord.SiblingName = ir.GetValue("兄弟姊妹姓名");
                            siblingRecord.SiblingTitle = ir.GetValue("稱謂");
                        }
                    }
                    else
                    {
                        // 更新資料
                        siblingRecord = SiblingRecordDict[key1];
                    }

                    // 判斷是否有值才更新
                    if (ir.Contains("生日"))
                    {
                        DateTime dt;
                        if (DateTime.TryParse(ir.GetValue("生日"), out dt))
                        {
                            siblingRecord.Birthday = dt;
                        }

                    }

                    if (ir.Contains("學校名稱"))
                    {
                        siblingRecord.SchoolName = ir.GetValue("學校名稱");
                    }

                    if (ir.Contains("班級名稱"))
                    {
                        siblingRecord.ClassName = ir.GetValue("班級名稱");
                    }

                    if (ir.Contains("備註"))
                    {
                        siblingRecord.Remark = ir.GetValue("備註");
                    }

                    if (isInsert)
                    {
                        InsertRecList.Add(siblingRecord);
                    }
                    else
                    {
                        UpdateRecList.Add(siblingRecord);
                    }
                }
            }

            // 資料回寫 UDT
            if(InsertRecList.Count > 0)
            {
                InsertRecList.SaveAll();
            }

            if (UpdateRecList.Count > 0)
            {
                UpdateRecList.SaveAll();
            }

            return "";
        }

        public override void Prepare(ImportOption Option)
        {
            // 畫面初始化後需要載入
            _Option = Option;
            _InsertRecList = new List<SiblingRecord>();
            _UpdateRecList = new List<SiblingRecord>();
        }
    }
}
