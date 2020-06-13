using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Campus.DocumentValidator;
using Campus.Import;
using MOD_siblings_ImportExport.UDT;

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
            // 學生ID List
            List<int> studIdList = new List<int>();
            foreach (IRowStream ir in Rows)
            {
                
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {
                    string key = ir.GetValue("學號") + "_" + ir.GetValue("狀態");
                   if (Global._AllStudentNumberStatusIDTemp.ContainsKey(key))
                        studIdList.Add(Global._AllStudentNumberStatusIDTemp[key]);
                }
            }

            int TotalCount = 0, NewIdx = 0;
            foreach (IRowStream ir in Rows)
            {
                TotalCount++;
                this.ImportProgress = TotalCount;

                int sid = 0;
                if (ir.Contains("學號") && ir.Contains("狀態"))
                {


                }
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
