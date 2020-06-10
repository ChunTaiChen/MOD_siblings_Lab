using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SmartSchool.API.PlugIn;
using SmartSchool.API.PlugIn.Export;
using FISCA.UDT;
using MOD_siblings_ImportExport.UDT;

namespace MOD_siblings_ImportExport.ImportExport
{
    /// <summary>
    /// 匯出兄弟姊妹資訊
    /// </summary>
    class ExportSiblingRecord : SmartSchool.API.PlugIn.Export.Exporter
    {
        // 可勾選選項
        List<string> ExportItemList;

        public ExportSiblingRecord()
        {
            this.Image = null;
            this.Text = "匯出兄弟姊妹資訊";
            ExportItemList = new List<string>();

            // 需要匯出欄位
            ExportItemList.Add("兄弟姊妹姓名");
            ExportItemList.Add("稱謂");
            ExportItemList.Add("生日");
            ExportItemList.Add("學校名稱");
            ExportItemList.Add("班級名稱");
            ExportItemList.Add("備註");

        }

        // 實作匯出
        public override void InitializeExport(ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange(ExportItemList);
            wizard.ExportPackage += delegate (object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
             {
                 // 透過學生編號取得UDT兄弟姊妹資訊
                 AccessHelper accessHelper = new AccessHelper();
                 string qry = "ref_student_id IN(" + string.Join(",", e.List.ToArray()) + ")";
                 List<SiblingRecord> SiblingRecordList = accessHelper.Select<SiblingRecord>(qry);

                 // 填入資料               
                 foreach (SiblingRecord sr in SiblingRecordList)
                 {
                     // 填入資料
                     RowData row = new RowData();
                     row.ID = sr.StudnetID.ToString();

                     foreach (string field in e.ExportFields)
                     {
                         // 檢查需要匯出欄位
                         if (wizard.ExportableFields.Contains(field))
                         {
                             switch (field)
                             {
                                 case "兄弟姊妹姓名":
                                     row.Add(field, sr.SiblingName);
                                     break;
                                 case "稱謂":
                                     row.Add(field, sr.SiblingTitle);
                                     break;
                                 case "生日":
                                     row.Add(field, sr.Birthday.ToShortDateString());
                                     break;
                                 case "學校名稱":
                                     row.Add(field, sr.SchoolName);
                                     break;
                                 case "班級名稱":
                                     row.Add(field, sr.ClassName);
                                     break;
                                 case "備註":
                                     row.Add(field, sr.Remark);
                                     break;
                             }
                         }
                     }
                     e.Items.Add(row);
                 }
             };

        }
    }
}
