using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation.Controls;
using K12.Data;
using System.Windows.Forms;
using Campus.DocumentValidator;

namespace MOD_siblings_ImportExport
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            ////第一個 RibbonBar 功能
            //FISCA.Presentation.RibbonBarItem item = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "康橋"];
            //item["第一個程式"].Enable = FISCA.Permission.UserAcl.Current["09613d9c-56b6-4511-b046-5b40689d5955"].Executable;
            //item["第一個程式"].Click += delegate
            //{
            //    MessageBox.Show("Hello FISCA!");
            //};

            FISCA.Presentation.RibbonBarItem item_e = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "康橋"];
            item_e["匯出兄弟姊妹資訊"].Enable = FISCA.Permission.UserAcl.Current["20DD29B1-20E8-42C4-ACB6-3958FEF8A8C1"].Executable;
            item_e["匯出兄弟姊妹資訊"].Click += delegate
            {
                // 最少選一位學生
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    // 使用匯出 API
                    SmartSchool.API.PlugIn.Export.Exporter exporter = new MOD_siblings_ImportExport.ImportExport.ExportSiblingRecord();

                    // 使用匯出精靈
                    ImportExport.ExportStudentV2 wizard = new ImportExport.ExportStudentV2(exporter.Text, exporter.Image);
                    exporter.InitializeExport(wizard);
                    wizard.ShowDialog();
                }
            };

            FISCA.Presentation.RibbonBarItem item_i = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "康橋"];
            item_i["匯入兄弟姊妹資訊"].Enable = FISCA.Permission.UserAcl.Current["800D8794-BB03-4497-A53B-E8BED38DFB2B"].Executable;
            item_i["匯入兄弟姊妹資訊"].Click += delegate
            {
                // 載入所有學生與狀態，資料匯入比對使用
                Global._AllStudentNumberStatusIDTemp = Global.GetAllStudenNumberStatusDict();
                ImportExport.ImportSiblingRecord importSiblingRecord = new ImportExport.ImportSiblingRecord();
                importSiblingRecord.Execute();
            };

            // 載入自訂驗證規則
            #region 自訂驗證規則
            FactoryProvider.RowFactory.Add(new ValidationRule.SiblingRowValidatorFactory());
            #endregion

            // 報表
            FISCA.Presentation.RibbonBarItem item = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "康橋"];
            item["兄弟姊妹資訊報表範例"].Enable = FISCA.Permission.UserAcl.Current["CBFDCD41-49D8-4972-A7B9-7CFD93FCB9C4"].Executable;
            item["兄弟姊妹資訊報表範例"].Click += delegate
            {

                // 最少選一位學生
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    Report.SiblingRecordRpt siblingRecordRpt = new Report.SiblingRecordRpt(K12.Presentation.NLDPanels.Student.SelectedSource);
                    siblingRecordRpt.Report();

                }
            };




            Catalog catalog_e = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog_e.Add(new RibbonFeature("20DD29B1-20E8-42C4-ACB6-3958FEF8A8C1", "匯出兄弟姊妹資訊"));
            catalog_e.Add(new RibbonFeature("800D8794-BB03-4497-A53B-E8BED38DFB2B", "匯入兄弟姊妹資訊"));
            catalog_e.Add(new RibbonFeature("CBFDCD41-49D8-4972-A7B9-7CFD93FCB9C4", "兄弟姊妹資訊報表範例"));

        }
    }

}
