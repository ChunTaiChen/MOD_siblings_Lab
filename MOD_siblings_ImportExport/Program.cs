using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation.Controls;
using K12.Data;
using System.Windows.Forms;

namespace MOD_siblings_ImportExport
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            //第一個 RibbonBar 功能
            FISCA.Presentation.RibbonBarItem item = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "康橋"];
            item["第一個程式"].Enable = FISCA.Permission.UserAcl.Current["09613d9c-56b6-4511-b046-5b40689d5955"].Executable;
            item["第一個程式"].Click += delegate
            {
                MessageBox.Show("Hello FISCA!");
            };

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

            Catalog catalog_e = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog_e.Add(new RibbonFeature("20DD29B1-20E8-42C4-ACB6-3958FEF8A8C1", "匯出兄弟姊妹資訊"));

        }
    }

}
