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

        }
    }

}
