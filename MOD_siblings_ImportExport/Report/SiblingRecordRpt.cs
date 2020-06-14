using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using Aspose.Words;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace MOD_siblings_ImportExport.Report
{
    public class SiblingRecordRpt
    {

        // 所選學生系統編號
        List<string> StudentIDList;

        public SiblingRecordRpt(List<string> sidList)
        {
            StudentIDList = sidList;
        }


        /// <summary>
        /// 產生報表
        /// </summary>
        public void Report()
        {
            try
            {
                if (StudentIDList != null && StudentIDList.Count > 0)
                {
                    // 使用 SQL Query讀取 學生兄弟姊妹資料
                    QueryHelper queryHelper = new QueryHelper();

                    // --班級   座號 姓名  兄弟姊妹姓名 稱謂  生日 學校名稱    班級名稱 備註
                    //SELECT class.class_name AS 班級,student.seat_no AS 座號,student.name AS 姓名,$kcbs.udt.siblings.sibling_name AS 兄弟姊妹姓名,$kcbs.udt.siblings.sibling_title AS 稱謂,$kcbs.udt.siblings.birthday AS 生日,$kcbs.udt.siblings.school_name AS 學校名稱,$kcbs.udt.siblings.class_name AS 班級名稱,$kcbs.udt.siblings.remark AS 備註 FROM student INNER JOIN class ON student.ref_class_id = class.id INNER JOIN $kcbs.udt.siblings ON student.id = $kcbs.udt.siblings.ref_student_id WHERE student.id IN(4013) ORDER BY class.grade_year,class.display_order,class.class_name,student.seat_no

                    string sql = "SELECT class.class_name AS 班級,student.seat_no AS 座號,student.name AS 姓名,$kcbs.udt.siblings.sibling_name AS 兄弟姊妹姓名,$kcbs.udt.siblings.sibling_title AS 稱謂,$kcbs.udt.siblings.birthday AS 生日,$kcbs.udt.siblings.school_name AS 學校名稱,$kcbs.udt.siblings.class_name AS 班級名稱,$kcbs.udt.siblings.remark AS 備註 FROM student INNER JOIN class ON student.ref_class_id = class.id INNER JOIN $kcbs.udt.siblings ON student.id = $kcbs.udt.siblings.ref_student_id WHERE student.id IN(" + string.Join(",", StudentIDList.ToArray()) + ") ORDER BY class.grade_year,class.display_order,class.class_name,student.seat_no";
                    DataTable dt = queryHelper.Select(sql);

                    // 讀取 Word 樣板
                    Document doc = new Document(new MemoryStream(Properties.Resources.兄弟姊妹資訊樣板));

                    // 使用 DatatTable 合併 Word 
                    DataTable dtTable = new DataTable();

                  

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 對照合併欄位
                        for (int i = 1; i <= dt.Rows.Count; i++)
                        {
                            dtTable.Columns.Add("班級" + i + "");
                            dtTable.Columns.Add("座號" + i + "");
                            dtTable.Columns.Add("姓名" + i + "");
                            dtTable.Columns.Add("兄弟姊妹姓名" + i + "");
                            dtTable.Columns.Add("稱謂" + i + "");
                            dtTable.Columns.Add("生日" + i + "");
                            dtTable.Columns.Add("學校名稱" + i + "");
                            dtTable.Columns.Add("班級名稱" + i + "");
                            dtTable.Columns.Add("備註" + i + "");
                        }

                        // 全部在一頁合併
                        DataRow newRow = dtTable.NewRow();
                        int idx = 1;
                        // 填入資料
                        foreach (DataRow dr in dt.Rows)
                        {
                            newRow["班級" + idx] = "" + dr["班級"];
                            newRow["座號" + idx] = "" + dr["座號"];
                            newRow["姓名" + idx] = "" + dr["姓名"];
                            newRow["兄弟姊妹姓名" + idx] = "" + dr["兄弟姊妹姓名"];
                            newRow["稱謂" + idx] = "" + dr["稱謂"];

                            DateTime dtd;
                            if (DateTime.TryParse(dr["生日"]+"",out dtd))
                            {
                                newRow["生日" + idx] = dtd.ToShortDateString();
                            }
                            
                            newRow["學校名稱" + idx] = "" + dr["學校名稱"];
                            newRow["班級名稱" + idx] = "" + dr["班級名稱"];
                            newRow["備註" + idx] = "" + dr["備註"];
                            idx++;
                        }

                        // 加入 DatatTable,一頁一筆 row
                        dtTable.Rows.Add(newRow);

                        doc.MailMerge.Execute(dtTable);

                        // 移除沒有合併到合併欄位
                        doc.MailMerge.RemoveEmptyParagraphs = true;
                        doc.MailMerge.DeleteFields();

                        // 存成Word 
                        doc.Save(Application.StartupPath + "\\Reports\\兄弟姊妹報表範例.doc", SaveFormat.Doc);
                        System.Diagnostics.Process.Start(Application.StartupPath + "\\Reports\\兄弟姊妹報表範例.doc");
                    }                  
                   

                }

            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生報表失敗," + ex.Message);
            }
        }


        public void temp()
        {
            // 產生合併欄位

            Aspose.Words.Document doc = new Document();
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln();
        
            List<string> m1 = new List<string>();
            m1.Add("班級");
            m1.Add("座號");
            m1.Add("姓名");
            m1.Add("兄弟姊妹姓名");
            m1.Add("稱謂");
            m1.Add("生日");
            m1.Add("學校名稱");
            m1.Add("班級名稱");
            m1.Add("備註");


            builder.StartTable();

            for (int i = 1; i <= 12; i++)
            {
                foreach (string name in m1)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + name + i + " \\* MERGEFORMAT ", "«" + name + i + "»");
                }
                builder.EndRow();
            }
            builder.EndTable();

            doc.Save(Application.StartupPath + "\\test.doc", SaveFormat.Doc);
            System.Diagnostics.Process.Start(Application.StartupPath + "\\test.doc");
        }
    }
}
