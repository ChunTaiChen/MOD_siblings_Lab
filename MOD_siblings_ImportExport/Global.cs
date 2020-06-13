using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;

namespace MOD_siblings_ImportExport
{
    public class Global
    {
        /// <summary>
        /// 所有學生學號與狀態對應的暫存
        /// </summary>
        public static Dictionary<string, int> _AllStudentNumberStatusIDTemp = new Dictionary<string, int>();

        
        /// <summary>
        /// 取得所有學生學號_狀態對應 StudentNumber_Status,id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, int> GetAllStudenNumberStatusDict()
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string strSQL = "SELECT " +
                "student.student_number" +
                ",CASE WHEN student.status = 1 THEN '一般'" +
                " WHEN student.status = 2 THEN '延修'" +
                " WHEN student.status = 4 THEN '休學'" +
                " WHEN student.status = 8 THEN '輟學'" +
                " WHEN student.status = 16 THEN '畢業或離校'" +
                " WHEN student.status = 256 THEN '刪除' ELSE '' END" +
                ",student.id FROM student;";
            DataTable dt = qh.Select(strSQL);
            foreach (DataRow dr in dt.Rows)
            {
                string key = dr[0].ToString() + "_" + dr[1].ToString();
                int id = int.Parse(dr[2].ToString());
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, id);
            }

            return retVal;
        }
                   
    }
}
