using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MOD_siblings_ImportExport.UDT
{
    [FISCA.UDT.TableName("kcbs.udt.siblings")]
    public class SiblingRecord : ActiveRecord
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = false)]
        public int StudnetID { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [Field(Field = "sibling_name", Indexed = false)]
        public string SiblingName { get; set; }

        /// <summary>
        /// 稱謂
        /// </summary>
        [Field(Field = "sibling_title", Indexed = false)]
        public string SiblingTitle { get; set; }

        /// <summary>
        /// 生日
        /// </summary>
        [Field(Field = "birthday", Indexed = false)]
        public DateTime Birthday { get; set; }

        /// <summary>
        /// 學校名稱
        /// </summary>
        [Field(Field = "school_name", Indexed = false)]
        public string SchoolName { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        [Field(Field = "class_name", Indexed = false)]
        public string ClassName { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        [Field(Field = "remark", Indexed = false)]
        public string Remark { get; set; }
    }
}
