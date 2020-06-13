using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Campus.DocumentValidator;

namespace MOD_siblings_ImportExport.ValidationRule
{
    public class SiblingRowValidatorFactory : IRowValidatorFactory
    {
        public IRowVaildator CreateRowValidator(string typeName, XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {                
                case "COUNSELSTUDCHECKSTUDENTNUMBERSTATUSVAL":
                    return new RowValidator.StudCheckStudentNumberStatusVal();
             
                default:
                    return null;
            }
        }
    }
}
