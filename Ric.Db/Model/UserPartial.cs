using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ric.Db.Model
{
    partial class User
    {
        public override string ToString()
        {
            return String.Format("{0} {1}", Familyname, Surname);
        }
    }
}
