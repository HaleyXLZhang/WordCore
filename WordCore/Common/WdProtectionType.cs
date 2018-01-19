using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordCore.Common
{
    public enum WdProtectionType
    {
        wdNoProtection = -1,
        wdAllowOnlyRevisions = 0,
        wdAllowOnlyComments = 1,
        wdAllowOnlyFormFields = 2,
        wdAllowOnlyReading = 3
    }
}
