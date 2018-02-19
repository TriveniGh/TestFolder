using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercer.Us.Ben.Automation
{
   public class Constants
    {
        /// <summary>
        /// A timeout value for ImplicitWaitTime (60).
        /// </summary>
       public const int ImplicitTimeout = 60;

       /// <summary>
       /// A timeout value for PageLoadTimeout (12)
       /// </summary>
       public const int PageLoadTimeout = 12;
        //WaitforthePageLoad

       public const int WaitforthePageLoad = 240;
       public const int ElementPresent = 10;

       public static  bool ApplicationLaunchFlag = true;
    }
}
