﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mercer.Us.Ben.Automation.Main;


namespace Mercer.Us.Ben.Automation.Scheduler
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class ScheduleRun
    {

        [TestMethod]
        public void StartScheduleRun()
        {
            Test execute = new Test();
            execute.RunTest("Smoke");

        }
        

    }
}
