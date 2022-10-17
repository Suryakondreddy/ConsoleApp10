using AventStack.ExtentReports;
using AventStack.ExtentReports.Model;
using AventStack.ExtentReports.Reporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using System;

namespace Reportextent
{
    [TestClass]
    public class Report
    {
        public ExtentReports extent;
        public ExtentTest test;

        [OneTimeSetUp]
        public void StartReport()
        {
            string pth = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            string actualpath = pth.Substring(0, pth.LastIndexOf("bin"));
            string projectpath = new Uri(actualpath).LocalPath;

            string reportpath = projectpath + "Reports\\MyOwnReport.html";

            extent = new ExtentReports();
            var extentHtml = new ExtentHtmlReporter(reportpath);
            extent.AttachReporter(extentHtml);


            extentHtml.LoadConfig(projectpath + "extent-config.xml");

            extent.AddSystemInfo("Host Name", "Surya");
            extent.AddSystemInfo("Environemnt", "QA");
            extent.AddSystemInfo("User Name", "Surya Kondreddy");

            extent.CreateTest(projectpath + "extent-config.xml");
        }

        [NUnit.Framework.Test]
        public void ReportPass()
        {
            test = extent.CreateTest("ReportPass");
            NUnit.Framework.Assert.IsTrue(true);
            test.Log(Status.Pass, "Assert pass as condition is True");

        }

        [NUnit.Framework.Test]

        public void ReportFail()
        {
            test = extent.CreateTest("ReportFail");
            NUnit.Framework.Assert.IsTrue(false);
            test.Log(Status.Pass, "Assert pass as condition is False");

        }

        [TearDown]
        public void GetResult()
        {
            var status = NUnit.Framework.TestContext.CurrentContext.Result.Outcome.Status;
            var stackTrace = "<pre>" + NUnit.Framework.TestContext.CurrentContext.Result.StackTrace + "</pre>";
            var errormessage = NUnit.Framework.TestContext.CurrentContext.Result.Message;

            if (status == NUnit.Framework.Interfaces.TestStatus.Failed)
            {
                test.Log(Status.Fail, status + errormessage);
            }


        }

        [OneTimeTearDown]

        public void EndReport()
        {
            extent.Flush();

        }
    }
}
