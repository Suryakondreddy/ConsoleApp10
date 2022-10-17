using Castle.Core.Internal;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;
using NHibernate.Driver;
using NHibernate.Mapping;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.iOS;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using Windows.Devices.SmartCards;
using Windows.System;
using Windows.UI.Shell;
using Windows.UI.Xaml;
using static System.Collections.Specialized.BitVector32;
using System.Xml.Serialization;


namespace SandR
{
    class Program
    {
        static void Main(string[] args)
        {

            WindowsDriver<WindowsElement> SandR;
            AppiumOptions desiredCapabilities = new AppiumOptions();

            //DesiredCapabilities appCapabilities = new DesiredCapabilities();


            //appCapabilities.SetCapability("appWorkingDir", "C:\\Program Files (x86)\\ReSound\\Palpatine6.7.4.21-RP-S\\AlgoLabtest.exe");

            //desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Lucan\App\Lucan.App.UI.exe");

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Camelot\WorkflowRuntime\Camelot.WorkflowRuntime.exe");


            desiredCapabilities.AddAdditionalCapability("createSessionTimeout", "60000");
            desiredCapabilities.AddAdditionalCapability("newCommandTimeout", "60000");
            desiredCapabilities.AddAdditionalCapability("platformName", "Windows");
            desiredCapabilities.AddAdditionalCapability("deviceName", "WindowsPC");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            //WebDriverWait waitForMe = new WebDriverWait();

            var wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(200));

            var dv = SandR.FindElementByName("LT988-DW (Final)");
            dv.Click();

            var sdi = SandR.FindElementByName("Continue >>");
            sdi.Click();
            Thread.Sleep(2000);

            SandR.SwitchTo().Window(SandR.WindowHandles.First());
            SandR.SwitchTo().ActiveElement();

            var r = SandR.FindElementByName("Read");
            r.Click();
            Thread.Sleep(2000);

            sdi = SandR.FindElementByName("Continue >>");
            sdi.Click();

            Thread.Sleep(2000);

            SandR.SwitchTo().Window(SandR.WindowHandles.First());
            SandR.SwitchTo().ActiveElement();

            Thread.Sleep(2000);

            if (SandR != null)
            {
                var sz = SandR.FindElementByName("Full");
                sz.Click();
                Thread.Sleep(3000);

            }
            else

                Thread.Sleep(3000);

            SandR.SwitchTo().Window(SandR.WindowHandles.First());
            SandR.SwitchTo().ActiveElement();
            Thread.Sleep(3000);

            SandR.Keyboard.SendKeys("1234");


            Thread.Sleep(3000);

            var si = SandR.FindElementByName("Continue >>");
            si.Click();

            Thread.Sleep(3000);

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Camelot\TestRuntimePC\Camelot.TestRuntimePC.exe");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            Thread.Sleep(2000);

            wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(200));

            wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("Close")));

            var cs = SandR.FindElementByName("Close");
            cs.Click();




            var bt = SandR.FindElementByName("Services");
            bt.Click();
            Thread.Sleep(15000);
            var btnrt = SandR.FindElementByName("Restore");
            btnrt.Click();
            Thread.Sleep(15000);
            var btnlo = SandR.FindElementByName("LOGIN REQUIRED");
            btnlo.Click();
            Thread.Sleep(2000);
            var textpass = SandR.FindElementByAccessibilityId("PasswordBox");
            textpass.SendKeys("svk01");
            Thread.Sleep(2000);
            //var btnlog = SandR.FindElementByName("LOGIN REQUIRED");
            //btnlog.Click();
            Thread.Sleep(1000);
            SandR.SwitchTo().Window(SandR.WindowHandles.First());
            SandR.SwitchTo().ActiveElement();
            Thread.Sleep(5000);
            SandR.FindElementByName("LOGIN").Click();
            // var btnlogin = sandr.findelementbyname("login");
            // btnlogin.click();
            Thread.Sleep(10000);
            var btnred = SandR.FindElementByName("READ");
            btnred.Click();
            Thread.Sleep(2000);
            var btnfnd = SandR.FindElementByName("FIND");
            btnfnd.Click();
            Thread.Sleep(8000);
            var btnslt = SandR.FindElementByName("SELECT");
            btnslt.Click();

            Thread.Sleep(2000);

            Actions actions2 = new Actions(SandR);

            var element = SandR.FindElementByName("RESTORE");

            actions2.MoveToElement(element);
            actions2.Perform();

            var sts = wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Name("OK")));

            // Thread.Sleep(50000);

            //var btnyes = SandR.FindElementByName("Yes");
            //btnyes.Click();

            //Thread.Sleep(20000);


            var txtresult = SandR.FindElementByAccessibilityId("textBlockMessage");
            Console.WriteLine($"Result shown by SandR:{txtresult.Text}");

            ////if (txtresults.text.endswith("operation completed successfully"))
            ////{
            ////    console.writeline("capture succeeded.");
            ////}
            ////else
            ////{
            ////    console.writeline("capture failed.");
            ////}

            var btno = SandR.FindElementByName("OK");
            btno.Click();

            Thread.Sleep(5000);

            var btncls = SandR.FindElementByAccessibilityId("PART_Close");
            btncls.Click();

            Thread.Sleep(1000);

            string path = ($"C:\\Users\\Public\\Documents\\Camelot\\Logs\\{Environment.MachineName}-{Environment.UserName}-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");

            if (File.Exists(path))
            {
                string last = File.ReadLines(path).Last();
                Console.WriteLine(last);
                string str = last;
                var result = str.LastIndexOf(':');
                int lastDotIndex = str.LastIndexOf(":", System.StringComparison.Ordinal);
                string firstPart = str.Remove(lastDotIndex);
                string secondPart = str.Substring(lastDotIndex + 1, str.Length - firstPart.Length - 1);
                //Console.WriteLine(last.LastIndexOf(':'));

                //foreach (var value in result)
                //    Console.WriteLine(value);
                int timeTaken = 0;

                if (result != null)
                {
                    try
                    {
                        timeTaken = Convert.ToInt32(secondPart.Replace(" ", "").Replace("|", ""));
                    }
                    catch (Exception ex)
                    {
                        //
                        throw ex;
                    }
                    if (timeTaken < 41000)
                        Console.WriteLine("Restore Operation performed in desired time");
                    else
                        Console.WriteLine("Failed");
                }

            }

            else
            {
                Console.WriteLine("file not found");
            }






            //WebDriverWait waitForMe = new WebDriverWait();
            //WebDriverWait webDriverWait = new WebDriverWait(SandR, new TimeSpan.Fromseconds(10));


            //var reorder = SandR.FindElementByName("ReorderGrip");
            //reorder.Click();

            //Thread.Sleep(1000);

            var pli = SandR.FindElementsByClassName("ListBoxItem");

            Thread.Sleep(2000);
            Actions actions = new Actions(SandR);
            WindowsElement menuHoverLink = pli[2]; //SandR.FindElement(By.linkText("Menu heading"));
            actions.MoveToElement(menuHoverLink).Perform();

            pli[2].FindElement(By.Name("Fit Patient")).Click();

            Thread.Sleep(10000);

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\ReSound\SmartFit\SmartFit.exe");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            var ID = SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems");
            ID.Click();

            Thread.Sleep(2000);

            var speed = SandR.FindElementByName("Speedlink");
            speed.Click();

            Thread.Sleep(2000);

            var connect = SandR.FindElementByName("Connect to ReSound Smart Fit");
            connect.Click();

            Thread.Sleep(5000);

            List<string> lstWindow = SandR.WindowHandles.ToList();
            foreach (string window in lstWindow)
            {

                SandR.SwitchTo().Window(windowName: window);


                //Thread.Sleep(10000);


                do
                {

                    wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(200));
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("Continue")));
                    var winname = SandR.FindElementsByClassName("TextBlock");
                    Console.WriteLine($"window name : {winname[1].Text}");
                    var cnti = SandR.FindElementByName("Continue");
                    cnti.Click();

                    //var cs = SandR.FindElementByName("Continue");
                    //cs.Click();
                    ////Thread.Sleep(10000);

                }

                while (!SandR.FindElements(By.Name("Continue")).IsNullOrEmpty());
            }


            var cnt = SandR.FindElementByAccessibilityId("ProgramStripAutomationIds.AddProgramAction");
            cnt.Click();

            Thread.Sleep(5000);

            cnt = SandR.FindElementByName("Outdoor");
            cnt.Click();

            Thread.Sleep(5000);

            cnt = SandR.FindElementByName("Save");
            cnt.Click();

            Thread.Sleep(5000);

            cnt = SandR.FindElementByName("Skip & Save");
            cnt.Click();

            Thread.Sleep(10000);

            cnt = SandR.FindElementByName("Exit ReSound Smart Fit");
            cnt.Click();

            Thread.Sleep(10000);




            //WindowsDriver<WindowsElement> SandR;
            //AppiumOptions desiredCapabilities = new AppiumOptions();

            //// DesiredCapabilities appCapabilities = new DesiredCapabilities();


            //// appCapabilities.SetCapability("appWorkingDir", "C:\\Program Files (x86)\\ReSound\\Palpatine6.4.4.16-RP-S\\AlgoLabtest.exe");

            //desiredCapabilities.AddAdditionalCapability("app", @"C:\Service & Repair Tool  4.6.exe");

            //SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            //var wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(30));






            //WindowsDriver<WindowsElement> SandR;
            //AppiumOptions desiredCapabilities = new AppiumOptions();

            DesiredCapabilities appCapabilities = new DesiredCapabilities();


            //appCapabilities.SetCapability("appWorkingDir", "C:\\Program Files (x86)\\ReSound\\Palpatine6.7.4.21-RP-S\\AlgoLabtest.exe");

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Camelot\WorkflowRuntime\Camelot.WorkflowRuntime.exe");

            desiredCapabilities.AddAdditionalCapability("createSessionTimeout", "60000");
            desiredCapabilities.AddAdditionalCapability("newCommandTimeout", "60000");
            desiredCapabilities.AddAdditionalCapability("platformName", "Windows");
            desiredCapabilities.AddAdditionalCapability("deviceName", "WindowsPC");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            //var wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(200));

            //var reorder = SandR.FindElementByName("ReorderGrip");
            //reorder.Click();

            //Thread.Sleep(1000);


            //var dv = SandR.FindElementByName("LT988-DW (Final)");
            //dv.Click();

            //var sdi = SandR.FindElementByName("Continue >>");
            //sdi.Click();
            //Thread.Sleep(2000);

            //SandR.SwitchTo().Window(SandR.WindowHandles.First());
            //SandR.SwitchTo().ActiveElement();

            //var r = SandR.FindElementByName("Read");
            //r.Click();
            //Thread.Sleep(2000);

            //sdi = SandR.FindElementByName("Continue >>");
            //sdi.Click();

            //Thread.Sleep(2000);

            //SandR.SwitchTo().Window(SandR.WindowHandles.First());
            //SandR.SwitchTo().ActiveElement();

            //Thread.Sleep(2000);

            //if (SandR != null)
            //{
            //    var sz = SandR.FindElementByName("Full");
            //    sz.Click();
            //    Thread.Sleep(3000);

            //}
            //else

            //    Thread.Sleep(3000);

            //SandR.SwitchTo().Window(SandR.WindowHandles.First());
            //SandR.SwitchTo().ActiveElement();
            //Thread.Sleep(3000);

            //SandR.Keyboard.SendKeys("1234");


            //Thread.Sleep(3000);

            //var si = SandR.FindElementByName("Continue >>");
            //si.Click();

            //Thread.Sleep(3000);

            //desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Camelot\TestRuntimePC\Camelot.TestRuntimePC.exe");

            //SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            //Thread.Sleep(2000);

            //wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(200));

            //wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("Close")));

            //var cs = SandR.FindElementByName("Close");
            //cs.Click();

            









            //var pli = SandR.FindElementsByClassName("ListBoxItem");

            //Thread.Sleep(2000);
            //Actions actions = new Actions(SandR);
            //WindowsElement menuHoverLink = pli[2]; //SandR.FindElement(By.linkText("Menu heading"));
            //actions.MoveToElement(menuHoverLink).Perform();

            pli[2].FindElement(By.Name("Fit Patient")).Click();


            Thread.Sleep(10000);

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\ReSound\SmartFit\SmartFit.exe");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);


            //var ID = SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems");
            //ID.Click();

            //Thread.Sleep(2000);

            //var speed = SandR.FindElementByName("Speedlink");
            //speed.Click();

            Thread.Sleep(2000);

            //var connect = SandR.FindElementByName("Connect to ReSound Smart Fit");
            connect.Click();

            Thread.Sleep(5000);

            //desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\ReSound\SmartFit\SmartFitSA.exe");

            //SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

            //Thread.Sleep(5000);

            //List<string> lstWindow = SandR.WindowHandles.ToList();
            foreach (string window in lstWindow)
            {

                SandR.SwitchTo().Window(windowName: window);
                Thread.Yield();

                do
                {
                    Thread.Yield();
                    var winname = SandR.FindElementsByClassName("TextBlock");
                    Console.WriteLine($"window name : {winname[1].Text}");
                    Thread.Yield();
                    var cnti = SandR.FindElementByName("Continue");
                    cnti.Click();
                    Thread.Yield();

                    //waitForMe.Until(pred => txtLocation.Enabled);


                }
                while (!SandR.FindElements(By.Name("Continue")).IsNullOrEmpty());

               // while (SandR.FindElements(By.Name("Continue")).EqualsIsNullOrEmpty()) ;


            }




            //var cnt = SandR.FindElementByAccessibilityId("ProgramStripAutomationIds.AddProgramAction");
            //cnt.Click();

            Thread.Sleep(5000);

            cnt = SandR.FindElementByName("Outdoor");
            cnt.Click();

            Thread.Sleep(5000);

            //cnt = SandR.FindElementByName("Save");
            //cnt.Click();

            //Thread.Sleep(5000);

            //cnt = SandR.FindElementByName("Skip & Save");
            //cnt.Click();

            //Thread.Sleep(15000);

            //cnt = SandR.FindElementByName("Exit ReSound Smart Fit");
            //cnt.Click();

            //Thread.Sleep(15000);






            ////var cnt = SandR.FindElementByName("Continue");
            ////cnt.Click();

            ////Thread.Sleep(10000);

            ////cnt = SandR.FindElementByName("Continue");
            ////cnt.Click();

            ////Thread.Sleep(10000);

            ////cnt = SandR.FindElementByName("Continue");
            ////cnt.Click();

            ////Thread.Sleep(15000);

            ////cnt = SandR.FindElementByName("Continue");
            ////cnt.Click();

            ////Thread.Sleep(15000);

            ////cnt = SandR.FindElementByName("Continue");
            ////cnt.Click();

            ////Thread.Sleep(10000);























            //AppiumOptions windowOptions = null;

            //WindowsDriver<WindowsElement> sessionSecondBest = null;

            //var listofAllWindows = SandR.FindElementsByClassName("Window");

            //Debug.WriteLine($"Elements found: {listofAllWindows.Count}");

            //foreach (var window in listofAllWindows)
            //{
            //    Console.WriteLine($"{window.Text}");
            //    if (window.Displayed && window.Text.Contains("Smart Launcher")) ;
            //    {
            //        var windowhandle = window.GetAttribute("NativeWindowHandle");
            //        Console.WriteLine($"Windowhandle:{windowhandle}");
            //        var handleint = int.Parse(windowhandle).ToString("X");
            //        windowOptions = new AppiumOptions();
            //        windowOptions.AddAdditionalCapability("appTopLevelWindow", handleint);
            //        break;

            //    }
            //}

            //if (windowOptions != null)
            //{
            //    sessionSecondBest = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), windowOptions);



            //}


            //var ID = SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems");
            //ID.Click();




            //List<string> lstWindow = SandR.WindowHandles.ToList();

            //foreach (string window in lstWindow)

            //{
            //    SandR.SwitchTo().Window(windowName: window);
            //}


            //Console.WriteLine("Number of windows opened by SmartFit" + SandR.WindowHandles.Count);

            //foreach (var item in SandR.WindowHandles)
            //{
            //    Console.WriteLine(item);
            //}
            //Console.WriteLine("current window handle" + SandR.CurrentWindowHandle);



            //Thread.Sleep(10000);

            //Console.WriteLine("Number of windows opened by SmartFit" + SandR.WindowHandles.Count);

            //foreach (var item in SandR.WindowHandles)
            //{
            //    Console.WriteLine(item);
            //}

            //SandR.SwitchTo().Window(SandR.CurrentWindowHandle);

            //Thread.Sleep(10000);

            //Console.WriteLine(SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems"));
            //Console.WriteLine("current window handle" + SandR.CurrentWindowHandle);
            //Thread.Sleep(3000);

            //pli[2].FindElementByName("Fit Patient").Click();
            //Thread.Sleep(10000);



            //var WindowHandle = SandR.WindowHandles;
            //Thread.Sleep(TimeSpan.FromSeconds(10));
            //var allWindowHandles2 = SandR.WindowHandles;
            //SandR.SwitchTo().Window(allWindowHandles2[0]);

            //var win = SandR.FindElementsByClassName("Window");
            //win[0].FindElementByWindowsUIAutomation("Smart Launcher");

            //Thread.Sleep(3000);



            //Thread.Sleep(5000);

            //SandR.SwitchTo().Window(windowName:"Smart Launcher");

            //SandR.SwitchTo().Window(windowName:"Smart Launcher");



            //Actions actions1 = new Actions(SandR);
            //WindowsElement menuHoverLink1 = SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems");

            //actions1.MoveToElement(menuHoverLink1).Perform();

            //var rn = SandR.FindElementByName("SmartFitSA - 2 running windows");
            //rn.Click();




            //SandR.SwitchTo().Window(SandR.WindowHandles.First());
            //SandR.SwitchTo().Window("Smart Launcher").Manage();


            var ft = SandR.FindElementByAccessibilityId("ConnectionAutomationIds.CommunicationInterfaceItems");
            ft.Click();

            Thread.Sleep(5000);

            var fit = SandR.FindElementByName("Speedlink");
            fit.Click();
            Thread.Sleep(2000);


            var patient = SandR.FindElementByName("ALL");
            patient.Click();

            Thread.Sleep(2000);

            var mx = SandR.FindElementsByClassName("TextBlock");
            mx[0].Click();

            Thread.Sleep(5000);

            var btnDevi = SandR.FindElementByName("Channel");
            btnDevi.Click();

            Thread.Sleep(1000);

            var chd = SandR.FindElementByName("ChannelDropDown");
            chd.Click();

            Thread.Sleep(1000);

            var ine = SandR.FindElementByName("SpeedLink:0");
            Thread.Sleep(1000);

            //var drp = SandR.FindElementByName("SpeedLink:0DropDown");
            //Thread.Sleep(1000);

            //var cik = SandR.FindElementByName("Left");
            //cik.Click();

            //Thread.Sleep(1000);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);
            SandR.Keyboard.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);
            SandR.Keyboard.SendKeys(Keys.Enter);
            Thread.Sleep(1000);



            //SandR.Keyboard.SendKeys(Keys.ArrowRight);
            //Thread.Sleep(1000);
            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(1000);




            var fine = SandR.FindElementByName("File");
            fine.Click();

            Thread.Sleep(2000);

            //var dpdn = SandR.FindElementByName("FileDropDown");
            ////dpdn.Click();

            //Thread.Sleep(1000);

            //var read = SandR.FindElementByName("Read from SpeedLink:0/Left");
            //read.Click();
            //Thread.Sleep(3000);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);
            SandR.Keyboard.SendKeys(Keys.ArrowDown);
            Thread.Sleep(1000);
            SandR.Keyboard.SendKeys(Keys.Enter);
            Thread.Sleep(5000);

            ine = SandR.FindElementByName("File");
            ine.Click();
            Thread.Sleep(2000);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);

            SandR.Keyboard.SendKeys(Keys.ArrowDown);
            Thread.Sleep(3000);

            SandR.Keyboard.SendKeys(Keys.Enter);

            Thread.Sleep(3000);

            var pd = SandR.FindElementByName("3e900:004c ProductionTestData");
            pd.Click();

            Thread.Sleep(2000);

            var data = SandR.FindElementByName("DataGridView");
            data.Click();
            var row = SandR.FindElementByName("Row 0");
            row.Click();
            var row1 = SandR.FindElementByName("Value Row 0");
            row1.SendKeys("2022-07-05 09:54:05Z");

            //var row = SandR.FindElementByName("Value Row 0");
            //row.SendKeys("2022-08-05 09:54:05Z");

            Thread.Sleep(3000);

            pd = SandR.FindElementByName("3e900:004c ProductionTestData");
            pd.Click();

            Thread.Sleep(3000);

            Actions action = new Actions(SandR);

            action.ContextClick(SandR.FindElement(By.Name("3e900:004c ProductionTestData"))).SendKeys(Keys.ArrowDown).SendKeys(Keys.ArrowDown).Build().Perform();
            SandR.Keyboard.SendKeys(Keys.Enter);

            // action.ContextClick(SandR.FindElement(By.Name("3e900:004c ProductionTestData")));
            //SandR.Keyboard.SendKeys(Keys.Enter);

            Thread.Sleep(2000);

            var min = SandR.FindElementByName("2a000:0026 MiniIdentification");
            min.Click();

            Thread.Sleep(2000);

            data = SandR.FindElementByName("DataGridView");
            data.Click();
            row = SandR.FindElementByName("Row 6");
            row.Click();
            row1 = SandR.FindElementByName("Value Row 6");
            row1.SendKeys("1659673446");
            Thread.Sleep(2000);

            min = SandR.FindElementByName("2a000:0026 MiniIdentification");
            min.Click();

            Thread.Sleep(3000);

            Actions action1 = new Actions(SandR);

            action1.ContextClick(SandR.FindElement(By.Name("2a000:0026 MiniIdentification"))).SendKeys(Keys.ArrowDown).SendKeys(Keys.ArrowDown).Build().Perform();
            SandR.Keyboard.SendKeys(Keys.Enter);


            Thread.Sleep(3000);

            var cl = SandR.FindElementByName("Close");
            cl.Click();

            desiredCapabilities.AddAdditionalCapability("app", @"C:\Program Files (x86)\GN Hearing\Lucan\App\Lucan.App.UI.exe");

            SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredCapabilities);

           // var wait = new WebDriverWait(SandR, TimeSpan.FromSeconds(30));


            var btnDeviceInfo = SandR.FindElementByName("Device Info");
            btnDeviceInfo.Click();

            Thread.Sleep(2000);

           // var cnt = SandR.FindElementByName("Connect");
           // cnt.Click();

            Thread.Sleep(30000);






            //var options = SandR.FindElement(By.Name("Check All nodes"));
            //options.Click();

            ////foreach (var option in options)
            //{
            //    option.SendKeys(Keys.ArrowDown);
            //}




            var ck = SandR.FindElementByName("Check All nodes");
            ck.Click();

            Thread.Sleep(3000);








            var sd = SandR.FindElementByName("Read from SpeedLink:0/Left");
            sd.Click();

            Thread.Sleep(7000);



            //var btnDeviceInfo = SandR.FindElementByAccessibilityId("labtestButton");
            //btnDeviceInfo.Click();


            Thread.Sleep(3000);

            var btnServices = SandR.FindElementByName("ADL");
            btnServices.Click();

            Thread.Sleep(1000);
            btnDeviceInfo = SandR.FindElementByAccessibilityId("labtestButton");
            btnDeviceInfo.Click();

            Thread.Sleep(2000);

            var cb = SandR.FindElementByAccessibilityId("BUTTON");
            cb.Click();

            Thread.Sleep(5000);

            var DI = SandR.FindElementsByAccessibilityId("DeviceIcon");
            //DI.Click();


            var intr = SandR.FindElementsByClassName("Image");
            

            var i = cb.FindElementByName("Image");

            Thread.Sleep(5000);

            var inter = SandR.FindElementsByClassName("UserControl");



            SandR.SwitchTo().Window(SandR.WindowHandles.First());
            //SandR.SwitchTo().();

            var uc = SandR.FindElementByClassName("Image");
            uc.Click();
            

                



            Thread.Sleep(1000);

            var im = SandR.FindElementsByClassName("Image");





            //DesiredCapabilities desktopCapabilities = new DesiredCapabilities();
            //desktopCapabilities.SetCapability("app", "Root");
            //SandR = new IOSDriver<IOSElement>(new Uri(""), desktopCapabilities);
            //SandR.FindElementByName("AlgoLabtest").Click();

            //SandR.FindElementByName("AlgoLabtest").Click();


            //AppiumOptions desiredcapabilities = new AppiumOptions();

            //desiredcapabilities.AddAdditionalCapability("appWorkingDir", @"C:\Program Files (x86)\ReSound\Palpatine6.4.4.16-RP-S");

            //SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredcapabilities);

            Thread.Sleep(1000);

            //var btnSettings = SandR.FindElementByName("Uninstall");
            //btnSettings.Click();

            //Thread.Sleep(8000);

            //var btnDeviceInfo = SandR.FindElementByName("Settings");
           // btnDeviceInfo.Click();
           // Thread.Sleep(2000);

            var c = SandR.FindElementsByClassName("ComboBox");

            c[1].Click();

            SandR.FindElementByName("SpeedLink").Click();

            c[2].Click();

            SandR.FindElementByName("Right").Click();

            c[3].Click();

            SandR.FindElementByName("Advanced").Click();

            c[0].Click();

            SandR.FindElementByName("Orange").Click();

            c[4].Click();

            SandR.FindElementByName("Automatically");


            var t = SandR.FindElementByName("Generate restoration PDF report");
            t.Click();

            var g = SandR.FindElementByName("Generate a ticket file");
            g.Click();




            //var n = SandR.FindElementByName("SpeedLink");
            //n.Click();

            Thread.Sleep(2000);

            var ch = SandR.FindElementsByClassName("Channels view");
            

            Thread.Sleep(2000);

            var n1 = SandR.FindElementByName("Right");
            n1.Click();

            var comboNumber4 = SandR.FindElementByClassName(nameof(ComboBox));
            comboNumber4.Click();



            var comboNumber5 = SandR.FindElementByClassName(nameof(ComboBox));
            comboNumber5.Click();

            comboNumber5.FindElementByName("Orange").Click();

             


            //var txtname = SandR.FindElementByClassName("DeviceInfoView");
            //txtname.Click();

            //var txtModelName = SandR.FindElementByClassName("DeviceInfoView");

            ////Console.WriteLine($"Result shown by SandR:{txtname.Text}");

            //var mname = SandR.FindElementByName("Model Name");
            //Console.WriteLine(mname.Text);


            //Console.WriteLine($"{mname}");



            //var res = SandR.FindElementByClassName("TextBox");
            //Console.WriteLine($"

            //var name = SandR.FindElementByName("Model Name");
            //name.Click();

            //var box = SandR.FindElementByClassName("TextBox");
            //box.ToString();

            //Console.WriteLine("box");


            // var btndiagnostics = SandR.FindElementByName("Diagnostics");
            //btndiagnostics.Click();

            //Thread.Sleep(5000);

            //var btndeviceinfo = SandR.FindElementByName("Device Info");
            //btndeviceinfo.Click();

            //Thread.Sleep(5000);

            //var btnrefresh = SandR.FindElementByName("Refresh");
            //btnrefresh.Click();

            //Thread.Sleep(10000);


            //var btnservices = SandR.FindElementByName("Services");
            //btnservices.Click();

            //Thread.Sleep(5000);

            //var btncapture = SandR.FindElementByName("Capture");
            //btncapture.Click();

            //Thread.Sleep(15000);

            //var btnlogin = SandR.FindElementByName("LOGIN REQUIRED");
            //btnlogin.Click();

            //Thread.Sleep(2000);

            //var textpassword = SandR.FindElementByAccessibilityId("PasswordBox");
            //textpassword.SendKeys("svk01");

            //Thread.Sleep(2000);

            //var btnlog = SandR.FindElementByName("LOGIN REQUIRED");
            //btnlog.Click();

            //Thread.Sleep(1000);

            //SandR.SwitchTo().Window(SandR.WindowHandles.First());

            //SandR.SwitchTo().ActiveElement();

            //Thread.Sleep(5000);

            //SandR.FindElementByName("LOGIN").Click();


            //// var btnlogin = sandr.findelementbyname("login");
            //// btnlogin.click();

            //Thread.Sleep(10000);

            //var btncap = SandR.FindElementByName("CAPTURE");
            //btncap.Click();

            //Thread.Sleep(40000);

            ////var btnyes = SandR.FindElementByName("Yes");
            ////btnyes.Click();

            ////Thread.Sleep(20000);

            //var txtresults = SandR.FindElementByAccessibilityId("textBlockMessage");
            //Console.WriteLine($"Result shown by SandR:{txtresults.Text}");

            ////if (txtresults.text.endswith("operation completed successfully"))
            ////{
            ////    console.writeline("capture succeeded.");
            ////}
            ////else
            ////{
            ////    console.writeline("capture failed.");
            ////}

            //var btnok = SandR.FindElementByName("OK");
            //btnok.Click();

            //Thread.Sleep(5000);

            //var btncls = SandR.FindElementByAccessibilityId("PART_Close");
            //btncls.Click();

            //Thread.Sleep(1000);
            //string path = $"c:\\users\\public\\documents\\camelot\\logs\\DKCPHHPF2PNCT7-xxdayred-{DateTime.Now.ToString():YYYY-MM-dd}.log";
            //string path = ($"C:\\Users\\Public\\Documents\\Camelot\\Logs\\{Environment.MachineName}-{Environment.UserName}-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");

            //if (File.Exists(path))
            //{
            //   string last = File.ReadLines(path).Last();
            //    Console.WriteLine(last);
            //    String str = last;
            //    var result = str.LastIndexOf(':');
            //    int lastDotIndex = str.LastIndexOf(":", System.StringComparison.Ordinal);
            //    string firstPart = str.Remove(lastDotIndex);
            //    string secondPart = str.Substring(lastDotIndex + 1, str.Length - firstPart.Length - 1);
            //    Console.WriteLine(last.LastIndexOf(':'));

            //    //foreach (var value in result)
            //    //    Console.WriteLine(value);
            //int timeTaken = 0;

            //if (result != null)
            //{
            //    try
            //    {
            //        timeTaken = Convert.ToInt32(secondPart.Replace(" ", "").Replace("|", ""));
            //    }
            //    catch (Exception ex)
            //    {
            //        //
            //        throw ex;
            //    }
            //    if (timeTaken < 41000)
            //        Console.WriteLine("capture Operation performed in desired time");
            //    else
            //        Console.WriteLine("Failed");
            //}

            //    //int temp = 41000;
            //    //while (temp < 41000)
            //    //{
            //    //    temp++;
            //    //    //Console.WriteLine("The value of desired time is : " + word);
            //    //    //word = word + 1;
            //    //    //if (word < 41000)
            //    //    //    break;
            //    //}
            //    //Console.WriteLine(temp);

            //    //if (str.Contains(str))
            //    //{
            //    //    Console.WriteLine("capture identity performed in desired time");

            //    //}
            //    //else
            //    //{
            //    //    Console.WriteLine("Capture identity failed");
            //    //}
            //}

            //else
            //{
            //    Console.WriteLine("file not found");
            //}



            //RESTORE

            //var btns = SandR.FindElementByName("Settings");
            //btnservices.Click();
            //Thread.Sleep(5000);


            //var boxsite = SandR.FindElementByClassName("ComboBox");
            //boxsite.Click();
            //Thread.Sleep(2000);
            //var tst = SandR.FindElementByName("TEST");
            //tst.SendKeys(Keys.Enter);

            //Thread.Sleep(2000);

            //boxsite.SendKeys(Keys.Tab);
            //SandR.Keyboard.SendKeys(Keys.ArrowUp);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.Tab);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowUp);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.Tab);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);


            //var cnt = SandR.FindElementByName("Connect to hearing instrument automatically");
            //cnt.Click();
            //Thread.Sleep(2000);


            //SandR.Keyboard.SendKeys(Keys.Tab);
            //Thread.Sleep(2000);

            ////SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //SandR.Keyboard.SendKeys(Keys.ArrowUp);

            //SandR.Keyboard.SendKeys(Keys.ArrowUp);

            //SandR.Keyboard.SendKeys(Keys.ArrowUp);

            //Thread.Sleep(2000);

            ////SandR.Keyboard.SendKeys(Keys.ArrowUp);
            ////SandR.Keyboard.SendKeys(Keys.ArrowUp);
            //SandR.Keyboard.SendKeys(Keys.Tab);
            //SandR.Keyboard.SendKeys(Keys.Tab);

            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowDown);
            //Thread.Sleep(2000);

            //SandR.Keyboard.SendKeys(Keys.ArrowUp);
            //Thread.Sleep(5000);


            ////SandR.Keyboard.SendKeys(Keys.Tab);

            ////Thread.Sleep(2000);


            ////SandR.Keyboard.SendKeys(Keys.Tab);


            //var restorationpdf = SandR.FindElementByName("Generate restoration PDF report");
            //restorationpdf.Click();

            //Thread.Sleep(3000);



            //var tctf = SandR.FindElementByName("Generate a ticket file");
            //tctf.Click();

            //Thread.Sleep(3000);
            ////device.Click();

            //var dev = SandR.FindElementByName("SpeedLink");
            //dev.Click();


            //tst = SandR.FindElementByName("SpeedLink");
            //tst.Click();
            //tst.SendKeys(Keys.ArrowUp);


            //device.SendKeys(Keys.ArrowUp);  
            //device.SendKeys(Keys.Enter);

            //boxsite = SandR.FindElementByClassName("ComboBox");
            //boxsite.Click();


            //Thread.Sleep(2000);

            //var device = SandR.FindElementByName("SpeedLink");
            //device.Click();
            //device.SendKeys(Keys.Enter);
            //Thread.Sleep(2000);

            //var nxtc = SandR.FindElementByName("SpeedLink");
            //nxtc.SendKeys(Keys.Enter);

            //boxsite.SendKeys(Keys.ArrowLeft);

            //var device = SandR.FindElementByName("SpeedLink");
            //device.Click();
            //device.SendKeys(Keys.Enter);
            //Thread.Sleep(2000);

            //boxsite.SendKeys(Keys.Tab);
            //boxsite.SendKeys(Keys.ArrowDown);
            //boxsite.SendKeys(Keys.ArrowUp);


            // var dev = SandR.FindElementByClassName("ComboBox1");
            // dev.Click();

            //Thread.Sleep(2000);

            //btnservices.SendKeys(Keys.Tab);
            //var intrabox = SandR.FindElementByClassName("ComboBox");
            //intrabox.Click();

            //Thread.Sleep(2000);

            //var device = SandR.FindElementByName("SpeedLink");

            //device.SendKeys(Keys.Enter);

           //     var bt = SandR.FindElementByName("Services");
           //     bt.Click();
           //     Thread.Sleep(15000);
           //     var btnrt = SandR.FindElementByName("Restore");
           //     btnrt.Click();
           //     Thread.Sleep(15000);
           //     var btnlo = SandR.FindElementByName("LOGIN REQUIRED");
           //     btnlo.Click();
           //     Thread.Sleep(2000);
           //     var textpass = SandR.FindElementByAccessibilityId("PasswordBox");
           //     textpass.SendKeys("svk01");
           //     Thread.Sleep(2000);
           //     //var btnlog = SandR.FindElementByName("LOGIN REQUIRED");
           //     //btnlog.Click();
           //     Thread.Sleep(1000);
           //     SandR.SwitchTo().Window(SandR.WindowHandles.First());
           //     SandR.SwitchTo().ActiveElement();
           //     Thread.Sleep(5000);
           //     SandR.FindElementByName("LOGIN").Click();
           //     // var btnlogin = sandr.findelementbyname("login");
           //     // btnlogin.click();
           //     Thread.Sleep(10000);
           //     var btnred = SandR.FindElementByName("READ");
           //     btnred.Click();
           //     Thread.Sleep(2000);
           //     var btnfnd = SandR.FindElementByName("FIND");
           //     btnfnd.Click();
           //     Thread.Sleep(8000);
           //     var btnslt = SandR.FindElementByName("SELECT");
           //     btnslt.Click();

           //     Thread.Sleep(2000);

           //     var element = SandR.FindElementByName("RESTORE");
           //     Actions actions2 = new Actions(SandR);
           //     actions.MoveToElement(element);
           //     actions.Perform();


           // //SandR.Keyboard.SendKeys(Keys.Enter);


           // //var btnres = SandR.FindElementsByClassName("Button");

           // //Thread.Sleep(5000);

           // //btnres[18].Click();


           // var sts = wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Name("OK")));

           //// Thread.Sleep(50000);

           //     //var btnyes = SandR.FindElementByName("Yes");
           //     //btnyes.Click();

           //     //Thread.Sleep(20000);

                
           //     var txtresult = SandR.FindElementByAccessibilityId("textBlockMessage");
           //     Console.WriteLine($"Result shown by SandR:{txtresult.Text}");

           //     ////if (txtresults.text.endswith("operation completed successfully"))
           //     ////{
           //     ////    console.writeline("capture succeeded.");
           //     ////}
           //     ////else
           //     ////{
           //     ////    console.writeline("capture failed.");
           //     ////}

           //     var btno = SandR.FindElementByName("OK");
           //     btno.Click();

           //     Thread.Sleep(5000);

           //     var btncls = SandR.FindElementByAccessibilityId("PART_Close");
           //     btncls.Click();

           //     Thread.Sleep(1000);

           //     string path = ($"C:\\Users\\Public\\Documents\\Camelot\\Logs\\{Environment.MachineName}-{Environment.UserName}-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");

           //     if (File.Exists(path))
           //     {
           //         string last = File.ReadLines(path).Last();
           //         Console.WriteLine(last);
           //         string str = last;
           //         var result = str.LastIndexOf(':');
           //         int lastDotIndex = str.LastIndexOf(":", System.StringComparison.Ordinal);
           //         string firstPart = str.Remove(lastDotIndex);
           //         string secondPart = str.Substring(lastDotIndex + 1, str.Length - firstPart.Length - 1);
           //         //Console.WriteLine(last.LastIndexOf(':'));

           //         //foreach (var value in result)
           //         //    Console.WriteLine(value);
           //         int timeTaken = 0;

           //         if (result != null)
           //         {
           //             try
           //             {
           //                 timeTaken = Convert.ToInt32(secondPart.Replace(" ", "").Replace("|", ""));
           //             }
           //             catch (Exception ex)
           //             {
           //                 //
           //                 throw ex;
           //             }
           //             if (timeTaken < 41000)
           //                 Console.WriteLine("Restore Operation performed in desired time");
           //             else
           //                 Console.WriteLine("Failed");
           //         }

           //     }

           //     else
           //     {
           //         Console.WriteLine("file not found");
           //     }



                //desiredcapabilities.AddAdditionalCapability("app", @"C:\Users\Public\Documents\Camelot\Logs\DKCPHHPF2PNCT7-xxdayred-2022-08-23.log");

                //SandR = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), desiredcapabilities);

                //SandR.SwitchTo().Window(SandR.WindowHandles.First());

                //SandR.SwitchTo().ActiveElement();

                //Thread.Sleep(5000);



                //var tx = SandR.FindElementByAccessibilityId("TitleBar");
                //tx.Click();

                //Thread.Sleep(2000);

                //var txt = SandR.FindElementByAccessibilityId("15");
                //txt.Click();

                //Thread.Sleep(2000);

                //String path = "C:\\Users\\Public\\Public Documents\\Camelot\\Logs\\DKCPHHPF2PNCT7-xxdayred-2022-08-23.log.txt";


                // string lastLine = File.ReadLines("C:\\Users\\Public\\Public Documents\\Camelot\\Logs\\DKCPHHPF2PNCT7-xxdayred-2022-08-23.log.txt").Last();

                //String last = File.ReadLines("").Last();
                //Console.WriteLine(lastLine);



                //String wordstobesearch = ("Total time for Capturing the hearing instrument identity in milliseconds <= 40000");


                // Console.WriteLine($"value shown by 15 {txt.Text}");

                //if (path.Length <= 40000) ;

                //{
                //   Console.WriteLine("Capture identity performed in desired time");
                // }
                //else
                //{
                //    Console.WriteLine("Capture identity operation failed");
                //}




                //String.("Total time for Capturing the hearing instrument identity in milliseconds <= 40000");

                //String wordstobesearch = "Total time for Capturing the hearing instrument identity in milliseconds: 40000";










                //var importantLines = File.ReadLines(@"C:\Users\Public\Documents\Camelot\Logs\DKCPHHPF2PNCT7-xxdayred-2022-08-22.log")
                //.SkipWhile(line => !line.Contains("CustomerEN"))
                //.TakeWhile(line => !line.Contains("Total time for Capturing the hearing instrument identity in milliseconds"));

                //var txtResult = SandR.FindElementByName("Total time for Capturing the hearing instrument identity in milliseconds");

                //Console.WriteLine($"Result shown after Capture Operation:{txtResults.Text}");

                //if (txtResults.Text.Equals("40000"))
                //{
                //     Console.WriteLine("Capture performed within desired time.");
                // }
                //else
                //{
                //    Console.WriteLine("Capture not performed in desired time.");
                //}


                //private string GetSasReadToken(string connectionString, string containerName, string blobName)
                //{

                //    var storageAccount = CloudStorageAccount.Parse(connectionString);
                //    CloudBlobClient cloudBlobClient = storageAccount.CreateCloudBlobClient();
                //    CloudBlobContainer cloudBlobContainer = cloudBlobClient.GetContainerReference(containerName);
                //    CloudBlockBlob cloudBlockBlob = cloudBlobContainer.GetBlockBlobReference(blobName);
                //    var sasConstraints = new SharedAccessBlobPolicy
                //    {
                //        SharedAccessExpiryTime = DateTimeOffset.UtcNow.AddMinutes(60),
                //        Permissions = SharedAccessBlobPermissions.Read,
                //    };
                //    var sasHeaders = new SharedAccessBlobHeaders();
                //    sasHeaders.ContentDisposition = "attachment;filename=<your-download-file-name>";
                //    var sasBlobToken = cloudBlockBlob.GetSharedAccessSignature(sharedAccessBlobPolicy, sasHeaders);
                //}


            
        }
    }
}

























