using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Windows.Automation.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading;

namespace UIAutomationTest1
{
    /*
    class Program
    {
        static System.Windows.Point MousePos
        {
            get
            {
                var pos = Control.MousePosition;
                return new System.Windows.Point(pos.X, pos.Y);
            }
        }

        public static void Main()
        {
            for (;;)
            {
                AutomationElement e = AutomationElement.FromPoint(MousePos);
                if (e != null)
                {
                    object o;
                    if (e.TryGetCurrentPattern(TextPattern.Pattern, out o))
                    {
                        var textPattern = (TextPattern)o;
                        var range = textPattern.RangeFromPoint(MousePos);
                        range.ExpandToEnclosingUnit(TextUnit.Word);
                        Console.WriteLine(range.GetText(-1));
                    }
                }
                System.Threading.Thread.Sleep(1000);
            }
        }
    }
    */

    /*
class Program
{
    static System.Windows.Point MousePos
    {
        get
        {
            var pos = Control.MousePosition;
            return new System.Windows.Point(pos.X, pos.Y);
        }
    }

    public static void Main()
    {
        for (;;)
        {
            AutomationElement e = AutomationElement.FromPoint(MousePos);
            if (e != null)
            {
                foreach (var prop in e.GetSupportedProperties())
                {
                    object o = e.GetCurrentPropertyValue(prop);
                    if (o != null)
                    {
                        var s = o.ToString();
                        if (s != "")
                        {
                            var id = o as AutomationIdentifier;
                            if (id != null) s = id.ProgrammaticName;
                            Console.WriteLine("{0}: {1}", Automation.PropertyName(prop), s);
                        }
                    }
                }
                foreach (var pattern in e.GetSupportedPatterns())
                {
                    Console.WriteLine("Pattern: {0}", Automation.PatternName(pattern));
                }
                Console.WriteLine();
            }
            System.Threading.Thread.Sleep(1000);
        }
    }
}
*/

    class Program
    {

        // Get a handle to an application window.
        /*[DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        */

        private static AutomationElement GetElement(AutomationElement root, string name, bool recursive)
        {
            PropertyCondition condName = new PropertyCondition(AutomationElement.NameProperty, name);
            return root.FindFirst(recursive ? TreeScope.Descendants : TreeScope.Children, condName);
        }
        //Search for the first occurance of a given type of control
        private static AutomationElement GetElement(AutomationElement root, ControlType controlType)
        {
            PropertyCondition condType = new PropertyCondition(AutomationElement.ControlTypeProperty, controlType);
            return root.FindFirst(TreeScope.Descendants, condType);
        }

        static System.Windows.Point MousePos
        {
            get
            {
                var pos = Control.MousePosition;
                return new System.Windows.Point(pos.X, pos.Y);
            }
        }

        public static void Main()
        {


            //Test2Calc();

            //testCalc();

            testOneNote();

            //SearchTest();

            //ActiveWindow aw = new ActiveWindow();
            //aw.GetActiveItemList();

            //AutomationElement desktop = AutomationElement.RootElement;
            //Console.WriteLine(desktop.GetCurrentPropertyValue(AutomationElement.ClassNameProperty) as string);

            //AutomationElement application = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
            //    new PropertyCondition(AutomationElement.ClassNameProperty, "Start"));

            //Console.WriteLine(application.GetCurrentPropertyValue(AutomationElement.NameProperty) as string);

            /*var process = Process.Start(@"C:\Windows\System32\calc.exe");

            Process[] processes;
            AutomationElement calc = null;

            do
            {
                processes = Process.GetProcessesByName("Calculator");
            }
            while (processes.Length == 0);
            

            foreach(var p in processes)
            {
                if(p.MainWindowHandle != IntPtr.Zero)
                {
                    calc = AutomationElement.FromHandle(p.MainWindowHandle);
                }
            }
            */

            //while (process.MainWindowHandle == IntPtr.Zero);

            //AutomationElement calc = AutomationElement.FromHandle(process.MainWindowHandle);

            // Locate calculator window
            /*AutomationElement calculator = GetElement(AutomationElement.RootElement, "Calculator", true);
            // locate the button 7
            AutomationElement btn7 = GetElement(calculator, "Sept", true);
            InvokePattern btn7InvPat = btn7.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btn7InvPat.Invoke();
            // locate the button *
            AutomationElement btnMult = GetElement(calculator, "*", true);
            InvokePattern btnMultInvPat = btnMult.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btnMultInvPat.Invoke();
            // hit on 7 again
            btn7InvPat.Invoke();
            // locate and invoke -
            AutomationElement btnMinus = GetElement(calculator, "-", true);
            InvokePattern btnMinusInvPat = btnMinus.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btnMinusInvPat.Invoke();
            // hit on 7 again
            btn7InvPat.Invoke();
            // locate and invoke =
            AutomationElement btnEq = GetElement(calculator, "=", true);
            InvokePattern btnEqInvPat = btnEq.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            btnEqInvPat.Invoke();
            // get the result edit box and verify the value is indeed 42
            AutomationElement editResult = GetElement(calculator, ControlType.Edit);
            string result = editResult.GetCurrentPropertyValue(ValuePattern.ValueProperty, false).ToString();
            Debug.Assert(result != "42.", result);
            */
            //valPattern.SetValue("abc");

            //session.FindElementByName("Un").Click();
            //session.FindElementByName("Plus").Click();
            //session.FindElementByName("Deux").Click();
            //session.FindElementByName("Est égal à").Click();

            //var calculatorResult = session.FindElementByAccessibilityId("CalculatorResults").Text.Replace("L’affichage est", "").Trim();
            //Console.WriteLine(calculatorResult);



            //session.FindElementByName("Un").Click();
            //session.FindElementByName("Plus").Click();
            //session.FindElementByName("Deux").Click();
            //session.FindElementByName("Est égal à").Click();

            //var calculatorResult = session.FindElementByAccessibilityId("CalculatorResults").Text.Replace("L’affichage est", "").Trim();
            //.WriteLine(calculatorResult);


            //var processes = Process.GetProcessesByName("chrome");
            //var automationElements = from chrome in processes
            //                         where chrome.MainWindowHandle != IntPtr.Zero
            //                         select AutomationElement.FromHandle(chrome.MainWindowHandle);

            //Condition propCondition = new PropertyCondition(
            //AutomationElement.ClassNameProperty, "toto", PropertyConditionFlags.IgnoreCase);

            //rootElement.FindFirst(TreeScope.Element | TreeScope.Children, propCondition);

            //TreeWalker

            /*
            for (;;)
            {
                AutomationElement e = AutomationElement.FromPoint(MousePos);
                if (e != null)
                {
                    Console.WriteLine("Name: {0}",
                     e.GetCurrentPropertyValue(AutomationElement.NameProperty));
                    Console.WriteLine("Value: {0}",
                     e.GetCurrentPropertyValue(ValuePattern.ValueProperty));
                    Console.WriteLine();
                }
                System.Threading.Thread.Sleep(1000);
            }
            */


        }

        static void SearchTest()
        {
            //AutomationElement desktop = AutomationElement.RootElement;
            //Console.WriteLine(desktop.GetCurrentPropertyValue(AutomationElement.ClassNameProperty) as string);

            AutomationElement Start = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Start"));

            ((InvokePattern)Start.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

            SendKeys.SendWait("word{ENTER}");

            //SendKeys.SendWait("{ENTER}");

            //SendKeys.Send("{ENTER}");

            //Zone de recherche
        }

        /*
        static void Test2Calc()
        {
            var process = Process.Start("calc.exe");

            Thread.Sleep(3000);

            Process[] processes;

            do
            {
                processes = Process.GetProcessesByName("Calculator");
            }
            while (processes.Length == 0);

            

            IntPtr calculatorHandle = processes[0].MainWindowHandle;

            // Verify that Calculator is a running process.
            if (calculatorHandle == IntPtr.Zero)
            {
                MessageBox.Show("Calculator is not running.");
                return;
            }

            // Make Calculator the foreground application and send it 
            // a set of calculations.
            //SetForegroundWindow(calculatorHandle);
            SendKeys.SendWait("111");
            SendKeys.SendWait("*");
            SendKeys.SendWait("11");
            SendKeys.SendWait("=");

        }
        */

        static void testOneNote()
        {

            AutomationElement Start = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Start"));

            ((InvokePattern)Start.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

            SendKeys.SendWait("one note{ENTER}");

            Process[] processes;

            do
            {
                processes = Process.GetProcessesByName("ONENOTE");
            }
            while (processes.Length == 0);


            //Thread.Sleep(2000);

            //AutomationElement calc = AutomationElement.FromHandle(processes[0].MainWindowHandle);
            AutomationElement OneNote = null;

            do
            {
                OneNote = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ProcessIdProperty, processes[0].Id));
            }
            while (OneNote == null);

            Thread.Sleep(2000);

            var NetUISectionTabs = OneNote.FindAll(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ClassNameProperty, "NetUISectionTab"));

            foreach(AutomationElement NetUISectionTab in NetUISectionTabs)
            {
                (NetUISectionTab.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern).Invoke();
            }

           /* AutomationElement NetUINotebookDropdownButton = OneNote.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "API - OneNote"));

            (NetUINotebookDropdownButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern).Invoke();
           */

        }

        static void testCalc()
        {

            var process = Process.Start("calc.exe");

            Process[] processes;

            do
            {
                processes = Process.GetProcessesByName("Calculator");
            }
            while (processes.Length == 0);


            //Thread.Sleep(2000);

            //AutomationElement calc = AutomationElement.FromHandle(processes[0].MainWindowHandle);
            AutomationElement calc = null;

            do
            {
                calc = AutomationElement.RootElement.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ProcessIdProperty, processes[0].Id));
            }
            while (calc == null);


            AutomationElement un = calc.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Un"));

            var invokePattern = un.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            AutomationElement bouton = calc.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Plus"));

            invokePattern = bouton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            bouton = calc.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Deux"));

            invokePattern = bouton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            bouton = calc.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Est égal à"));

            invokePattern = bouton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();


            bouton = calc.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "CalculatorResults"));

            //string result = bouton.GetCurrentPropertyValue(ValuePattern.ValueProperty, false).ToString();
            //Console.WriteLine(result);

            //string result = bouton.GetCurrentPattern(ValuePattern.ValueProperty, false).ToString();
            //Console.WriteLine(result);

            //ValuePattern valPattern = bouton.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            Console.WriteLine(bouton.GetCurrentPropertyValue(AutomationElement.NameProperty).ToString().Replace("L’affichage est", "").Trim());
            //valPattern.SetValue("abc");
        }
    }

}
