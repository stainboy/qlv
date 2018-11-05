using qvl2mp4.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace qvl2mp4
{
    class Program
    {
        static void Main(string[] args)
        {
            DoConvertProcess();
        }

        private static void DoConvertProcess()
        {
            var src = Directory.GetFiles(Settings.Default.VideoSourceDir, "*.qlv", SearchOption.AllDirectories);
            foreach (var qvl in src)
            {
                Console.WriteLine($"Processing {qvl}...");
                var mp4 = TransformMP4Ext(qvl);
                if (File.Exists(mp4))
                {
                    Console.WriteLine($"Target file {mp4} exits");
                }
                else
                {
                    ManipulateVC(qvl);
                    ValidateTarget(mp4);
                }
            }
            Console.WriteLine("Bye");
        }

        private static void ValidateTarget(string mp4)
        {
            if (File.Exists(mp4))
            {
                Console.WriteLine($"Target file {mp4} was created");

                //var proc = new Process
                //{
                //    StartInfo = new ProcessStartInfo
                //    {
                //        FileName = Settings.Default.FFProbeExePath,
                //        Arguments = mp4,
                //        UseShellExecute = false,
                //        RedirectStandardOutput = true,
                //        CreateNoWindow = true
                //    }
                //};
                //proc.Start();
                //proc.WaitForExit();
                //while (!proc.StandardOutput.EndOfStream)
                //{
                //    string line = proc.StandardOutput.ReadLine();
                //    Console.WriteLine(line);
                //}
                //proc.Dispose();
            }
            else
            {
                Console.WriteLine($"Target file {mp4} was not successfully created");
            }
        }

        private static void ManipulateVC(string qvl)
        {
            foreach (var vc in Process.GetProcessesByName("VideoConverter"))
            {
                try
                {
                    vc.Kill();
                }
                catch (Exception)
                {
                }
            }
            var p = Process.Start(Settings.Default.VideoConverterExePath);

            while (p.MainWindowHandle == IntPtr.Zero)
            {
                p.Refresh();
                Thread.Sleep(200);
            }

            // Click "添加文件"
            var e = AutomationElement.FromHandle(p.MainWindowHandle);
            var add = e.FindFirst(TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, "添加文件"))
                );
            var ip = add.GetCurrentPattern(InvokePatternIdentifiers.Pattern) as InvokePattern;

            new Thread(() =>
            {
                try
                {
                    ip.Invoke();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: {0}", ex.Message);
                }
            }).Start();
            Thread.Sleep(5000);

            // Find ChooseFileDialog
            var cfd = e.FindFirst(TreeScope.Children,
               new AndCondition(
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                   new PropertyCondition(AutomationElement.ClassNameProperty, "#32770"))
               );
            var inputFile = cfd.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "1148"));
            var vp = inputFile.GetCurrentPattern(ValuePatternIdentifiers.Pattern) as ValuePattern;
            vp.SetValue(qvl);

            // Click "Open"
            var openBtn = cfd.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "1"));
            var iip = openBtn.GetCurrentPattern(InvokePatternIdentifiers.Pattern) as InvokePattern;
            iip.Invoke();
            Thread.Sleep(3000);

            // Find Convert Dialog
            e = AutomationElement.FromHandle(p.MainWindowHandle);
            var cvd = e.FindFirst(TreeScope.Children,
            // new AndCondition(
            //     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                 new PropertyCondition(AutomationElement.ClassNameProperty, "Qt5QWindow") //)
             );
            // Find Start Button
            var startTxt = cvd.FindFirst(TreeScope.Descendants,
             new AndCondition(
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                 new PropertyCondition(AutomationElement.NameProperty, "开始转码"))
             );
            TreeWalker walker = TreeWalker.ControlViewWalker;
            var parent = walker.GetParent(startTxt);
            //var parent = startTxt.FindFirst(TreeScope.Parent, Condition.TrueCondition);
            var startBtn = parent.FindFirst(TreeScope.Children, Condition.TrueCondition);
            var startBtnIP = startBtn.GetCurrentPattern(InvokePatternIdentifiers.Pattern) as InvokePattern;
            startBtnIP.Invoke();

            // Find "继续转换" and click it
            Thread.Sleep(2000);
            e = AutomationElement.FromHandle(p.MainWindowHandle);
            var confirmDialog = e.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Qt5QWindow")
            );
            var contBtns = confirmDialog.FindAll(TreeScope.Descendants,
               //new AndCondition(
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)//,
                  // new PropertyCondition(AutomationElement.NameProperty, "继续转换"))
               );
            //foreach (AutomationElement btn in contBtns)
            //{
            //    Console.WriteLine($"button name -> {btn.Current.Name}");
            //}
            var contBtn = contBtns[1];
            //Console.WriteLine($"button name -> {contBtn.Current.Name}");
            var contBtnIP = contBtn.GetCurrentPattern(InvokePatternIdentifiers.Pattern) as InvokePattern;
            contBtnIP.Invoke();


            // wait until done
            var counter = 0;
            while(IsStillRunning(p))
            {
                Thread.Sleep(2000);
                counter++;
                if (counter %5 == 0)
                {
                    Console.WriteLine($"Waiting for convert complete ({counter/5})...");
                }
            }
            Console.WriteLine("Convert completed");
            p.Kill();
            Thread.Sleep(5000);

            //throw new NotImplementedException();
        }

        private static bool IsStillRunning(Process p)
        {
           var e = AutomationElement.FromHandle(p.MainWindowHandle);
            var cvd = e.FindFirst(TreeScope.Children,
            // new AndCondition(
            //     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane),
                 new PropertyCondition(AutomationElement.ClassNameProperty, "Qt5QWindow") //)
             );
            // Find Start Button
            var startTxt = cvd.FindFirst(TreeScope.Descendants,
             new AndCondition(
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text),
                 new PropertyCondition(AutomationElement.NameProperty, "开始转码"))
             );
            return startTxt == null;
        }

        private static string TransformMP4Ext(string qvl)
        {
            return qvl.Substring(0, qvl.Length - 3) + "MP4";
        }
    }
}
