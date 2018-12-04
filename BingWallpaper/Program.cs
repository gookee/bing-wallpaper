using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

namespace BingWallpaper
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Mutex mutex = new Mutex(false, "b1c063de-2104-468e-ab02-4ca06b0c213e");
            try
            {
                if (mutex.WaitOne(0, false))
                {
                    //获得当前登录的Windows用户标示
                    System.Security.Principal.WindowsIdentity identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                    //创建Windows用户主题
                    Application.EnableVisualStyles();

                    System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(identity);
                    //判断当前登录用户是否为管理员
                    if (principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator))
                    {
                        //如果是管理员，则直接运行
                        Application.EnableVisualStyles();
                        Application.Run(new MainForm(new BingImageProvider(), new Settings()));
                    }
                    else
                    {
                        //创建启动对象
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        //设置运行文件
                        startInfo.FileName = Application.ExecutablePath;
                        //设置启动动作,确保以管理员身份运行
                        startInfo.Verb = "runas";
                        //如果不是管理员，则启动UAC
                        Process.Start(startInfo);
                    }
                }
                else
                {
                    MessageBox.Show(Resource.AppStatus, Resource.AppName);
                }
            }
            finally
            {
                if (mutex != null)
                {
                    mutex.Close();
                    mutex = null;
                }
            }
        }
    }
}
