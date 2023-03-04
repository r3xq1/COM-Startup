namespace ComRega
{
    using System;

    public static class COMStartupLib
    {
        /// <summary>
        /// windows styles
        /// </summary>
        enum ShortcutWindowStyles : int
        {
            /// <summary>
            /// Hide
            /// </summary>
            WshHide = 0,
            /// <summary>
            /// NormalFocus
            /// </summary>
            WshNormalFocus = 1,
            /// <summary>
            /// MinimizedFocus
            /// </summary>
            WshMinimizedFocus = 2,
            /// <summary>
            /// MaximizedFocus
            /// </summary>
            WshMaximizedFocus = 3,
            /// <summary>
            /// NormalNoFocus
            /// </summary>
            WshNormalNoFocus = 4,
            /// <summary>
            /// MinimizedNoFocus
            /// </summary>
            WshMinimizedNoFocus = 6,
        }

        /*  Include: IWshRuntimeLibrary -> References -> COM > Windows Script Host Object Model */

        //public static void ComAddToStartup(bool addOrDelete, string pathToFile, string startupName, string descript = "")
        //{
        //    try
        //    {
        //        string fullPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), $"{startupName}.lnk");
        //        Type m_type = Type.GetTypeFromProgID("WScript.Shell");
        //        IWshShortcut shortcut = (IWshShortcut)m_type.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, Activator.CreateInstance(m_type), new object[] { fullPath });

        //        if (addOrDelete)
        //        {
        //            shortcut.Description = !string.IsNullOrWhiteSpace(descript) ? descript : null;
        //            //shortcut.IconLocation = @"C:\Windows\System32\user32.dll,1";//  @"C:\Windows\System32\user32.dll,1";
        //            shortcut.IconLocation = typeof(Program).Assembly.Location; // Добавляет иконку от Вашего приложения.
        //            shortcut.TargetPath = pathToFile;
        //            shortcut.WindowStyle = (int)ShortcutWindowStyles.WshHide;
        //            shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(pathToFile);
        //            shortcut.Save();
        //        }
        //        else
        //        {
        //            if (System.IO.File.Exists(fullPath))
        //            {
        //                System.IO.File.Delete(fullPath);
        //            }
        //        }
        //    }
        //    catch { }
        //}
    }
}
