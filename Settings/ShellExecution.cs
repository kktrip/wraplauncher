using System.Diagnostics;

namespace WrapLauncher
{

    public static class ShellExecution
    {
        /// <summary>
        /// プログラム実行
        /// </summary>
        /// <param name="cmd">実行するコマンド</param>
        public static void Run(string cmd)
        {
            var p = new Process();
            p.StartInfo.FileName = cmd;
            p.StartInfo.UseShellExecute = true;
            // 実行
            p.Start();
        }
    }

}