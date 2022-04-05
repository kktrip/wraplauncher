using System;
using System.IO;
using System.Windows;

namespace WrapLauncher
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// アプリケーション設定ファイル名
        /// </summary>
        public const string AppSettingsFileName = "appsettings.json";

        /// <summary>
        /// ランチャー定義ファイル名
        /// </summary>
        public const string LaunchDefFileName = "WrapLauncher.path";

        /// <summary>
        /// アプリケーションパス取得
        /// </summary>
        /// <returns></returns>
        public static string GetAppPath()
        {
            string? appPath = AppContext.BaseDirectory;
            if (appPath is null)
            {
                throw new DirectoryNotFoundException("実行ファイルのパス取得失敗");
            }

            return appPath;
        }

        /// <summary>
        /// アプリケーション設定ファイルパス取得
        /// </summary>
        /// <returns></returns>
        public static string GetAppSettingsFilePath()
        {
            return Path.Combine(GetAppPath(), AppSettingsFileName);
        }

        /// <summary>
        /// ランチャー定義ファイルパス取得
        /// </summary>
        /// <returns></returns>
        public static string GetLaunchDefFilePath()
        {
            return Path.Combine(GetAppPath(), LaunchDefFileName);
        }
    }
}