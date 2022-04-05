using System.IO;
using System.Text.Json;

namespace WrapLauncher.Settings
{
    /// <summary>
    /// アプリケーション設定読み込み
    /// </summary>
    public class AppSettingsReader
    {
        /// <summary>
        /// ファイル存在確認
        /// </summary>
        /// <returns></returns>
        public bool ExistsFile()
        {
            return File.Exists(App.GetAppSettingsFilePath());
        }

        /// <summary>
        /// ファイル読み込み
        /// </summary>
        /// <returns>設定情報</returns>
        public AppSettings ReadFromFile()
        {
            if (!ExistsFile())
            {
                // ファイルがなければ初期値を返す
                return new AppSettings();
            }

            // 読み込み
            string jsonStr = File.ReadAllText(App.GetAppSettingsFilePath());
            var stg = JsonSerializer.Deserialize<AppSettings>(jsonStr);

            return stg ?? new AppSettings();
        }
    }
}