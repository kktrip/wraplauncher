using System.IO;
using System.Text.Json;
using JE = System.Text.Encodings.Web;
using UN = System.Text.Unicode;

namespace WrapLauncher.Settings
{
    /// <summary>
    /// アプリケーション設定書き込み
    /// </summary>
    public class AppSettingsWriter
    {
        /// <summary>
        /// ファイル書き込み
        /// </summary>
        /// <param name="stg"></param>
        public void WriteToFile(AppSettings stg)
        {
            var opt = new JsonSerializerOptions
            {
                // シリアライズするUnicodeの範囲
                Encoder = JE.JavaScriptEncoder.Create(UN.UnicodeRanges.All),
                // インデントする
                WriteIndented = true,
            };

            // JSONオブジェクトを文字列化
            string jsonStr = JsonSerializer.Serialize(stg, opt);
            // ファイル書き込み
            File.WriteAllText(App.GetAppSettingsFilePath(), jsonStr);
        }
    }
}