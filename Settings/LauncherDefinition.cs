using System.Collections.Generic;

namespace WrapLauncher.Settings
{
    /// <summary>
    /// ランチャー定義情報
    /// </summary>
    public class LauncherDefinition
    {
        /// <summary>
        /// 区切り文字
        /// </summary>
        public const char Delimiter = '\t';

        /// <summary>
        /// グループ見出しの先頭記号
        /// </summary>
        public const string GroupTitleHeader = "//";

        /// <summary>
        /// グループ見出しのカラム位置
        /// </summary>
        public const int GroupTitleColumnIndex = 0;

        /// <summary>
        /// カラム位置
        /// </summary>
        public static IReadOnlyDictionary<string, int> Columns = new Dictionary<string, int>
        {
            {"Color", 0},
            {"ButtonTitle", 1},
            {"Path", 2},
        };

        /// <summary>
        /// グループ見出し判定
        /// </summary>
        /// <param name="values"></param>
        /// <returns>グループ見出しならtrue</returns>
        public bool IsGroupTitle(string[] values)
        {
            return values[GroupTitleColumnIndex].StartsWith(GroupTitleHeader);
        }

        /// <summary>
        /// グループ見出し取得
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public string GetGroupTitle(string[] values)
        {
            return values[GroupTitleColumnIndex].Substring(GroupTitleHeader.Length);
        }

    }
}