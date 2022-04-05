namespace WrapLauncher.Settings
{
    /// <summary>
    /// アプリケーション設定
    /// </summary>
    public class AppSettings
    {
        /// <summary>
        /// アプリ起動後にランチャー最小化
        /// </summary>
        public bool MinimizedAfterLaunch { get; set; } = false;

        /// <summary>
        /// 表示倍率
        /// </summary>
        /// <value></value>
        public double Scale { get; set; } = 1.0;
    }
}