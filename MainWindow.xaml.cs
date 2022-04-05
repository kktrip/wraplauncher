using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using WrapLauncher.Settings;
using System.Diagnostics;

namespace WrapLauncher
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private readonly LauncherDefinition _btnDef = new();

		// ヘルプテキスト
		private const string HelpText = @"設定ファイルはEXEと同じ場所に「WrapLauncher.path」のファイル名で配置する。
文字コードはBOM無しのUTF8。

データ構造
------------------------------
//グループ見出し
色名  ボタンテキスト   起動プログラム/フォルダのフルパス
色名  ボタンテキスト   起動プログラム/フォルダのフルパス
：
//グループ見出し
色名  ボタンテキスト   起動プログラム/フォルダのフルパス
：
------------------------------
グループ見出しは先頭「//」で始める。
色名、ボタンテキスト、起動するパスはTABで区切る。

[色名一覧]
";

		private int _buttonCount = 0;
		private int _groupCount = 0;

		/// <summary>
		/// コンストラクタ
		/// </summary>
		public MainWindow()
		{
			InitializeComponent();
		}

		/// <summary>
		/// 起動時
		/// </summary>
		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				// アプリケーション設定読み込み、画面に反映
				LoadAppSettings();
				// ボタン設定読み込み、画面に反映
				LoadLauncherDef();
			}
			catch (System.Exception ex)
			{
				ShowException(ex);
			}
		}

		/// <summary>
		/// 終了時
		/// </summary>
		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			try
			{
				// アプリケーション設定保存
				SaveAppSettings();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(
					$"アプリケーション設定の保存に失敗しました。\n{ex}",
					"エラー", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		/// <summary>
		/// コンテキストメニュー「設定再読み込み」
		/// </summary>
		private void MenuReload_Click(object sender, RoutedEventArgs e)
		{
			// 画面クリア
			MainContainer.Children.Clear();

			try
			{
				// 設定読み込み、画面に反映
				LoadLauncherDef();
			}
			catch (System.Exception ex)
			{
				ShowException(ex);
			}
		}

		/// <summary>
		/// コンテキストメニュー「ランチャーの場所を開く」
		/// </summary>
		private void MenuFolderOpen_Click(object sender, RoutedEventArgs e)
		{
			string appPath = App.GetAppPath();
			Execute(appPath);
		}

		/// <summary>
		/// コンテキストメニュー「表示倍率」
		/// </summary>
		private void ScaleItem_Click(object sender, RoutedEventArgs e)
		{
			string? scaleStr = (sender as RadioButton)?.Content?.ToString();
			if (!double.TryParse(scaleStr, out double scale))
			{
				scale = 1.0;
				DefaultScaleItem.IsChecked = true;
			}
			SetContainerScale(scale);
			WindowContextMenu.IsOpen = false;
		}

		/// <summary>
		/// コンテナ拡大縮小処理
		/// </summary>
		/// <param name="scaleXY"></param>
		private void SetContainerScale(double scaleXY)
		{
			MainContainer.LayoutTransform = new ScaleTransform(scaleXY, scaleXY);
		}

		/// <summary>
		/// コンテキストメニュー「情報」
		/// </summary>
		private void MenuInfo_Click(object sender, RoutedEventArgs e)
		{
			MessageBox.Show($"グループ数：{_groupCount}\nボタン数：{_buttonCount}", "情報");
		}

		/// <summary>
		/// コンテキストメニュー「ヘルプ」
		/// </summary>
		private void MenuHelp_Click(object sender, RoutedEventArgs e)
		{
			ShowHelp();
		}

		/// <summary>
		/// エラー情報表示
		/// </summary>
		private void ShowException(System.Exception ex)
		{
			MainContainer.Children.Clear();

			var txt = new TextBox
			{
				TextWrapping = TextWrapping.Wrap,
				Margin = new Thickness(5),
				IsReadOnly = true,
				Text = ex.ToString()
			};

			MainContainer.Children.Add(txt);
		}

		/// <summary>
		/// ヘルプ内容表示
		/// </summary>
		private void ShowHelp()
		{
			MainContainer.Children.Clear();

			var helpText = new RichTextBox
			{
				IsReadOnly = true,
				FontSize = 16
			};

			// テキストの説明文追加
			helpText.Document.Blocks.Add(new Paragraph(new Run(HelpText)));

			var bc = new BrushConverter();
			var docList = new System.Windows.Documents.List();

			// 色名一覧作成
			var bList = typeof(Brushes).GetProperties()
				.Where(x => x.Name != "Transparent")
				.OrderBy(x => x.Name);
			foreach (var b in bList)
			{
				var li = new ListItem();
				var p = new Paragraph();
				var r = new Run("■ ")
				{
					Foreground = (Brush)bc.ConvertFromString(b.Name)!
				};
				p.Inlines.Add(r);
				p.Inlines.Add(new Run(b.Name));

				li.Blocks.Add(p);
				docList.ListItems.Add(li);
			}

			helpText.Document.Blocks.Add(docList);
			MainContainer.Children.Add(helpText);
		}

		/// <summary>
		/// アプリケーション設定読み込み、画面に反映
		/// </summary>
		private void LoadAppSettings()
		{
			var asr = new AppSettingsReader();
			SetScreen(asr.ReadFromFile());
		}

		/// <summary>
		/// アプリケーション設定を画面に反映
		/// </summary>
		/// <param name="appStg"></param>
		private void SetScreen(AppSettings appStg)
		{
			// アプリを起動したらランチャー最小化
			MinimizedMenuItem.IsChecked = appStg.MinimizedAfterLaunch;
			// 表示倍率
			DefaultScaleItem.IsChecked = true;
			foreach (var child in LogicalTreeHelper.GetChildren(ScaleMenu))
			{
				if (child is RadioButton radio)
				{
					if (double.TryParse(radio.Content.ToString(), out double scale))
					{
						if (scale == appStg.Scale)
						{
							radio.IsChecked = true;
							SetContainerScale(scale);
							break;
						}
					}
				}
			}
		}

		/// <summary>
		/// アプリケーション設定保存
		/// </summary>
		private void SaveAppSettings()
		{
			// 画面から設定取得
			var appStg = GetScreen();

			var asw = new AppSettingsWriter();

			// 書き込み
			asw.WriteToFile(appStg);
		}

		/// <summary>
		/// 画面からアプリケーション設定取得
		/// </summary>
		/// <returns></returns>
		private AppSettings GetScreen()
		{
			var appStg = new AppSettings
			{
				// アプリを起動したらランチャー最小化
				MinimizedAfterLaunch = MinimizedMenuItem.IsChecked
			};
			// 表示倍率
			double scale = 1.0;
			foreach (var child in LogicalTreeHelper.GetChildren(ScaleMenu))
			{
				if (child is RadioButton radio)
				{
					if (radio.IsChecked == true)
					{
						if (double.TryParse(radio.Content.ToString(), out scale))
						{
							break;
						}
					}
				}
			}
			appStg.Scale = scale;

			return appStg;
		}

		/// <summary>
		/// ランチャー定義読み込み、画面に反映
		/// </summary>
		private void LoadLauncherDef()
		{
			// 設定ファイルのフルパス取得
			string filePath = App.GetLaunchDefFilePath();

			// 設定ファイルが見つからなければ終了
			if (!File.Exists(filePath))
			{
				throw new FileNotFoundException("設定ファイルなし");
			}

			// 設定ファイル読み込み
			using var reader = new StreamReader(filePath);
			WrapPanel? btnContainer = null;

			_buttonCount = 0;
			_groupCount = 0;




			// Outlook空きスケジュール抽出　コンテナ作成
			var grpOutlookContainer = new StackPanel();
			// グループコンテナをメインコンテナに追加
			MainContainer.Children.Add(grpOutlookContainer);
			// グループ見出しを生成し、グループコンテナに追加
			grpOutlookContainer.Children.Add(CreateGroupTitle("Outlook"));
			WrapPanel? btnOutlookContainer = new WrapPanel();
			// ボタンコンテナをグループコンテナに追加
			grpOutlookContainer.Children.Add(btnOutlookContainer);
			// ボタン作成
			Button outlookBtn = CreateOutlookButton("Blue", "スケジュール空き日時取得");
			// ボタンコンテナにボタンを追加
			grpOutlookContainer?.Children.Add(outlookBtn);






			while (!reader.EndOfStream)
			{
				// TABで分解
				var item = reader.ReadLine()?.Split(LauncherDefinition.Delimiter);
				if (item is null ||
					item.Length < 1)
				{
					// データなしの行
					continue;
				}

				/*
                MainContainer           ..... StackPanel
                    grpContainer        ..... StackPanel
                        見出し          ..... TextBlock
                        btnContainer    ..... WrapPanel
                            ボタン
                            ボタン
                            ：
                    grpContainer        ..... StackPanel
                        見出し          ..... TextBlock
                        btnContainer    ..... WrapPanel
                            ボタン
                            ボタン
                            ：
                */

				// グループ作成するローカル関数
				void MakeGroup(ref WrapPanel? btnContainer, string grpTitle = "")
				{
					// グループコンテナ作成
					var grpContainer = new StackPanel();
					// グループコンテナをメインコンテナに追加
					MainContainer.Children.Add(grpContainer);

					if (string.IsNullOrEmpty(grpTitle) == false)
					{
						// グループ見出しを生成し、グループコンテナに追加
						grpContainer.Children.Add(CreateGroupTitle(grpTitle));
					}

					// ボタンコンテナを作成
					btnContainer = new WrapPanel();
					// ボタンコンテナをグループコンテナに追加
					grpContainer.Children.Add(btnContainer);
					_groupCount++;
				}

				if (_btnDef.IsGroupTitle(item))
				{
					// 見出し

					// グループ作成
					MakeGroup(ref btnContainer, _btnDef.GetGroupTitle(item));
				}
				else if (item.Length == LauncherDefinition.Columns.Count)
				{
					// ボタン

					// まだボタンコンテナが作成されていない？
					if (btnContainer is null)
					{
						// グループ作成（見出し無し）
						MakeGroup(ref btnContainer);
					}

					// ボタン作成
					Button btn = CreateLaunchButton(
						item[LauncherDefinition.Columns["Color"]],
						item[LauncherDefinition.Columns["ButtonTitle"]],
						item[LauncherDefinition.Columns["Path"]]);
					// ボタンコンテナにボタンを追加
					btnContainer?.Children.Add(btn);
					_buttonCount++;
				}
				else
				{
					throw new System.Exception($"カラム数不正\n{string.Join(LauncherDefinition.Delimiter, item)}");
				}
			}

			if (_buttonCount == 0 &&
				_groupCount == 0)
			{
				throw new System.Exception("定義内容なし");
			}
		}

		/// <summary>
		/// グループ見出し生成
		/// </summary>
		/// <param name="title"></param>
		/// <returns></returns>
		private TextBlock CreateGroupTitle(string title)
		{
			var txt = new TextBlock
			{
				Style = (Style)(this.Resources["GroupTitleStyle"]),
				Text = title
			};

			return txt;
		}

		/// <summary>
		/// ボタン作成
		/// </summary>
		/// <param name="colorName">色名</param>
		/// <param name="text">ボタンテキスト</param>
		/// <param name="execute">起動プログラムのファイルパス</param>
		private Button CreateLaunchButton(string colorName, string text, string execute)
		{
			var btn = new Button();
			var txtContainer = new StackPanel
			{
				Orientation = Orientation.Horizontal
			};
			// ■テキスト作成
			var txtMark = new TextBlock
			{
				Text = "■",
				Margin = new Thickness(0, 0, 2, 0)
			};

			try
			{
				// 色指定が有効なら■の色を変える。無効な色名なら例外発生で終了させる。
				var bcnv = new BrushConverter();
				txtMark.Foreground = (Brush)bcnv.ConvertFromString(colorName)!;
			}
			catch
			{
				throw new System.Exception($"無効な色名[{colorName}]");
			}

			txtContainer.Children.Add(txtMark);

			// ボタン名テキスト作成
			var txt = new TextBlock
			{
				Text = text
			};
			txtContainer.Children.Add(txt);

			// ボタンテキスト設定
			btn.Content = txtContainer;
			// ボタンクリック時の処理
			btn.Click += (_, _) =>
			{
				if (Execute(execute) &&
					MinimizedMenuItem.IsChecked &&
					!Keyboard.IsKeyDown(Key.LeftCtrl) &&
					!Keyboard.IsKeyDown(Key.RightCtrl))
				{
					// ウィンドウ最小化（Ctrl+クリックの場合は最小化しない）
					WindowState = WindowState.Minimized;
				}
			};

			return btn;
		}

		/// <summary>
		/// ボタン作成
		/// </summary>
		/// <param name="colorName">色名</param>
		/// <param name="text">ボタンテキスト</param>
		/// <param name="execute">起動プログラムのファイルパス</param>
		private Button CreateOutlookButton(string colorName, string text)
		{
			var btn = new Button();
			var txtContainer = new StackPanel
			{
				Orientation = Orientation.Horizontal
			};
			// ■テキスト作成
			var txtMark = new TextBlock
			{
				Text = "■",
				Margin = new Thickness(0, 0, 2, 0)
			};

			try
			{
				// 色指定が有効なら■の色を変える。無効な色名なら例外発生で終了させる。
				var bcnv = new BrushConverter();
				txtMark.Foreground = (Brush)bcnv.ConvertFromString(colorName)!;
			}
			catch
			{
				throw new System.Exception($"無効な色名[{colorName}]");
			}

			txtContainer.Children.Add(txtMark);

			// ボタン名テキスト作成
			var txt = new TextBlock
			{
				Text = text
			};
			txtContainer.Children.Add(txt);

			// ボタンテキスト設定
			btn.Content = txtContainer;
			// ボタンクリック時の処理
			btn.Click += (_, _) =>
			{
				try
				{
					var stdOut = "";
					var stdErr = "";
					var exitCode = 0;
					var ps1FilePath = "C:\\Users\\kazua\\Documents\\C#\\wraplauncher\\outlook.ps1";
					try
					{
						string cmdStr = "";
						cmdStr = $"-File {ps1FilePath} ";
						// foreach (var item in args)
						// {
						// cmdStr = cmdStr + $"-{item.argName} {item.argValue}";
						// }
						System.Diagnostics.Process process = new System.Diagnostics.Process();
						ProcessStartInfo processStartInfo = new ProcessStartInfo("powershell.exe", cmdStr);

						processStartInfo.CreateNoWindow = true;
						processStartInfo.UseShellExecute = false;

						processStartInfo.RedirectStandardOutput = true;
						processStartInfo.RedirectStandardError = true;

						process = System.Diagnostics.Process.Start(processStartInfo);
						process.WaitForExit();

						stdOut = process.StandardOutput.ReadToEnd();
						stdErr = process.StandardError.ReadToEnd();
						exitCode = process.ExitCode;

						process.Close();

					}
					catch
					{
						throw;
					}

					MessageBox.Show("クリップボードに空きスケジュールをコピーしました。");

				}
				catch
				{
					MessageBox.Show("エラーが出ました", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
				}
			};

			return btn;
		}







		/// <summary>
		/// プログラム実行
		/// </summary>
		/// <param name="cmd">実行するコマンド</param>
		/// <returns>成否</returns>
		private bool Execute(string cmd)
		{
			try
			{
				MessageBox.Show("ボタン押下");
				ShellExecution.Run(cmd);
				return true;
			}
			catch
			{
				MessageBox.Show(
					$"起動に失敗しました。\n{cmd}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
				return false;
			}
		}

	}
}