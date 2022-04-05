# 起動済みのOutlookがあるか確認
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
$needQuit = $false
if ($outlookProcess -eq $null) {
    # Write-Host -Object "Outlook起動していない" -ForegroundColor Red
    $needQuit = $true
}

$app = New-Object -ComObject Outlook.Application

# プロファイルを選択してログオン
try {
    $namespace = $app.GetNamespace("MAPI")
    $namespace.Logon("Outlook")

    [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Outlook")

    $OlDefaultFolders = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]
    # $OlItemType = [Microsoft.Office.Interop.Outlook.OlItemType]
    # $OlBusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]

    $date = Get-Date
    $schStartTime = "09:00:00"
    $schEndTime = "18:00:00"

    #当日の10：00以前のものは前日分としてRestrictメソッドで処理されている？ようなので前日を開始時刻に設定
    $Start = $date.AddDays(-1).ToShortDateString()
    #あとでif文で当日分をきっちり区別するために使用
    $Start2 = $date.AddDays(0).ToShortDateString()
    $End = $date.AddDays(14).ToShortDateString()
    $Filter = "[Start] >= '$Start' AND [End] < '$End'"

    $folder = $namespace.GetDefaultFolder($OlDefaultFolders::olFolderCalendar)
    $allItems = $folder.Items
    $allItems.Sort("[Start]")
    $allItems.IncludeRecurrences = $true

    $items = $allItems.Restrict($filter)

    $befDate = $date.ToString("yyyy/MM/dd")
    $befEndTime = $schStartTime

    $output = $befDate + "(" + $date.ToString("ddd") + ")`r`n"

    $i = 0
    $items | Sort-Object Start | ForEach-Object {
        # Write-Host -Object ($i.ToString() + " | " + ($items.Count - 1).ToString() + " | curDate:" + $_.Start.ToString("yyyy/MM/dd") + " | subject:" + $_.Subject) -ForegroundColor Green
        if ( ( $_.RecurrenceState -eq 1) -and ($_.Start.DayOfWeek -eq $date.DayOfWeek) ) {
            #定期的なアイテムであれば、曜日が一致していれば出力結果に含む
            $teiki_flg = 1
        }
        if ( ($teiki_flg -eq 1) -or ( ($_.Start -gt $Start2) -and ($_.End -lt $End) )  ) {
            $curDate = $_.Start.ToString("yyyy/MM/dd")
            $curStartTime = $_.Start.ToString("HH:mm:ss")
            $curEndTime = $_.End.ToString("HH:mm:ss")
            if ($curDate -ne $befDate) {
                # 前の行において、終業までの空き時間を出力に追加
                if ($befEndTime -lt $schEndTime) {
                    $output += $befEndTime + " ～ " + $schEndTime + "`r`n"
                }
                # 日付が変わるとき
                $befDate = $curDate
                # 出力に日付追加
                $output += "`r`n" + $curDate + "(" + $_.Start.ToString("ddd") + ")" + "`r`n"
                # 初期化
                $befEndTime = $schStartTime
            }

            # 空いている時間を探す　前のタスク終了時刻 < タスク開始時刻
            # Write-Host -Object ($curDate + " | Subject:" + $_.Subject + " | befEndTime:" + $befEndTime + " | curStartTime:" + $curStartTime + " | curEndTime:" + $curEndTime) -ForegroundColor Yellow
            if ($befEndTime -lt $curStartTime) {
                # 今のタスクから次のタスクまでの時間は空いている
                # Write-Host -Object ("Subject:" + $_.Subject + " | befEndTime -lt curStartTime") -ForegroundColor Green
                $output += $befEndTime + " ～ " + $curStartTime + "`r`n"
            }

            if ($i -eq $items.Count - 1) {
                # 最終行
                if ($curEndTime -lt $schEndTime) {
                    # タスク終了時刻～終業時刻まで空き時間
                    $output += $curEndTime + " ～ " + $schEndTime + "`r`n"
                }
            }

            # if ($schEndTime -lt $curEndTime) {
            # その日最後のタスクか判定し、終業まで空き時間なので通知する必要あり
            # $output += $befEndTime + " ～ " + $curStartTime + "`r`n"
            # }
            # Write-Host -Object ("Subject:" + $_.Subject + " | befEndTime:" + $befEndTime) -ForegroundColor White
            $befEndTime = $curEndTime
            #定期的なアイテムでない場合は、開始時刻と終了時刻を再判定
            # $start_ar = $_.Start -split " "
            # $end_ar = $_.End -split " "
            # $start_t = $start_ar[1] -split ":"
            # $end_t = $end_ar[1] -split ":"
            # $start_h = $start_t[0]
            # $start_m = $start_t[1]
            # $end_h = $end_t[0]
            # $end_m = $end_t[1]
            # $output += "--------------------`r`n"
            # $output += $_.Start.ToString("yyyy/MM/dd") + "(" + $_.Start.ToString("ddd") + ") " + $_.Start.ToString("HH:mm:ss") + " "`
            #     + $_.End.ToString("yyyy/MM/dd") + "(" + $_.End.ToString("ddd") + ") " + $_.End.ToString("HH:mm:ss") + "`r`n"
            # $output += $_.Subject, "`r`n"
            # $output += $start_h + ":" + $start_m + "-" + $end_h + ":" + $end_m + "`r`n"
            # $output += $_.Location, "`r`n"
        }
        $i += 1
    }
    $OutputEncoding = [console]::OutputEncoding;
    # $output += "--------------------"
    $output
    $output | clip

    # foreach ($folder in $folders) {
    #     Write-Host -Object $folder.Name -ForegroundColor Red
    #     $folders = $folder.Folders
    #     $result = $result + $folder.Name + "`n"
    # }
    # $result = ""

    # https://docs.microsoft.com/ja-jp/office/vba/api/outlook.olitemtype
    # 0:olmailitem
    # $mail = $app.CreateItem(0)
    # $mail.To = "kazuaki.kiyomi@falcs.jp" 
    # $mail.Subject = "タイトルでごわす" 
    # $mail.Body = "本文でごわす`n" + $result
    # $mail.Send()

    # 送受信を行う
    # このメソッドは非同期なので、10秒ほどスリープしている
    # https://docs.microsoft.com/ja-jp/office/vba/api/outlook.namespace.sendandreceive
    # $session = $app.Session
    # $session.SendAndReceive($False)
    # Start-Sleep 10

    $namespace.Logoff()

    # ↓良く分からん処理 ==================================================================
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($attachments) | Out-Null
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Mail) | Out-Null
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($session) | Out-Null


    # $namespace = $null
    # Remove-Variable namespace -ErrorAction SilentlyContinue
    # $attachments = $null
    # Remove-Variable attachments -ErrorAction SilentlyContinue
    # $mail = $null
    # Remove-Variable mail -ErrorAction SilentlyContinue
    # $namespace = $null
    # Remove-Variable namespace -ErrorAction SilentlyContinue
    # $session = $null
    # Remove-Variable session -ErrorAction SilentlyContinue

    # [System.GC]::Collect()
    # [System.GC]::WaitForPendingFinalizers()
    # [System.GC]::Collect()

    # $app.Quit() 


    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    # $app = $null
    # Remove-Variable app -ErrorAction SilentlyContinue

    # [System.GC]::Collect()
    # [System.GC]::WaitForPendingFinalizers()
    # [System.GC]::Collect()
    # ↑良く分からん処理 ==================================================================
}
finally {
    # Write-Host -Object "finallyに遷移" -ForegroundColor Red
    if ($needQuit) {
        [void]$outlook.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
    }
}

