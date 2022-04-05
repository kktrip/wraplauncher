# �N���ς݂�Outlook�����邩�m�F
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
$needQuit = $false
if ($outlookProcess -eq $null) {
    # Write-Host -Object "Outlook�N�����Ă��Ȃ�" -ForegroundColor Red
    $needQuit = $true
}

$app = New-Object -ComObject Outlook.Application

# �v���t�@�C����I�����ă��O�I��
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

    #������10�F00�ȑO�̂��̂͑O�����Ƃ���Restrict���\�b�h�ŏ�������Ă���H�悤�Ȃ̂őO�����J�n�����ɐݒ�
    $Start = $date.AddDays(-1).ToShortDateString()
    #���Ƃ�if���œ����������������ʂ��邽�߂Ɏg�p
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
            #����I�ȃA�C�e���ł���΁A�j������v���Ă���Ώo�͌��ʂɊ܂�
            $teiki_flg = 1
        }
        if ( ($teiki_flg -eq 1) -or ( ($_.Start -gt $Start2) -and ($_.End -lt $End) )  ) {
            $curDate = $_.Start.ToString("yyyy/MM/dd")
            $curStartTime = $_.Start.ToString("HH:mm:ss")
            $curEndTime = $_.End.ToString("HH:mm:ss")
            if ($curDate -ne $befDate) {
                # �O�̍s�ɂ����āA�I�Ƃ܂ł̋󂫎��Ԃ��o�͂ɒǉ�
                if ($befEndTime -lt $schEndTime) {
                    $output += $befEndTime + " �` " + $schEndTime + "`r`n"
                }
                # ���t���ς��Ƃ�
                $befDate = $curDate
                # �o�͂ɓ��t�ǉ�
                $output += "`r`n" + $curDate + "(" + $_.Start.ToString("ddd") + ")" + "`r`n"
                # ������
                $befEndTime = $schStartTime
            }

            # �󂢂Ă��鎞�Ԃ�T���@�O�̃^�X�N�I������ < �^�X�N�J�n����
            # Write-Host -Object ($curDate + " | Subject:" + $_.Subject + " | befEndTime:" + $befEndTime + " | curStartTime:" + $curStartTime + " | curEndTime:" + $curEndTime) -ForegroundColor Yellow
            if ($befEndTime -lt $curStartTime) {
                # ���̃^�X�N���玟�̃^�X�N�܂ł̎��Ԃ͋󂢂Ă���
                # Write-Host -Object ("Subject:" + $_.Subject + " | befEndTime -lt curStartTime") -ForegroundColor Green
                $output += $befEndTime + " �` " + $curStartTime + "`r`n"
            }

            if ($i -eq $items.Count - 1) {
                # �ŏI�s
                if ($curEndTime -lt $schEndTime) {
                    # �^�X�N�I�������`�I�Ǝ����܂ŋ󂫎���
                    $output += $curEndTime + " �` " + $schEndTime + "`r`n"
                }
            }

            # if ($schEndTime -lt $curEndTime) {
            # ���̓��Ō�̃^�X�N�����肵�A�I�Ƃ܂ŋ󂫎��ԂȂ̂Œʒm����K�v����
            # $output += $befEndTime + " �` " + $curStartTime + "`r`n"
            # }
            # Write-Host -Object ("Subject:" + $_.Subject + " | befEndTime:" + $befEndTime) -ForegroundColor White
            $befEndTime = $curEndTime
            #����I�ȃA�C�e���łȂ��ꍇ�́A�J�n�����ƏI���������Ĕ���
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
    # $mail.Subject = "�^�C�g���ł��킷" 
    # $mail.Body = "�{���ł��킷`n" + $result
    # $mail.Send()

    # ����M���s��
    # ���̃��\�b�h�͔񓯊��Ȃ̂ŁA10�b�قǃX���[�v���Ă���
    # https://docs.microsoft.com/ja-jp/office/vba/api/outlook.namespace.sendandreceive
    # $session = $app.Session
    # $session.SendAndReceive($False)
    # Start-Sleep 10

    $namespace.Logoff()

    # ���ǂ�������񏈗� ==================================================================
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
    # ���ǂ�������񏈗� ==================================================================
}
finally {
    # Write-Host -Object "finally�ɑJ��" -ForegroundColor Red
    if ($needQuit) {
        [void]$outlook.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
    }
}

