' run-hidden.vbs <script.ps1>
' PowerShell スクリプトをウィンドウ非表示で起動するランチャー。
' タスクスケジューラから wscript.exe で呼ぶと、黒いコンソールが一切表示されない。
Set args = WScript.Arguments
If args.Count < 1 Then WScript.Quit 1
Set sh = CreateObject("WScript.Shell")
cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File """ & args(0) & """"
' 第2引数 0 = ウィンドウ非表示, 第3引数 False = 完了を待たない
sh.Run cmd, 0, False
