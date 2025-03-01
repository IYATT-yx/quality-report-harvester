$startTime = Get-Date

python -m venv venv
.\venv\Scripts\activate
pip install -r .\requirements.txt
pip install nuitka==2.6.4

python .\savebuildtime.py

nuitka --standalone --onefile --remove-output --windows-console-mode=disable `
--enable-plugin=tk-inter `
--windows-icon-from-ico=.\icon.ico --include-data-file=.\icon.ico=.\ `
--include-data-file=.\venv\Lib\site-packages\jieba\dict.txt=.\jieba\ `
--output-dir=dist --output-filename=quality-report-harvester_win_amd64 `
.\quality-report-harvester.py

$endTime = Get-Date
$elapsedTime = New-TimeSpan -Start $startTime -End $endTime
Write-Output "程序构建用时：$($elapsedTime.TotalSeconds) 秒"