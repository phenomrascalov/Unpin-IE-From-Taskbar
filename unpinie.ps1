$appnames = @("Internet Explorer")
((New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items() | 
  Where-Object { $appnames -contains $_.Name }).Verbs() | 
  Where-Object {$_.Name.replace('&','') -match 'Unpin from taskbar'} | 
  ForEach-Object {$_.DoIt()}