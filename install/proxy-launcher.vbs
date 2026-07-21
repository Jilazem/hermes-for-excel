' HermesExcel localhost proxy'yi GIZLI baslatir (konsol penceresi acmadan).
' Yalniz UZAK dagitimda (kopru baska bir makinede/sunucuda) gereklidir.
' Once .env / PROXY_UPSTREAM ayarlayin (bkz. README "Remote deployment").
' Bu dosyayi C:\HermesExcel (ya da kendi kurulum klasorunuz) altina kopyalayip
' PROXY_UPSTREAM'i asagida kendi sunucunuza gore duzenleyin, sonra Startup
' klasorune kisayol/kopya olarak ekleyin.
Dim sh, node, script, envSet
Set sh = CreateObject("WScript.Shell")

' --- BURAYI DUZENLEYIN: uzak kopru adresi ---
envSet = "set PROXY_UPSTREAM=https://<sunucu-adi-veya-ip>:8799 && "
' ---------------------------------------------

node = "node"
If CreateObject("Scripting.FileSystemObject").FileExists("C:\Program Files\nodejs\node.exe") Then
  node = """C:\Program Files\nodejs\node.exe"""
End If

' Bu VBS'in bulundugu klasoru bul, proxy.mjs'i ona gore coz (repo koku/bridge/proxy.mjs)
Dim fso, thisDir, repoRoot, scriptPath
Set fso = CreateObject("Scripting.FileSystemObject")
thisDir = fso.GetParentFolderName(WScript.ScriptFullName)
repoRoot = fso.GetParentFolderName(thisDir)   ' install/ -> repo koku
scriptPath = fso.BuildPath(repoRoot, "bridge\proxy.mjs")
If Not fso.FileExists(scriptPath) Then
  ' install/ disinda (ör. C:\HermesExcel) kopyalandiysa: proxy.mjs ayni klasorde olabilir
  scriptPath = fso.BuildPath(thisDir, "proxy.mjs")
End If

sh.Run "cmd /c " & envSet & node & " " & Chr(34) & scriptPath & Chr(34), 0, False
