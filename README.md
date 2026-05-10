# RPA
使用するpythyon ver 3.14.4
インストーラーダウンロードページ
https://www.python.org/downloads/release/pymanager-261/

https://qiita.com/harupython/items/5aabf7fd592feca1a267

使用するライブラリ

pywinauto

https://note.com/dapper_pony2721/n/ndeebe3ed0d36

pywinautoに関するtips
https://note.com/dapper_pony2721/n/n55fdc4c63e00#4d53287a-9b00-4260-b1c9-8ac6ef3910ba

やること
clipstampのインストール
ファイル更新のためのサービス設定とその方法をまとめる


VBAよりメールアドレスを取得
Sub GetMyEmailAddress()
    Dim olApp As Object
    Dim olNS As Object
    Dim myAddress As String

    ' Outlookアプリケーションを取得
    Set olApp = CreateObject("Outlook.Application")
    ' 名前空間を取得
    Set olNS = olApp.GetNamespace("MAPI")

    ' 現在のユーザーのメールアドレスを取得
    myAddress = olNS.CurrentUser.Address

    ' 結果を表示（イミディエイトウィンドウ）
    Debug.Print "メールアドレス: " & myAddress
    MsgBox "メールアドレス: " & myAddress
End Sub

メールアドレスが取得できたら⇛確認が取れた。
jsonファイルからメールアドレスをキーとして名前とユーザネームを配置する
