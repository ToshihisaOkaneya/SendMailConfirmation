# SendMailConfirmation
Show confirmation message before you send Outlook mail

Outlookでのメール送信時に確認メッセージを表示する。  
Alt+sでメールを作成途中で送信することを防ぐためにAlt+sを無効化しようとしたが、  
設定方法が見つからなかったので作成。  

まずOutlookでメニューから以下のように設定することでマクロを有効にする。  
ファイル  
→オプション  
→セキュリティセンター  
→セキュリティセンターの設定  
→マクロの設定  
→マクロの設定を「すべてのマクロに対して警告を表示する」に変更  

次にAlt+F11押下でマクロ編集画面を表示し、左上のプロジェクトウィンドウから  
Project1  
→Microsoft Outlook Objects  
→ThisOutlookSession  
を選択して編集画面を表示し、このリポジトリのThisOutlookSessionファイルの  
中身を貼り付けて使用する。  

ただOutlook起動時に毎回警告が表示される。  
宛先を書く前に本文を書くように習慣づければ良いだけかもしれない。。 
