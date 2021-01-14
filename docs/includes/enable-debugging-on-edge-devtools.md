Microsoft Edge でアドインを実行している場合、UI のないコードは既定でデバッガーにアタッチできません。
UI のないコードは、アドイン コマンドなどの作業ウィンドウが表示されていないときに実行されます。 デバッグを有効にするには、[Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) の次のコマンドを実行する必要があります。

1. 次のコマンドを実行して、**Microsoft.Win32WebViewHost** アプリ パッケージの情報を取得します。
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    このコマンドを実行すると、次のようなアプリ パッケージ情報が表示されます。
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. 次のコマンドを実行してデバッグを有効にします。 前のコマンドに記載されている **PackageFullName** の値を使用します。
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. Office が既に実行されている場合は、Office を終了して再起動すると、デバッグの変更が反映されます。