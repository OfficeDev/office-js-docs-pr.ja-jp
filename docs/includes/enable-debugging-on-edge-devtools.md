<span data-ttu-id="77924-101">Microsoft Edge でアドインを実行している場合、UI のないコードは既定でデバッガーにアタッチできません。</span><span class="sxs-lookup"><span data-stu-id="77924-101">When the add-in is running in Microsoft Edge, UI-less code will not be able to attach to a debugger by default.</span></span>
<span data-ttu-id="77924-102">UI のないコードは、アドイン コマンドなどの作業ウィンドウが表示されていないときに実行されます。</span><span class="sxs-lookup"><span data-stu-id="77924-102">UI-less code is any code running while the task pane is not visible, such as add-in commands.</span></span> <span data-ttu-id="77924-103">デバッグを有効にするには、[Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) の次のコマンドを実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="77924-103">To enable debugging, you need to run the following [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) commands.</span></span>

1. <span data-ttu-id="77924-104">次のコマンドを実行して、**Microsoft.Win32WebViewHost** アプリ パッケージの情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="77924-104">Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.</span></span>
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    <span data-ttu-id="77924-105">このコマンドを実行すると、次のようなアプリ パッケージ情報が表示されます。</span><span class="sxs-lookup"><span data-stu-id="77924-105">The command lists app package information similar to the following output.</span></span>
    
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
    
2. <span data-ttu-id="77924-106">次のコマンドを実行してデバッグを有効にします。</span><span class="sxs-lookup"><span data-stu-id="77924-106">Run the following command to enable debugging.</span></span> <span data-ttu-id="77924-107">前のコマンドに記載されている **PackageFullName** の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="77924-107">Use the value for the **PackageFullName** listed from the previous command.</span></span>
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. <span data-ttu-id="77924-108">Office が既に実行されている場合は、Office を終了して再起動すると、デバッグの変更が反映されます。</span><span class="sxs-lookup"><span data-stu-id="77924-108">If Office was already running, close and restart Office so that it picks up the debugging change.</span></span>