---
title: iPad と Mac で Office アドインをデバッグする
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 5bf626c4c18bcedccd331570b6b892a8c6a903fd
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451604"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="13e22-102">iPad と Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="13e22-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="13e22-p101">Windows でのアドインの開発とデバッグには Visual Studio を使用できますが、iPad と Mac で使用して アドインをデバッグすることはできません。アドインは HTML と Javascript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、iPad または Mac で動作するアドインをデバッグする方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="13e22-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span>

## <a name="debugging-with-vorlonjs-on-ipad-or-mac"></a><span data-ttu-id="13e22-106">iPad または Mac での Vorlon.JS を使用したデバッグ</span><span class="sxs-lookup"><span data-stu-id="13e22-106">Debugging with Vorlon.JS on iPad or Mac</span></span>

<span data-ttu-id="13e22-107">iPad または Mac でアドインをデバッグするには、Vorlon.JS (F12 ツールに似ている Web ページのデバッガー) を使用できます。</span><span class="sxs-lookup"><span data-stu-id="13e22-107">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="13e22-108">リモートで動作するように設計されているため、異なるデバイス間で Web ページをデバッグすることができます。</span><span class="sxs-lookup"><span data-stu-id="13e22-108">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="13e22-109">詳細については、[Vorlon の Web サイト](http://www.vorlonjs.com)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-109">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="13e22-110">Vorlon をインストールしてセットアップする</span><span class="sxs-lookup"><span data-stu-id="13e22-110">Install and set up Vorlon.JS</span></span>  

1.  <span data-ttu-id="13e22-111">管理者としてデバイスにログオンします。</span><span class="sxs-lookup"><span data-stu-id="13e22-111">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="13e22-112">まだ [Node.js](https://nodejs.org) をインストールしていない場合は、インストールします。</span><span class="sxs-lookup"><span data-stu-id="13e22-112">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span>

3.  <span data-ttu-id="13e22-p103">**[ターミナル]** ウィンドウを開き、コマンド `npm i -g vorlon` を入力します。ツールが `/usr/local/lib/node_modules/vorlon` にインストールされます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p103">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="13e22-115">Vorlon.JS を構成して HTTPS を使用する</span><span class="sxs-lookup"><span data-stu-id="13e22-115">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="13e22-p104">Vorlon.JS を使用してアプリケーションをデバッグするには、既知の場所から Vorlon.JS スクリプトを読み込むアプリケーションの開始ページに `<script>` タグを追加します (詳細については、次の手順を参照してください)。アドインが SSL 保護付き (HTTPS) の場合、アドインで使用するすべてのスクリプトは HTTPS サーバーからホストされるように拡張する必要があります。これには、Vorlon.JS スクリプトも含まれます。そのため、アドインで Vorlon.JS を使用するには、Vorlon.JS を構成して SSL を使用することが必要になります。</span><span class="sxs-lookup"><span data-stu-id="13e22-p104">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span>

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="13e22-119">**[Finder]** で、`/usr/local/lib/node_modules/vorlon` に移動し、`/Server` フォルダーのコンテキスト メニュー (右クリック) を開き、**[情報を見る]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="13e22-119">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="13e22-120">**[サーバー情報]** ウィンドウの右下隅にある南京錠アイコンを選択して、フォルダーのロックを解除します。</span><span class="sxs-lookup"><span data-stu-id="13e22-120">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="13e22-121">ウィンドウの **[共有とアクセス権]** セクションで、**スタッフ** グループの **[特権]** を **[読み取り/書き込み]** に設定します。</span><span class="sxs-lookup"><span data-stu-id="13e22-121">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="13e22-122">南京錠アイコンをもう一度選択して、フォルダーを***再度ロック***します。</span><span class="sxs-lookup"><span data-stu-id="13e22-122">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="13e22-123">**[Finder]** に戻り、`/Server` サブフォルダーを展開し、ファイル `config.json` を右クリックして、**[情報を見る]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="13e22-123">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="13e22-p105">**[config.json 情報]** ウィンドウで、親 `/Server` フォルダーに対して行ったものと同じ方法でファイルの特権を変更します。必ず再度ロックしてからウィンドウを閉じてください。</span><span class="sxs-lookup"><span data-stu-id="13e22-p105">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="13e22-p106">**[Finder]** に戻り、ファイル `config.json` を右クリックして、**[このアプリケーションで開く]**、**[テキストエディット]** の順に選択します。ファイルがテキスト エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p106">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="13e22-128">**useSSL** プロパティの値を `true` に変更します。</span><span class="sxs-lookup"><span data-stu-id="13e22-128">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="13e22-p107">**[プラグイン]** セクションで、**ID** が `OFFICE` で**名前**が `Office Addin` のプライグインを検索します。プラグインの **enabled** プロパティがまだ `true` になっていない場合は、`true` に設定します。</span><span class="sxs-lookup"><span data-stu-id="13e22-p107">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="13e22-131">ファイルを保存し、エディターを閉じます。</span><span class="sxs-lookup"><span data-stu-id="13e22-131">Save the file and close the editor.</span></span>

11. <span data-ttu-id="13e22-132">**[検索]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="13e22-132">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span>

12. <span data-ttu-id="13e22-p108">**[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。管理者パスワードの入力を求めるダイアログ ボックスが表示されます。Vorlon サーバーが起動します。**[ターミナル]** ウィンドウを開いたままにしておきます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p108">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="13e22-p109">ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。ダイアログ ボックスが表示されたら、**[常に]** を選択して、セキュリティ証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="13e22-p109">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span>

    > [!NOTE]
    > <span data-ttu-id="13e22-p110">ダイアログ ボックスが表示されない場合は、手動で証明書を信頼する必要があります。証明書ファイルは `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt` です。次の手順を実行し、問題が発生した場合は、Macintosh または iPad のヘルプを参照してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-p110">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span>
    >
    > 1. <span data-ttu-id="13e22-143">ブラウザー ウィンドウを閉じ、Vorlon サーバーを実行している **[ターミナル]** ウィンドウで、Control-C を使用してサーバーを停止します。</span><span class="sxs-lookup"><span data-stu-id="13e22-143">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="13e22-p111">**[Finder]** で、`server.crt` ファイルを右クリックして、**[キーチェーンアクセス]** を選択します。**[キーチェーンアクセス]** ウィンドウが開きます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p111">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="13e22-p112">左側の **[キーチェーン]** リストで、**[ログイン]** がまだ選択されていない場合は選択し、**[カテゴリ]** セクションで **[証明書]** を選択します。証明書 **localhost** が一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p112">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="13e22-p113">証明書 **localhost** を右クリックし、**[情報を見る]** を選択します。**[localhost]** ウィンドウが開きます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p113">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="13e22-150">**[信頼]** セクションで、**[この証明書を使用する場合]** というラベルの付いたセレクターを選択し、**[常に信頼する]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="13e22-150">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="13e22-p114">**[localhost]** ウィンドウを閉じます。アクションが成功すると、**[キーチェーンアクセス]** ウィンドウの **localhost** 証明書のアイコンに青い円で囲まれた白い十字が表示されます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p114">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="13e22-153">Vorlon.JS デバッグ用のアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="13e22-153">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="13e22-154">次のスクリプト タグを、アドインの home.html ファイル (またはメイン HTML ファイル) の `<head>` セクションに追加します。</span><span class="sxs-lookup"><span data-stu-id="13e22-154">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. <span data-ttu-id="13e22-155">Azure Web サイトなど、Mac または iPad からアクセス可能な Web サーバーにアドイン Web アプリケーションを展開します。</span><span class="sxs-lookup"><span data-stu-id="13e22-155">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span>

3. <span data-ttu-id="13e22-156">アドイン マニフェストに URL が表示されるすべての場所で、アドインの URL を更新します。</span><span class="sxs-lookup"><span data-stu-id="13e22-156">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="13e22-157">アドイン マニフェストを Mac または iPad 上の次のフォルダーにコピーします: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`。ここで、*{host_name}* は、Word、Excel、PowerPoint、または Outlook です。</span><span class="sxs-lookup"><span data-stu-id="13e22-157">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="13e22-158">Vorlon.JS でアドインを検査する</span><span class="sxs-lookup"><span data-stu-id="13e22-158">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="13e22-159">Vorlon サーバーが実行されていない場合、**[Finder]** で `/usr/local/lib/node_modules/vorlon` に移動して、`Server` サブフォルダーを右クリックし、**[フォルダーの新しいターミナル]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="13e22-159">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 

2.  <span data-ttu-id="13e22-p115">**[ターミナル]** ウィンドウで、`sudo vorlon` と入力します。管理者パスワードの入力を求めるダイアログ ボックスが表示されます。Vorlon サーバーが起動します。**[ターミナル]** ウィンドウを開いたままにしておきます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p115">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="13e22-164">ブラウザー ウィンドウを開き、Vorlon.JS インターフェイスの `https://localhost:1337` に進みます。</span><span class="sxs-lookup"><span data-stu-id="13e22-164">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="13e22-165">アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="13e22-165">Sideload the add-in.</span></span> <span data-ttu-id="13e22-166">アドインが Excel、PowerPoint、Word 用の場合は、「[iPad または Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)」の説明に従ってサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="13e22-166">If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="13e22-167">アドインが Outlook アドインである場合は、「[テストのために Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」の説明に従ってサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="13e22-167">If it is an Outlook add-in, sideload it as described in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span> <span data-ttu-id="13e22-168">アドインでアドイン コマンドを使用しない場合は、アドインが直ちに開きます。</span><span class="sxs-lookup"><span data-stu-id="13e22-168">If the add-in does not use add-in commands, it will open immediately.</span></span> <span data-ttu-id="13e22-169">それ以外の場合は、ボタンを選択してアドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="13e22-169">Otherwise, choose the button to open the add-in.</span></span> <span data-ttu-id="13e22-170">Office ホスト アプリケーションのビルドに応じて、ボタンは **[ホーム]** タブまたは **[アドイン]** タブのいずれかに表示されます。</span><span class="sxs-lookup"><span data-stu-id="13e22-170">Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="13e22-171">アドインは、Vorlon.JS のクライアントのリスト (Vorlon.JS インターフェイスの左側) に **{OS} - n** として表示されます。*n* は数値、*{OS}* は "Macintosh" などのデバイスの種類です。</span><span class="sxs-lookup"><span data-stu-id="13e22-171">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span>

![Vorlon.js インターフェイスを示すスクリーンショット](../images/vorlon-interface.png)

<span data-ttu-id="13e22-173">Vorlon ツールには、さまざまなプラグインがあります。現在有効になっているプラグインはツールの上部にタブとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="13e22-173">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool.</span></span> <span data-ttu-id="13e22-174">(左側にある歯車アイコンを選択すると、さらに別のプラグインを有効にすることができます。)これらのプラグインは、F12 ツールの機能に似ています。</span><span class="sxs-lookup"><span data-stu-id="13e22-174">(You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools.</span></span> <span data-ttu-id="13e22-175">たとえば、DOM 要素の強調表示、コマンドの実行などを行えます。</span><span class="sxs-lookup"><span data-stu-id="13e22-175">For example, you can highlight DOM elements, execute commands, and more.</span></span> <span data-ttu-id="13e22-176">詳細については、[Vorlon ドキュメントの「コア プラグイン」](http://vorlonjs.com/documentation/#console)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-176">For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console).</span></span>

<span data-ttu-id="13e22-p118">**Office アドイン** プラグインにより Office.js に特別な機能 (オブジェクト モデルを調査する機能、Office.js の呼び出しを実行する機能、およびオブジェクト プロパティの値を読み取る機能など) が追加されます。手順については、「[Office アドインをデバッグするための VorlonJS プラグイン](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-p118">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="13e22-179">Vorlon.JS にブレーク ポイントを設定する方法はありません。</span><span class="sxs-lookup"><span data-stu-id="13e22-179">There is no way to set break points in Vorlon.JS.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="13e22-180">Mac での Safari Web インスペクタを使用したデバッグ</span><span class="sxs-lookup"><span data-stu-id="13e22-180">Debugging with Safari Web Inspector on a Mac</span></span>

> [!IMPORTANT]
> <span data-ttu-id="13e22-181">**要素の検査**アドイン コンテキスト メニュー オプションは試験的な機能であり、Office アプリケーションの将来のバージョンでこの機能が維持されるかどうかは保証されない点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-181">Please note that the **Inspect Element** add-in context menu option is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

<span data-ttu-id="13e22-182">作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="13e22-182">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="13e22-183">Mac の Office アドインをデバッグするには、Mac OS High Sierra 以降 と Mac Office バージョン 16.9.1 (ビルド 18012504) 以降の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="13e22-183">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra or later AND Mac Office version 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="13e22-184">Office for Mac ビルドをまだお持ちでない場合は、[Office 365 Developer Program](https://aka.ms/o365devprogram) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="13e22-184">If you don't have an Office for Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="13e22-185">最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。</span><span class="sxs-lookup"><span data-stu-id="13e22-185">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="13e22-186">次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="13e22-186">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="13e22-187">アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="13e22-187">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="13e22-188">このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="13e22-188">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="13e22-189">インスペクタを使用するとダイアログのちらつきが発生する場合は、次の回避策を試してください。</span><span class="sxs-lookup"><span data-stu-id="13e22-189">If you are trying to use the inspector and the dialog flickers, try the following workaround:</span></span>
> 1. <span data-ttu-id="13e22-190">ダイアログのサイズを変更します。</span><span class="sxs-lookup"><span data-stu-id="13e22-190">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="13e22-191">**[要素の検査]** を選択します (新しいウィンドウが開きます)。</span><span class="sxs-lookup"><span data-stu-id="13e22-191">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="13e22-192">ダイアログを元のサイズに変更します。</span><span class="sxs-lookup"><span data-stu-id="13e22-192">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="13e22-193">必要に応じてインスペクタを使用します。</span><span class="sxs-lookup"><span data-stu-id="13e22-193">Use the inspector as required.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="13e22-194">Mac または iPad 上の Office アプリケーションのキャッシュのクリア</span><span class="sxs-lookup"><span data-stu-id="13e22-194">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="13e22-p121">アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="13e22-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="13e22-198">Mac では、`/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダー内にあるすべてを削除することによってキャッシュを手動でクリアできます。</span><span class="sxs-lookup"><span data-stu-id="13e22-198">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

<span data-ttu-id="13e22-p122">iPad では、アドインの JavaScript から `window.location.reload(true)` を呼び出して、強制的に再読み込みすることができます。または、Office を再インストールすることができます。</span><span class="sxs-lookup"><span data-stu-id="13e22-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
