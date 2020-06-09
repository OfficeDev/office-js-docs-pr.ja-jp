### <a name="configuration"></a><span data-ttu-id="4d50a-101">構成</span><span class="sxs-lookup"><span data-stu-id="4d50a-101">Configuration</span></span>

<span data-ttu-id="4d50a-102">次のファイルは、アドインの構成設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="4d50a-102">The following files specify configuration settings for the add-in.</span></span>

- <span data-ttu-id="4d50a-103">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="4d50a-103">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="4d50a-104">**./.ENV**プロジェクトのルートディレクトリにあるファイルには、アドインプロジェクトで使用される定数が定義されています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-104">The **./.ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>

### <a name="task-pane"></a><span data-ttu-id="4d50a-105">作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d50a-105">Task pane</span></span> 

<span data-ttu-id="4d50a-106">次のファイルは、アドインの作業ウィンドウの UI と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="4d50a-106">The following files define the add-in's task pane UI and functionality.</span></span>

- <span data-ttu-id="4d50a-107">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-107">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>

- <span data-ttu-id="4d50a-108">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-108">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>

- <span data-ttu-id="4d50a-109">JavaScript プロジェクトでは、 **/src/taskpane/taskpane.js**ファイルにアドインを初期化するコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-109">In a JavaScript project, the **./src/taskpane/taskpane.js** file contains code to initialize the add-in.</span></span> <span data-ttu-id="4d50a-110">TypeScript プロジェクトでは、/src/taskpane/taskpane.ts ファイルにアドインを初期化するコードと、Office JavaScript API ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードも記述されてい**ます**。</span><span class="sxs-lookup"><span data-stu-id="4d50a-110">In a TypeScript project, the **./src/taskpane/taskpane.ts** file contains code to initialize the add-in and also code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document.</span></span>

### <a name="authentication"></a><span data-ttu-id="4d50a-111">認証</span><span class="sxs-lookup"><span data-stu-id="4d50a-111">Authentication</span></span>

<span data-ttu-id="4d50a-112">次のファイルにより、SSO プロセスが容易になり、Office ドキュメントにデータが書き込まれます。</span><span class="sxs-lookup"><span data-stu-id="4d50a-112">The following files facilitate the SSO process and write data to the Office document.</span></span>

- <span data-ttu-id="4d50a-113">JavaScript プロジェクトの/Src/helpers/documentHelper.js ファイルには、Office JavaScript API ライブラリを使用して Microsoft Graph のデータを Office ドキュメントに追加するコードが含まれてい**ます**。</span><span class="sxs-lookup"><span data-stu-id="4d50a-113">In a JavaScript project, the **./src/helpers/documentHelper.js** file contains code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document.</span></span> <span data-ttu-id="4d50a-114">このようなファイルは TypeScript プロジェクトには含まれていません。Office JavaScript API ライブラリを使用して Microsoft Graph から Office ドキュメントにデータを追加するコードは、代わりに **/src/taskpane/taskpane.ts**にあります。</span><span class="sxs-lookup"><span data-stu-id="4d50a-114">There is no such file in a TypeScript project; the code that uses the Office JavaScript API library to add the data from Microsoft Graph to the Office document exists in **./src/taskpane/taskpane.ts** instead.</span></span>

- <span data-ttu-id="4d50a-115">**./Src/helpers/fallbackauthdialog.html**ファイルは、フォールバック認証戦略の JavaScript を読み込む UI レスページです。</span><span class="sxs-lookup"><span data-stu-id="4d50a-115">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the JavaScript for the fallback authentication strategy.</span></span>

- <span data-ttu-id="4d50a-116">**./Src/helpers/fallbackauthdialog.js**ファイルには、msal .js を使用してユーザーにサインするフォールバック認証戦略の JavaScript が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-116">The **./src/helpers/fallbackauthdialog.js** file contains the JavaScript for the fallback authentication strategy that signs in the user with msal.js.</span></span>

- <span data-ttu-id="4d50a-117">**/Src/helpers/fallbackauthhelper.js**ファイルには、SSO 認証がサポートされていないシナリオでフォールバック認証戦略を呼び出す作業ウィンドウ JavaScript が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d50a-117">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication strategy in scenarios when SSO authentication is not supported.</span></span>

- <span data-ttu-id="4d50a-118">**./src/helpers/ssoauthhelper.js** ファイルには、SSO API `getAccessToken` へのJavaScript 呼び出しが含まれ、ブートストラップ トークンの受信し、Microsoft Graph へのアクセス トークンのブートストラップ トークン交換の開始、データのための Microsoft Graph への呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="4d50a-118">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>