<span data-ttu-id="4e431-101">開発プロジェクトを設定して、このチュートリアルを始めます。</span><span class="sxs-lookup"><span data-stu-id="4e431-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="4e431-p101">このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="4e431-p101">This page describes an individual step of a Word add-in tutorial. If you’ve arrived at this page via search engine results or other direct link, please go to the [Word add-in tutorial](../tutorials/word-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

> [!TIP]
> <span data-ttu-id="4e431-104">「[最初の Word アドインをビルドする](../quickstarts/word-quickstart.md?tabs=visual-studio-code)」をまだ読んでいない場合は、最初にその記事をご確認ください。</span><span class="sxs-lookup"><span data-stu-id="4e431-104">If you haven't already done so, please read [Build your first Word add-in](../quickstarts/word-quickstart.md?tabs=visual-studio-code).</span></span> <span data-ttu-id="4e431-105">具体的には、テスト用に Word アドインをサイドロードする方法をしっかりと理解します。</span><span class="sxs-lookup"><span data-stu-id="4e431-105">In particular, be sure that you know how to sideload a Word add-in for testing.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4e431-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="4e431-106">Prerequisites</span></span>

<span data-ttu-id="4e431-107">このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="4e431-107">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="4e431-108">Word 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。</span><span class="sxs-lookup"><span data-stu-id="4e431-108">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="4e431-109">このバージョンを入手するには、Office Insider への参加が必要になることがあります。</span><span class="sxs-lookup"><span data-stu-id="4e431-109">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="4e431-110">詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4e431-110">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>
- [<span data-ttu-id="4e431-111">Node と npm</span><span class="sxs-lookup"><span data-stu-id="4e431-111">Node and npm</span></span>](https://nodejs.org/en/) 
- <span data-ttu-id="4e431-112">[Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="4e431-112">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="setup"></a><span data-ttu-id="4e431-113">セットアップ</span><span class="sxs-lookup"><span data-stu-id="4e431-113">Setup</span></span>

1. <span data-ttu-id="4e431-114">「[Word アドインのチュートリアル](https://github.com/OfficeDev/Word-Add-in-Tutorial)」で、GitHub リポジトリを複製します。</span><span class="sxs-lookup"><span data-stu-id="4e431-114">Clone the GitHub repository [Word Add-in Tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>
2. <span data-ttu-id="4e431-115">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="4e431-115">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
3. <span data-ttu-id="4e431-116">`npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。</span><span class="sxs-lookup"><span data-stu-id="4e431-116">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 
4. <span data-ttu-id="4e431-117">「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="4e431-117">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

