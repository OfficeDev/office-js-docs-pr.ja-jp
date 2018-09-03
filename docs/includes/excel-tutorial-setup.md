<span data-ttu-id="c13cf-101">開発プロジェクトを設定して、このチュートリアルを始めます。</span><span class="sxs-lookup"><span data-stu-id="c13cf-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="c13cf-102">このページでは、Excel のアドインのチュートリアルの個々 の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="c13cf-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="c13cf-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[Excel アドインのチュートリアル](../tutorials/excel-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="c13cf-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c13cf-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="c13cf-104">Prerequisites</span></span>

<span data-ttu-id="c13cf-105">このチュートリアルを使用するには、以下のバージョンがインストールされている必要があります。</span><span class="sxs-lookup"><span data-stu-id="c13cf-105">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="c13cf-106">Excel 2016、バージョン 1711 (ビルド 8730.1000 クイック実行) 以降。</span><span class="sxs-lookup"><span data-stu-id="c13cf-106">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="c13cf-107">このバージョンを入手するには、Office Insider への参加が必要になることがあります。</span><span class="sxs-lookup"><span data-stu-id="c13cf-107">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="c13cf-108">詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c13cf-108">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>
- [<span data-ttu-id="c13cf-109">Node と npm</span><span class="sxs-lookup"><span data-stu-id="c13cf-109">Node and npm</span></span>](https://nodejs.org/en/) 
- <span data-ttu-id="c13cf-110">[Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="c13cf-110">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="setup"></a><span data-ttu-id="c13cf-111">セットアップ</span><span class="sxs-lookup"><span data-stu-id="c13cf-111">Setup</span></span>

1. <span data-ttu-id="c13cf-112">「[Excel アドインのチュートリアル](https://github.com/OfficeDev/Excel-Add-in-Tutorial)」で、GitHub リポジトリを複製します。</span><span class="sxs-lookup"><span data-stu-id="c13cf-112">Clone the GitHub repository [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>
2. <span data-ttu-id="c13cf-113">Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="c13cf-113">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
3. <span data-ttu-id="c13cf-114">`npm install` コマンドを実行して、package.json ファイルに一覧表示されているツールとライブラリをインストールします。</span><span class="sxs-lookup"><span data-stu-id="c13cf-114">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 
4. <span data-ttu-id="c13cf-115">「[自己署名証明書を信頼されたルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」の手順を実行して、開発用コンピューターのオペレーティング システムの証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="c13cf-115">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

