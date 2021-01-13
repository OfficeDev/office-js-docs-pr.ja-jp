---
title: 開発環境をセットアップする
description: カスタム アドインを構築するためのOfficeを設定します。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: eddf8bdf7b20a54667e6f8eb38bdace801ea1813
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839713"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="1f3e6-103">開発環境をセットアップする</span><span class="sxs-lookup"><span data-stu-id="1f3e6-103">Set up your development environment</span></span>

<span data-ttu-id="1f3e6-104">このガイドは、クイック スタートまたはチュートリアルに従Officeアドインを作成するためのツールをセットアップする場合に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="1f3e6-105">次の一覧からツールをインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="1f3e6-106">既にインストールされている場合は、この Excel React クイック スタートなど、クイック スタートを [開始する準備が整っています](../quickstarts/excel-quickstart-react.md)。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="1f3e6-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="1f3e6-107">Node.js</span></span>
- <span data-ttu-id="1f3e6-108">npm</span><span class="sxs-lookup"><span data-stu-id="1f3e6-108">npm</span></span>
- <span data-ttu-id="1f3e6-109">Microsoft 365 アカウント。サブスクリプション バージョンのアカウントが含Office</span><span class="sxs-lookup"><span data-stu-id="1f3e6-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="1f3e6-110">選択したコード エディター</span><span class="sxs-lookup"><span data-stu-id="1f3e6-110">A code editor of your choice</span></span>

<span data-ttu-id="1f3e6-111">このガイドでは、コマンド ライン ツールの使い方を知っている必要があります。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="1f3e6-112">Node.js. のインストール</span><span class="sxs-lookup"><span data-stu-id="1f3e6-112">Install Node.js</span></span>

<span data-ttu-id="1f3e6-113">Node.jsは、モダン アドインを開発するために必要Office JavaScript ランタイムです。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="1f3e6-114">Web Node.js [から最新の推奨バージョンをダウンロードしてインストールします](https://nodejs.org)。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="1f3e6-115">オペレーティング システムのインストール手順に従います。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="1f3e6-116">npm をインストールする</span><span class="sxs-lookup"><span data-stu-id="1f3e6-116">Install npm</span></span>

<span data-ttu-id="1f3e6-117">npm はオープン ソース のソフトウェア レジストリで、アドインの開発に使用されるパッケージOfficeダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="1f3e6-118">npm をインストールするには、コマンド ラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="1f3e6-119">npm が既にインストール済みで、インストールされているバージョンを確認するには、コマンド ラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="1f3e6-120">ノード バージョン マネージャーを使用して、複数のバージョンの Node.js と npm を切り替える場合がありますが、これは厳密には必要ありません。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="1f3e6-121">これを行う方法の詳細については [、npm の手順を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="1f3e6-122">Get Office 365</span><span class="sxs-lookup"><span data-stu-id="1f3e6-122">Get Office 365</span></span>

<span data-ttu-id="1f3e6-123">Microsoft 365 アカウントをまだお持ちでない場合は、[Microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで 90 日間の更新可能な無料の Microsoft 365 サブスクリプションを入手できます。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="1f3e6-124">コード エディターのインストール</span><span class="sxs-lookup"><span data-stu-id="1f3e6-124">Install a code editor</span></span>

<span data-ttu-id="1f3e6-125">以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="1f3e6-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="1f3e6-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="1f3e6-127">Atom</span><span class="sxs-lookup"><span data-stu-id="1f3e6-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="1f3e6-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="1f3e6-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="1f3e6-129">次の手順</span><span class="sxs-lookup"><span data-stu-id="1f3e6-129">Next steps</span></span>

<span data-ttu-id="1f3e6-130">独自のアドインを作成するか、Script Lab を使用して組み込みのサンプルを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="1f3e6-131">Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="1f3e6-131">Create an Office add-in</span></span>

<span data-ttu-id="1f3e6-132">[5 分間のクイック スタート](../index.yml)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml).</span></span> <span data-ttu-id="1f3e6-133">以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.yml)を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="1f3e6-134">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="1f3e6-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="1f3e6-135">Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1f3e6-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="1f3e6-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="1f3e6-136">See also</span></span>

- [<span data-ttu-id="1f3e6-137">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="1f3e6-137">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="1f3e6-138">アドインOffice開発する</span><span class="sxs-lookup"><span data-stu-id="1f3e6-138">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="1f3e6-139">Office アドインの設計</span><span class="sxs-lookup"><span data-stu-id="1f3e6-139">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1f3e6-140">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="1f3e6-140">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="1f3e6-141">Office アドインの公開</span><span class="sxs-lookup"><span data-stu-id="1f3e6-141">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="1f3e6-142">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="1f3e6-142">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)