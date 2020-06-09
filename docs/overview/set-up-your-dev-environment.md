---
title: 開発環境をセットアップする
description: Office アドインをビルドするための開発環境をセットアップする
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: 6c3f533b56cafc8300837cc835b26361490afedb
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611955"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="70ad8-103">開発環境をセットアップする</span><span class="sxs-lookup"><span data-stu-id="70ad8-103">Set up your development environment</span></span>

<span data-ttu-id="70ad8-104">このガイドでは、クイックスタートまたはチュートリアルに従って Office アドインを作成するためのツールのセットアップを支援します。</span><span class="sxs-lookup"><span data-stu-id="70ad8-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="70ad8-105">次の一覧からツールをインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="70ad8-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="70ad8-106">これらが既にインストールされている場合は、クイックスタートを開始する準備ができています。たとえば、この[Excel はクイックスタートを反応](../quickstarts/excel-quickstart-react.md)します。</span><span class="sxs-lookup"><span data-stu-id="70ad8-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="70ad8-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="70ad8-107">Node.js</span></span>
- <span data-ttu-id="70ad8-108">npm</span><span class="sxs-lookup"><span data-stu-id="70ad8-108">npm</span></span>
- <span data-ttu-id="70ad8-109">Office 365 (サブスクリプション版 Office) アカウント</span><span class="sxs-lookup"><span data-stu-id="70ad8-109">An Office 365 (the subscription version of Office) account</span></span>
- <span data-ttu-id="70ad8-110">任意のコードエディター</span><span class="sxs-lookup"><span data-stu-id="70ad8-110">A code editor of your choice</span></span>

<span data-ttu-id="70ad8-111">このガイドでは、コマンドラインツールの使用方法について理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="70ad8-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="70ad8-112">Node.js. のインストール</span><span class="sxs-lookup"><span data-stu-id="70ad8-112">Install Node.js</span></span>

<span data-ttu-id="70ad8-113">Node.js は JavaScript ランタイムです。モダンな Office アドインを開発する必要があります。</span><span class="sxs-lookup"><span data-stu-id="70ad8-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="70ad8-114">[Web サイトから最新の推奨バージョンをダウンロード](https://nodejs.org)して、node.js をインストールします。</span><span class="sxs-lookup"><span data-stu-id="70ad8-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="70ad8-115">オペレーティングシステムのインストール手順に従います。</span><span class="sxs-lookup"><span data-stu-id="70ad8-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="70ad8-116">Npm をインストールする</span><span class="sxs-lookup"><span data-stu-id="70ad8-116">Install npm</span></span>

<span data-ttu-id="70ad8-117">npm は、Office アドインの開発に使用されたパッケージをダウンロードするためのオープンソースソフトウェアレジストリです。</span><span class="sxs-lookup"><span data-stu-id="70ad8-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="70ad8-118">Npm をインストールするには、コマンドラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="70ad8-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="70ad8-119">既に npm がインストールされているかどうかを確認し、インストールされているバージョンを確認するには、コマンドラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="70ad8-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="70ad8-120">ノードバージョンマネージャーを使用して、node.js と npm の複数のバージョンを切り替えることができますが、これは厳密には必要ありません。</span><span class="sxs-lookup"><span data-stu-id="70ad8-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="70ad8-121">この方法の詳細については、 [「npm の手順」を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。</span><span class="sxs-lookup"><span data-stu-id="70ad8-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="70ad8-122">Office 365 を取得する</span><span class="sxs-lookup"><span data-stu-id="70ad8-122">Get Office 365</span></span>

<span data-ttu-id="70ad8-123">Office 365 アカウントをまだお持ちでない場合は、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで 90 日間の更新可能な無料の Office 365 サブスクリプションを入手できます。</span><span class="sxs-lookup"><span data-stu-id="70ad8-123">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="70ad8-124">コード エディターのインストール</span><span class="sxs-lookup"><span data-stu-id="70ad8-124">Install a code editor</span></span>

<span data-ttu-id="70ad8-125">以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。</span><span class="sxs-lookup"><span data-stu-id="70ad8-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="70ad8-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="70ad8-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="70ad8-127">Atom</span><span class="sxs-lookup"><span data-stu-id="70ad8-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="70ad8-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="70ad8-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="70ad8-129">次の手順</span><span class="sxs-lookup"><span data-stu-id="70ad8-129">Next steps</span></span>

<span data-ttu-id="70ad8-130">独自のアドインを作成するか、スクリプトラボを使用して組み込みサンプルを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="70ad8-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="70ad8-131">Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="70ad8-131">Create an Office add-in</span></span>

<span data-ttu-id="70ad8-132">[5 分間のクイック スタート](../index.md)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。</span><span class="sxs-lookup"><span data-stu-id="70ad8-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.md).</span></span> <span data-ttu-id="70ad8-133">以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.md)を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="70ad8-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.md).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="70ad8-134">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="70ad8-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="70ad8-135">Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="70ad8-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="70ad8-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="70ad8-136">See also</span></span>

- [<span data-ttu-id="70ad8-137">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="70ad8-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="70ad8-138">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="70ad8-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="70ad8-139">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="70ad8-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="70ad8-140">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="70ad8-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="70ad8-141">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="70ad8-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="70ad8-142">Office アドインの公開</span><span class="sxs-lookup"><span data-stu-id="70ad8-142">Publish Office Add-ins</span></span>](../publish/publish.md)
