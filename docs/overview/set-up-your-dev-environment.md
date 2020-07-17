---
title: 開発環境をセットアップする
description: Office アドインをビルドするための開発環境をセットアップする
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 1948cd83a252ea713c9b9a41941ceaef09d4a524
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159410"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="06ea4-103">開発環境をセットアップする</span><span class="sxs-lookup"><span data-stu-id="06ea4-103">Set up your development environment</span></span>

<span data-ttu-id="06ea4-104">このガイドでは、クイックスタートまたはチュートリアルに従って Office アドインを作成するためのツールのセットアップを支援します。</span><span class="sxs-lookup"><span data-stu-id="06ea4-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="06ea4-105">次の一覧からツールをインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="06ea4-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="06ea4-106">これらが既にインストールされている場合は、クイックスタートを開始する準備ができています。たとえば、この[Excel はクイックスタートを反応](../quickstarts/excel-quickstart-react.md)します。</span><span class="sxs-lookup"><span data-stu-id="06ea4-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="06ea4-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="06ea4-107">Node.js</span></span>
- <span data-ttu-id="06ea4-108">npm</span><span class="sxs-lookup"><span data-stu-id="06ea4-108">npm</span></span>
- <span data-ttu-id="06ea4-109">Office のサブスクリプション版を含む Microsoft 365 アカウント</span><span class="sxs-lookup"><span data-stu-id="06ea4-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="06ea4-110">任意のコードエディター</span><span class="sxs-lookup"><span data-stu-id="06ea4-110">A code editor of your choice</span></span>

<span data-ttu-id="06ea4-111">このガイドでは、コマンドラインツールの使用方法について理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="06ea4-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="06ea4-112">Node.js. のインストール</span><span class="sxs-lookup"><span data-stu-id="06ea4-112">Install Node.js</span></span>

<span data-ttu-id="06ea4-113">Node.js は JavaScript ランタイムで、モダン Office アドインを開発する必要があります。</span><span class="sxs-lookup"><span data-stu-id="06ea4-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="06ea4-114">[Web サイトから最新の推奨バージョンをダウンロード](https://nodejs.org)して、Node.js をインストールします。</span><span class="sxs-lookup"><span data-stu-id="06ea4-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="06ea4-115">オペレーティングシステムのインストール手順に従います。</span><span class="sxs-lookup"><span data-stu-id="06ea4-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="06ea4-116">Npm をインストールする</span><span class="sxs-lookup"><span data-stu-id="06ea4-116">Install npm</span></span>

<span data-ttu-id="06ea4-117">npm は、Office アドインの開発に使用されたパッケージをダウンロードするためのオープンソースソフトウェアレジストリです。</span><span class="sxs-lookup"><span data-stu-id="06ea4-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="06ea4-118">Npm をインストールするには、コマンドラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="06ea4-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="06ea4-119">既に npm がインストールされているかどうかを確認し、インストールされているバージョンを確認するには、コマンドラインで次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="06ea4-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="06ea4-120">ノードバージョンマネージャーを使用して、複数のバージョンの Node.js と npm を切り替えることができますが、これは厳密には必要ありません。</span><span class="sxs-lookup"><span data-stu-id="06ea4-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="06ea4-121">この方法の詳細については、 [「npm の手順」を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。</span><span class="sxs-lookup"><span data-stu-id="06ea4-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="06ea4-122">Office 365 を取得する</span><span class="sxs-lookup"><span data-stu-id="06ea4-122">Get Office 365</span></span>

<span data-ttu-id="06ea4-123">Microsoft 365 アカウントをまだお持ちでない場合は、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することによって、更新可能な90日間の microsoft 365 サブスクリプションを無料で入手できます。</span><span class="sxs-lookup"><span data-stu-id="06ea4-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="06ea4-124">コード エディターのインストール</span><span class="sxs-lookup"><span data-stu-id="06ea4-124">Install a code editor</span></span>

<span data-ttu-id="06ea4-125">以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。</span><span class="sxs-lookup"><span data-stu-id="06ea4-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="06ea4-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="06ea4-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="06ea4-127">Atom</span><span class="sxs-lookup"><span data-stu-id="06ea4-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="06ea4-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="06ea4-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="06ea4-129">次の手順</span><span class="sxs-lookup"><span data-stu-id="06ea4-129">Next steps</span></span>

<span data-ttu-id="06ea4-130">独自のアドインを作成するか、スクリプトラボを使用して組み込みサンプルを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="06ea4-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="06ea4-131">Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="06ea4-131">Create an Office add-in</span></span>

<span data-ttu-id="06ea4-132">[5 分間のクイック スタート](/office/dev/add-ins/)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。</span><span class="sxs-lookup"><span data-stu-id="06ea4-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="06ea4-133">以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](/office/dev/add-ins/)を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="06ea4-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="06ea4-134">Script Lab を使用して API を調べる</span><span class="sxs-lookup"><span data-stu-id="06ea4-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="06ea4-135">Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="06ea4-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="06ea4-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="06ea4-136">See also</span></span>

- [<span data-ttu-id="06ea4-137">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="06ea4-137">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="06ea4-138">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="06ea4-138">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="06ea4-139">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="06ea4-139">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="06ea4-140">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="06ea4-140">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="06ea4-141">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="06ea4-141">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="06ea4-142">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="06ea4-142">Publish Office Add-ins</span></span>](../publish/publish.md)
