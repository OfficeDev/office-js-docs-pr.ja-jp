---
title: Visual Studio Code を使用して Office アドインを開発する
description: Visual Studio Code を使用して Office アドインを開発する方法
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: a18d8a74ff269b32e83c836b06629850873e507b
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670495"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="e4c04-103">Visual Studio Code を使用して Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="e4c04-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="e4c04-104">この記事では、[Visual Studio Code (VS Code)](https://code.visualstudio.com) を使用して Office アドインを開発する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e4c04-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e4c04-105">Visual Studio を使用して Office アドインを作成する方法については、「[Visual Studio での Office アドインの作成とデバッグ](create-and-debug-office-add-ins-in-visual-studio.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4c04-105">For information about using Visual Studio to create an Office Add-in, see [Create and debug Office Add-ins in Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e4c04-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="e4c04-106">Prerequisites</span></span>

- [<span data-ttu-id="e4c04-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="e4c04-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="e4c04-108">Yeoman ジェネレーターを使用してアドイン プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="e4c04-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="e4c04-109">統合開発環境 (IDE) として VS Code を使用している場合、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で Office アドイン プロジェクトを作成する必要があります。Yeoman ジェネレーターは、VS Code またはその他のエディターで管理できる Node.js プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="e4c04-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="e4c04-110">Yeoman ジェネレーターを使用して Office アドインを作成するには、作成するアドインの種類に対応する [5 分間のクイック スタート](../index.md)の指示に従います。</span><span class="sxs-lookup"><span data-stu-id="e4c04-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="e4c04-111">VS Code を使用してアドインを開発する</span><span class="sxs-lookup"><span data-stu-id="e4c04-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="e4c04-112">Yeoman ジェネレーターがアドイン プロジェクトの作成を完了したら、VS Code でプロジェクトのルート フォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="e4c04-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="e4c04-113">Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。</span><span class="sxs-lookup"><span data-stu-id="e4c04-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="e4c04-114">Mac では、VS Code でプロジェクト フォルダーを開くためにそのコマンドを使用する前に、[`code` コマンドをパスに追加する](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)必要があります。</span><span class="sxs-lookup"><span data-stu-id="e4c04-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="e4c04-115">Yeoman ジェネレーターは、機能が制限された基本的なアドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="e4c04-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="e4c04-116">VS Code で[マニフェスト](add-in-manifests.md)、HTML、JavaScript または TypeScript、および CSS ファイルを編集することにより、アドインをカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="e4c04-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="e4c04-117">Yeoman ジェネレーターが作成するアドイン プロジェクトのプロジェクト構造とファイルの概要については、作成したアドインの種類に対応する [5 分間のクイック スタート](../index.md)内の Yeoman ジェネレーターのガイダンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4c04-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="e4c04-118">アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="e4c04-118">To run and debug the add-in</span></span>

<span data-ttu-id="e4c04-119">Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e4c04-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="e4c04-120">詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e4c04-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="e4c04-121">アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="e4c04-121">Publish the add-in.</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="e4c04-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="e4c04-122">See also</span></span>

- [<span data-ttu-id="e4c04-123">5 分間のクイック スタート</span><span class="sxs-lookup"><span data-stu-id="e4c04-123">5-Minute Quick Starts</span></span>](../index.md)
- <span data-ttu-id="e4c04-124">[Script Lab を使用して Office JavaScript API を探索する](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="e4c04-124">To learn more, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).</span></span>
- [<span data-ttu-id="e4c04-125">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="e4c04-125">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="e4c04-126">Office アドインを展開し、発行する</span><span class="sxs-lookup"><span data-stu-id="e4c04-126">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)