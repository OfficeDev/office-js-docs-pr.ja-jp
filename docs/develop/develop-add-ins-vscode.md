---
title: Visual Studio Code を使用して Office アドインを開発する
description: Visual Studio Code を使用して Office アドインを開発する方法
ms.date: 01/16/2020
localization_priority: Priority
ms.openlocfilehash: 25f9fa258ec1c6f9cc8344340cdf294091473168
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265481"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="2c662-103">Visual Studio Code を使用して Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2c662-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="2c662-104">この記事では、[Visual Studio Code (VS Code)](https://code.visualstudio.com) を使用して Office アドインを開発する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="2c662-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2c662-105">Visual Studio を使用して Office アドインを作成する方法については、「[Visual Studio を使用して Office アドインを作成する](develop-add-ins-visual-studio.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c662-105">For information about using Visual Studio to create an Office Add-in, see [Develop Office Add-ins with Visual Studio](develop-add-ins-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2c662-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="2c662-106">Prerequisites</span></span>

- [<span data-ttu-id="2c662-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="2c662-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="2c662-108">Yeoman ジェネレーターを使用してアドイン プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="2c662-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="2c662-109">統合開発環境 (IDE) として VS Code を使用している場合、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で Office アドイン プロジェクトを作成する必要があります。Yeoman ジェネレーターは、VS Code またはその他のエディターで管理できる Node.js プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="2c662-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="2c662-110">Yeoman ジェネレーターを使用して Office アドインを作成するには、作成するアドインの種類に対応する [5 分間のクイック スタート](../index.md)の指示に従います。</span><span class="sxs-lookup"><span data-stu-id="2c662-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="2c662-111">VS Code を使用してアドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2c662-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="2c662-112">Yeoman ジェネレーターがアドイン プロジェクトの作成を完了したら、VS Code でプロジェクトのルート フォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="2c662-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="2c662-113">Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。</span><span class="sxs-lookup"><span data-stu-id="2c662-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="2c662-114">Mac では、VS Code でプロジェクト フォルダーを開くためにそのコマンドを使用する前に、[`code` コマンドをパスに追加する](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)必要があります。</span><span class="sxs-lookup"><span data-stu-id="2c662-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="2c662-115">Yeoman ジェネレーターは、機能が制限された基本的なアドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="2c662-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="2c662-116">VS Code で[マニフェスト](add-in-manifests.md)、HTML、JavaScript または TypeScript、および CSS ファイルを編集することにより、アドインをカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="2c662-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="2c662-117">Yeoman ジェネレーターが作成するアドイン プロジェクトのプロジェクト構造とファイルの概要については、作成したアドインの種類に対応する [5 分間のクイック スタート](../index.md)内の Yeoman ジェネレーターのガイダンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c662-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="2c662-118">アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="2c662-118">Test and debug the add-in</span></span>

<span data-ttu-id="2c662-119">Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。</span><span class="sxs-lookup"><span data-stu-id="2c662-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="2c662-120">詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2c662-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="2c662-121">アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="2c662-121">Publish the add-in</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="2c662-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="2c662-122">See also</span></span>

- [<span data-ttu-id="2c662-123">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="2c662-123">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="2c662-124">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="2c662-124">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="2c662-125">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2c662-125">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="2c662-126">Office アドインを設計する</span><span class="sxs-lookup"><span data-stu-id="2c662-126">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="2c662-127">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="2c662-127">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="2c662-128">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="2c662-128">Publish Office Add-ins</span></span>](../publish/publish.md)