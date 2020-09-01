---
title: Office アドインを開発する
description: Office アドイン開発の概要を説明します。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 419880e8872df20be5a3de40f480f70be2b18859
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292779"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="b490d-103">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="b490d-103">Develop Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="b490d-104">この記事を読む前に、「[Building Office Add-ins (Office アドインの構築)](../overview/office-add-ins-fundamentals.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b490d-104">Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.</span></span>

<span data-ttu-id="b490d-105">すべての Office アドインは、Office アドイン プラットフォーム上で構築します。</span><span class="sxs-lookup"><span data-stu-id="b490d-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="b490d-106">すべての Office アドインでは共通のフレームワークが共有され、これにより特定の機能の実装が可能になります。</span><span class="sxs-lookup"><span data-stu-id="b490d-106">They share a common framework through which certain capabilities can be implemented.</span></span> <span data-ttu-id="b490d-107">どのようなアドインを構築する場合でも、アプリケーションやプラットフォームの可用性、Office JavaScript API のプログラミング パターン、アドインの設定と機能をマニフェスト ファイル上で指定する方法など、重要な概念を理解する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b490d-107">For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, and more.</span></span> <span data-ttu-id="b490d-108">開発に関するこれらの中心概念については、ドキュメントの「**Core concepts (中心概念)**」 > 「**Develop (開発)**」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b490d-108">Core development concepts like these are covered here in the **Core concepts** > **Develop** section of the documentation.</span></span> <span data-ttu-id="b490d-109">構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、 [Excel](../excel/index.yml)) を詳しく見る前に、ここに記載される情報を確認してください。</span><span class="sxs-lookup"><span data-stu-id="b490d-109">Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

> [!NOTE]
> <span data-ttu-id="b490d-110">「**Core concepts (中心概念)**」 > 「**Develop (開発)**」 > 「**How to (方法)**」セクションには、開発に関する具体的な概念やタスクについての記事があります。</span><span class="sxs-lookup"><span data-stu-id="b490d-110">The **Core concepts** > **Develop** > **How to** section of this documentation contains articles focused on specific development concepts or tasks.</span></span> <span data-ttu-id="b490d-111">同セクションでは、[Visual Studio Code を使用したアドイン開発](develop-add-ins-vscode.md)、[タスク ウィンドウをドキュメントと共に自動的に開く](automatically-open-a-task-pane-with-a-document.md)、[アドイン コマンドの作成](create-addin-commands.md)、[ダイアログ ボックスを開く](dialog-api-in-office-add-ins.md)などに関する情報が提供されています。</span><span class="sxs-lookup"><span data-stu-id="b490d-111">For example, you'll find information there about tasks like [developing add-ins with Visual Studio Code](develop-add-ins-vscode.md), [automatically opening a task pane with a document](automatically-open-a-task-pane-with-a-document.md), [creating add-in commands](create-addin-commands.md), and [opening a dialog box](dialog-api-in-office-add-ins.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="b490d-112">次の手順</span><span class="sxs-lookup"><span data-stu-id="b490d-112">Next steps</span></span>

<span data-ttu-id="b490d-113">ここで説明する中心概念について理解したら、構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、[Excel](../excel/index.yml)) を確認します。</span><span class="sxs-lookup"><span data-stu-id="b490d-113">After you're familiar with the core concepts covered here, explore the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span> <span data-ttu-id="b490d-114">ドキュメントの各アプリケーション固有のセクションには、特定の Office アプリケーション用のアドインの構築に関する具体的な情報が記載されています。</span><span class="sxs-lookup"><span data-stu-id="b490d-114">Each application-specific section of the documentation contains information specifically about building add-ins for a certain Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="b490d-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="b490d-115">See also</span></span>

- [<span data-ttu-id="b490d-116">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="b490d-116">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="b490d-117">Office アドインの構築</span><span class="sxs-lookup"><span data-stu-id="b490d-117">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="b490d-118">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="b490d-118">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="b490d-119">Office アドインの設計</span><span class="sxs-lookup"><span data-stu-id="b490d-119">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="b490d-120">Office アドインのテストとデバッグ</span><span class="sxs-lookup"><span data-stu-id="b490d-120">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="b490d-121">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="b490d-121">Publish Office Add-ins</span></span>](../publish/publish.md)
