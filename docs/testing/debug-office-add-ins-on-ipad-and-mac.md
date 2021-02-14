---
title: Mac で Office アドインをデバッグする
description: Mac を使用してアドインをデバッグするOffice説明します。
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: b2164e3ed672b2911db6841fad24441b67882204
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237946"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="4f892-103">Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="4f892-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="4f892-p101">アドインは HTML と JavaScript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、Mac で動作するアドインをデバッグする方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="4f892-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="4f892-106">Mac での Safari Web インスペクタを使用したデバッグ</span><span class="sxs-lookup"><span data-stu-id="4f892-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="4f892-107">作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="4f892-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="4f892-108">Mac で Office アドインをデバッグするには、Mac OS High Sierra と Mac Office バージョン 16.9.1 (ビルド 18012504) 以降が必要です。</span><span class="sxs-lookup"><span data-stu-id="4f892-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office version 16.9.1 (build 18012504) or later.</span></span> <span data-ttu-id="4f892-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="4f892-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="4f892-110">最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。</span><span class="sxs-lookup"><span data-stu-id="4f892-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > <span data-ttu-id="4f892-111">Mac App Store ビルドのOfficeフラグはサポート `OfficeWebAddinDeveloperExtras` されていません。</span><span class="sxs-lookup"><span data-stu-id="4f892-111">Mac App Store builds of Office do not support the `OfficeWebAddinDeveloperExtras` flag.</span></span>

<span data-ttu-id="4f892-112">次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="4f892-112">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="4f892-113">アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="4f892-113">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="4f892-114">このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="4f892-114">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4f892-115">インスペクタとダイアログ フリッカーを使おうとしている場合は、Office を最新バージョンに更新してください。</span><span class="sxs-lookup"><span data-stu-id="4f892-115">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="4f892-116">それでも、ちらつきが解消しない場合は、次の回避策を試してください。</span><span class="sxs-lookup"><span data-stu-id="4f892-116">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="4f892-117">ダイアログのサイズを変更します。</span><span class="sxs-lookup"><span data-stu-id="4f892-117">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="4f892-118">**[要素の検査]** を選択します (新しいウィンドウが開きます)。</span><span class="sxs-lookup"><span data-stu-id="4f892-118">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="4f892-119">ダイアログを元のサイズに変更します。</span><span class="sxs-lookup"><span data-stu-id="4f892-119">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="4f892-120">必要に応じてインスペクタを使用します。</span><span class="sxs-lookup"><span data-stu-id="4f892-120">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="4f892-121">Mac 上の Office アプリケーションのキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="4f892-121">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
