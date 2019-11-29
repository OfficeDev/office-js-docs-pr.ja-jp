---
title: Mac で Office アドインをデバッグする
description: ''
ms.date: 11/26/2019
localization_priority: Priority
ms.openlocfilehash: 4a14638a6db4e1b3edb9de62a8977885547ad8e3
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629750"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="3e251-102">Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="3e251-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="3e251-p101">アドインは HTML と JavaScript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、Mac で動作するアドインをデバッグする方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="3e251-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="3e251-105">Mac での Safari Web インスペクタを使用したデバッグ</span><span class="sxs-lookup"><span data-stu-id="3e251-105">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="3e251-106">作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="3e251-106">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="3e251-107">Mac の Office アドインをデバッグするには、Mac OS High Sierra と Mac Office バージョン 16.9.1 (ビルド 18012504) 以降の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="3e251-107">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="3e251-108">Office Mac ビルドをまだお持ちでない場合は、[Office 365 Developer Program](https://aka.ms/o365devprogram) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="3e251-108">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="3e251-109">最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。</span><span class="sxs-lookup"><span data-stu-id="3e251-109">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="3e251-110">次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="3e251-110">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="3e251-111">アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="3e251-111">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="3e251-112">このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="3e251-112">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3e251-113">インスペクタとダイアログ フリッカーを使おうとしている場合は、Office を最新バージョンに更新してください。</span><span class="sxs-lookup"><span data-stu-id="3e251-113">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="3e251-114">それでも、ちらつきが解消しない場合は、次の回避策を試してください。</span><span class="sxs-lookup"><span data-stu-id="3e251-114">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="3e251-115">ダイアログのサイズを変更します。</span><span class="sxs-lookup"><span data-stu-id="3e251-115">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="3e251-116">**[要素の検査]** を選択します (新しいウィンドウが開きます)。</span><span class="sxs-lookup"><span data-stu-id="3e251-116">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="3e251-117">ダイアログを元のサイズに変更します。</span><span class="sxs-lookup"><span data-stu-id="3e251-117">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="3e251-118">必要に応じてインスペクタを使用します。</span><span class="sxs-lookup"><span data-stu-id="3e251-118">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="3e251-119">Mac 上の Office アプリケーションのキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="3e251-119">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
