---
title: Mac で Office アドインをデバッグする
description: ''
ms.date: 05/21/2019
localization_priority: Priority
ms.openlocfilehash: 0505dcc49ea98040f1c4891621c8e30a8cbeaff4
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432279"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="fcb2f-102">Mac で Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="fcb2f-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="fcb2f-p101">Windows でのアドインの開発とデバッグには Visual Studio を使用できますが、Mac で使用してアドインをデバッグすることはできません。アドインは HTML と JavaScript を使用して開発されているため、さまざまなプラットフォームで機能するように設計されていますが、さまざまなブラウザーで HTML の表示方法に微妙な違いがあります。この記事では、Mac で動作するアドインをデバッグする方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on a Mac. Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="fcb2f-106">Mac での Safari Web インスペクタを使用したデバッグ</span><span class="sxs-lookup"><span data-stu-id="fcb2f-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="fcb2f-107">作業ウィンドウまたはコンテンツ アドインに UI を表示するアドインを使用している場合は、Safari Web インスペクタを使用して Office アドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="fcb2f-108">Mac の Office アドインをデバッグするには、Mac OS High Sierra と Mac Office バージョン 16.9.1 (ビルド 18012504) 以降の両方が必要です。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="fcb2f-109">Office Mac ビルドをまだお持ちでない場合は、[Office 365 Developer Program](https://aka.ms/o365devprogram) に参加することで入手できます。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="fcb2f-110">最初に端末を開き、該当する Office アプリケーションの `OfficeWebAddinDeveloperExtras` プロパティを以下のように設定します。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="fcb2f-111">次に Office アプリケーションを開き、[アドインをサイドロードします](sideload-an-office-add-in-on-ipad-and-mac.md)。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-111">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="fcb2f-112">アドインを右クリックします。コンテキスト メニューに **[要素の検査]** オプションが表示されるはずです。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="fcb2f-113">このオプションを選択するとインスペクタが表示されます。インスペクタでは、ブレークポイントを設定してアドインをデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="fcb2f-114">インスペクタとダイアログ フリッカーを使おうとしている場合は、Office を最新バージョンに更新してください。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-114">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="fcb2f-115">それでも、ちらつきが解消しない場合は、次の回避策を試してください。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-115">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="fcb2f-116">ダイアログのサイズを変更します。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="fcb2f-117">**[要素の検査]** を選択します (新しいウィンドウが開きます)。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="fcb2f-118">ダイアログを元のサイズに変更します。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="fcb2f-119">必要に応じてインスペクタを使用します。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-119">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="fcb2f-120">Mac 上の Office アプリケーションのキャッシュのクリア</span><span class="sxs-lookup"><span data-stu-id="fcb2f-120">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="fcb2f-p105">アドインはパフォーマンス上の理由から、Office for Mac でキャッシュされることが多いです。通常、キャッシュはアドインを再読み込みすることでクリアされます。同じドキュメント内に複数のアドインが存在する場合、再読み込み時にキャッシュを自動的にクリアするプロセスは信頼できない場合があります。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-p105">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="fcb2f-124">Mac では、`~/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーのコンテンツを削除することによってキャッシュを手動でクリアできます。</span><span class="sxs-lookup"><span data-stu-id="fcb2f-124">On a Mac, you can clear the cache manually by deleting everything in the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
