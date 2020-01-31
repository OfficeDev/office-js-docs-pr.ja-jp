---
title: Office のキャッシュをクリアする
description: コンピューターで Office のキャッシュをクリアする方法について説明します。
ms.date: 01/21/2020
localization_priority: Priority
ms.openlocfilehash: 68e5c022671844ee44bf8ca8ac00bc5af6564bad
ms.sourcegitcommit: 43166612e9b4bf7a73312a572663c8696353dbc6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/29/2020
ms.locfileid: "41580968"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="46a9d-103">Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="46a9d-103">Clear the Office cache</span></span>

<span data-ttu-id="46a9d-104">以前に Windows、Mac、または iOS にサイドロードしたアドインは、コンピューターで Office のキャッシュをクリアすることにより削除できます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="46a9d-105">また、アドインのマニフェストに変更を加えた場合は (アイコンのファイル名やアドイン コマンドのテキストを更新した場合など)、Office のキャッシュをクリアし、更新されたマニフェストを使用してアドインをサイドロードし直す必要があります。</span><span class="sxs-lookup"><span data-stu-id="46a9d-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="46a9d-106">これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="46a9d-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="46a9d-107">Windows で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="46a9d-107">Clear the Office cache on Windows</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="46a9d-108">Excel、Word、および PowerPoint</span><span class="sxs-lookup"><span data-stu-id="46a9d-108">Excel, Word, and PowerPoint</span></span> 

<span data-ttu-id="46a9d-109">Windows で Excel、Word、および PowerPoint の Office のキャッシュをクリアするには、フォルダー `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 内のコンテンツを削除します。</span><span class="sxs-lookup"><span data-stu-id="46a9d-109">To clear the Office cache on Windows for Excel, Word, and PowerPoint, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

### <a name="outlook-windows-10"></a><span data-ttu-id="46a9d-110">Outlook (Windows 10)</span><span class="sxs-lookup"><span data-stu-id="46a9d-110">Outlook (Windows 10)</span></span>

<span data-ttu-id="46a9d-111">アドインが Microsoft Edge で実行されているときに Windows 10 で Office のキャッシュをクリアするには、Microsoft Edge DevTools を使用します。</span><span class="sxs-lookup"><span data-stu-id="46a9d-111">To clear the Outlook cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.</span></span>

> [!TIP]
> <span data-ttu-id="46a9d-112">サイドロードされたアドインに HTML や JavaScript のソース ファイルへの最近の変更を反映させたいなら、次の手順でキャッシュをクリアする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="46a9d-112">If you're just wanting the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to use the following steps to clear the cache.</span></span> <span data-ttu-id="46a9d-113">代わりに、アドインの作業ウィンドウにフォーカスを置き (タスク ウィンドウ内の任意の場所をクリック)、**F5** キーを押してアドインをリロードします。</span><span class="sxs-lookup"><span data-stu-id="46a9d-113">Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.</span></span> 

> [!NOTE]
> <span data-ttu-id="46a9d-114">次の手順を使用して Outlook のキャッシュをクリアするには、アドインに作業ウィンドウが必要です。</span><span class="sxs-lookup"><span data-stu-id="46a9d-114">To clear the Outlook cache using the following steps, your add-in must have a task pane.</span></span> <span data-ttu-id="46a9d-115">アドインが UI を使用しない場合 (たとえば、[送信時](/outlook/add-ins/outlook-on-send-addins)機能を使用するアドインの場合)、次の手順でキャッシュをクリアする前に、同じドメインを [SourceLocation](../reference/manifest/sourcelocation.md) に使用するアドインに作業ウィンドウを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="46a9d-115">If your add-in is a UI-less add-in -- for example, one that uses the [on-send](/outlook/add-ins/outlook-on-send-addins) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.</span></span>

1. <span data-ttu-id="46a9d-116">[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj) をインストールします。</span><span class="sxs-lookup"><span data-stu-id="46a9d-116">Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span></span>

2. <span data-ttu-id="46a9d-117">Outlook でアドインを開きます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-117">Open your add-in in Outlook.</span></span>

3. <span data-ttu-id="46a9d-118">Microsoft Edge DevTools を実行します。</span><span class="sxs-lookup"><span data-stu-id="46a9d-118">Run the Microsoft Edge DevTools.</span></span>

4. <span data-ttu-id="46a9d-119">Microsoft Edge DevTools で、[**ローカル**] タブを開きます。アドインの名前が一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-119">In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

5. <span data-ttu-id="46a9d-120">アドイン名を選択して、アドインにデバッガーをアタッチします。</span><span class="sxs-lookup"><span data-stu-id="46a9d-120">Select the add-in name to attach the debugger to your add-in.</span></span> <span data-ttu-id="46a9d-121">デバッガーがアドインにアタッチされると、新しい Microsoft Edge DevTools ウィンドウが開きます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-121">A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.</span></span>

6. <span data-ttu-id="46a9d-122">新しいウィンドウの [**ネットワーク**] タブで、[**キャッシュのクリア**] ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="46a9d-122">On the **Network** tab of the new window, select the **Clear cache** button.</span></span>

    ![[キャッシュのクリア] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-clear-cache.png)

7. <span data-ttu-id="46a9d-124">これらの手順を完了しても望む結果が得られない場合は、[**常にサーバーから更新する**] ボタンを選択することもできます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-124">If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.</span></span>

    ![[常にサーバーから更新する] ボタンが強調表示された Microsoft Edge DevTools のスクリーンショット](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="46a9d-126">Mac で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="46a9d-126">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="46a9d-127">iOS で Office のキャッシュをクリアする</span><span class="sxs-lookup"><span data-stu-id="46a9d-127">Clear the Office cache on iOS</span></span>

<span data-ttu-id="46a9d-128">iOS で Office のキャッシュをクリアするには、アドイン内の JavaScript から `window.location.reload(true)` を呼び出し、強制的に再読み込みを行います。</span><span class="sxs-lookup"><span data-stu-id="46a9d-128">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="46a9d-129">別の方法として、Office を再インストールすることもできます。</span><span class="sxs-lookup"><span data-stu-id="46a9d-129">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="46a9d-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="46a9d-130">See also</span></span>

- [<span data-ttu-id="46a9d-131">Office アドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="46a9d-131">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [<span data-ttu-id="46a9d-132">ランタイム ログを使用してアドインをデバッグする</span><span class="sxs-lookup"><span data-stu-id="46a9d-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="46a9d-133">テスト用に Office アドインをサイドロードする</span><span class="sxs-lookup"><span data-stu-id="46a9d-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="46a9d-134">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="46a9d-134">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="46a9d-135">Office アドインのマニフェストを検証する</span><span class="sxs-lookup"><span data-stu-id="46a9d-135">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)

