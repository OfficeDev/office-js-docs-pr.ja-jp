---
title: ドキュメントが開Officeアドインでコードを実行する
description: ドキュメントが開Officeアドインでコードを実行する方法について学習します。
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789245"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a><span data-ttu-id="4fb43-103">ドキュメントが開Officeアドインでコードを実行する</span><span class="sxs-lookup"><span data-stu-id="4fb43-103">Run code in your Office Add-in when the document opens</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="4fb43-104">ドキュメントが開Officeすぐにコードを読み込み、実行するアドインを構成できます。</span><span class="sxs-lookup"><span data-stu-id="4fb43-104">You can configure your Office Add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="4fb43-105">これは、アドインが表示される前に、イベント ハンドラーの登録、作業ウィンドウのデータの事前読み込み、UI の同期、その他のタスクの実行が必要な場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="4fb43-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="4fb43-106">ドキュメントが開くと読み込むアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="4fb43-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="4fb43-107">次のコードは、ドキュメントを開く際に読み込み、実行を開始するアドインを構成します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="4fb43-108">メソッド `setStartupBehavior` は非同期です。</span><span class="sxs-lookup"><span data-stu-id="4fb43-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="4fb43-109">ドキュメントを開く場合の読み込み時の動作が発生しなか、アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="4fb43-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="4fb43-110">次のコードは、ドキュメントを開く際にアドインが起動しなく設定します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="4fb43-111">代わりに、リボン ボタンを選択したり、作業ウィンドウを開くなど、何らかの方法でユーザーが操作を開始します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-111">Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="4fb43-112">現在の読み込み動作を取得する</span><span class="sxs-lookup"><span data-stu-id="4fb43-112">Get the current load behavior</span></span>

<span data-ttu-id="4fb43-113">現在のスタートアップ動作を確認するには、オブジェクトを返す次の関数を実行 `Office.StartupBehavior` します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-113">To determine what the current startup behavior is, run the following function, which returns an `Office.StartupBehavior` object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="4fb43-114">ドキュメントを開く際にコードを実行する方法</span><span class="sxs-lookup"><span data-stu-id="4fb43-114">How to run code when the document opens</span></span>

<span data-ttu-id="4fb43-115">ドキュメントを開く際に読み込むアドインが構成されている場合は、すぐに実行されます。</span><span class="sxs-lookup"><span data-stu-id="4fb43-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="4fb43-116">イベント `Office.initialize` ハンドラーが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4fb43-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="4fb43-117">スタートアップ コードをイベント ハンドラー `Office.initialize` に `Office.onReady` 配置します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-117">Place your startup code in the `Office.initialize` or `Office.onReady` event handler.</span></span>

<span data-ttu-id="4fb43-118">次の Excel アドイン コードは、アクティブ ワークシートから変更イベントのイベント ハンドラーを登録する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4fb43-118">The following Excel add-in code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="4fb43-119">ドキュメントを開く際に読み込むアドインを構成する場合、このコードはドキュメントを開く際にイベント ハンドラーを登録します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="4fb43-120">作業ウィンドウを開く前に変更イベントを処理できます。</span><span class="sxs-lookup"><span data-stu-id="4fb43-120">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}
```

<span data-ttu-id="4fb43-121">次の PowerPoint アドイン コードは、PowerPoint ドキュメントから選択変更イベントのイベント ハンドラーを登録する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4fb43-121">The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document.</span></span> <span data-ttu-id="4fb43-122">ドキュメントを開く際に読み込むアドインを構成する場合、このコードはドキュメントを開く際にイベント ハンドラーを登録します。</span><span class="sxs-lookup"><span data-stu-id="4fb43-122">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="4fb43-123">作業ウィンドウを開く前に変更イベントを処理できます。</span><span class="sxs-lookup"><span data-stu-id="4fb43-123">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="see-also"></a><span data-ttu-id="4fb43-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="4fb43-124">See also</span></span>

- [<span data-ttu-id="4fb43-125">共有 JavaScript Office使用する新しいアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="4fb43-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="4fb43-126">Excel カスタム関数と作業ウィンドウのチュートリアルの間でデータとイベントを共有する</span><span class="sxs-lookup"><span data-stu-id="4fb43-126">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="4fb43-127">Excel JavaScript API を使用してイベントを操作する</span><span class="sxs-lookup"><span data-stu-id="4fb43-127">Work with Events using the Excel JavaScript API</span></span>](../excel/excel-add-ins-events.md)
