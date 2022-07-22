---
title: ドキュメントが開いたら、Office アドインでコードを実行する
description: ドキュメントが開いたときに Office アドイン アドインでコードを実行する方法について説明します。
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1a1c3277a349dc4054da5f089c62331296590021
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958440"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>ドキュメントが開いたら、Office アドインでコードを実行する

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

ドキュメントが開いたらすぐにコードを読み込んで実行するように Office アドインを構成できます。 これは、アドインが表示される前に、イベント ハンドラーの登録、作業ウィンドウのデータの事前読み込み、UI の同期、またはその他のタスクの実行を行う必要がある場合に便利です。

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>ドキュメントが開いたときに読み込むアドインを構成する

次のコードは、ドキュメントを開いたときにアドインを読み込んで実行を開始するように構成します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> メソッドは `setStartupBehavior` 非同期です。

## <a name="place-startup-code-in-officeinitialize"></a>Office.initialize にスタートアップ コードを配置する

ドキュメントを開いたときに読み込むようにアドインが構成されると、すぐに実行されます。 `Office.initialize`イベント ハンドラーが呼び出されます。 スタートアップ コードをイベント ハンドラーに`Office.initialize``Office.onReady`配置します。

次の Excel アドイン コードは、作業中のワークシートから変更イベントのイベント ハンドラーを登録する方法を示しています。 ドキュメントを開くときに読み込むアドインを構成した場合、このコードはドキュメントを開いたときにイベント ハンドラーを登録します。 作業ウィンドウを開く前に、変更イベントを処理できます。

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
    await Excel.run(async (context) => {    
        await context.sync();
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
  });
}
```

次の PowerPoint アドイン コードは、PowerPoint ドキュメントから選択変更イベントのイベント ハンドラーを登録する方法を示しています。 ドキュメントを開くときに読み込むアドインを構成した場合、このコードはドキュメントを開いたときにイベント ハンドラーを登録します。 作業ウィンドウを開く前に、変更イベントを処理できます。

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>ドキュメントを開いたときに読み込み動作が行われなくなった場合にアドインを構成する

次のコードでは、ドキュメントを開いたときにアドインを開始しないように構成します。 代わりに、リボン ボタンの選択や作業ウィンドウの開きなど、ユーザーが何らかの方法で操作したときに開始されます。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>現在の読み込み動作を取得する

現在のスタートアップ動作を確認するには、オブジェクトを返す次のメソッドを `Office.StartupBehavior` 実行します。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>関連項目

- [Office アドインを構成して共有 JavaScript ランタイムを使用する](configure-your-add-in-to-use-a-shared-runtime.md)
- [Excel カスタム関数と作業ウィンドウのチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Excel JavaScript API を使用してイベントを操作する](../excel/excel-add-ins-events.md)
