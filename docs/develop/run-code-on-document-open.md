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
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>ドキュメントが開Officeアドインでコードを実行する

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

ドキュメントが開Officeすぐにコードを読み込み、実行するアドインを構成できます。 これは、アドインが表示される前に、イベント ハンドラーの登録、作業ウィンドウのデータの事前読み込み、UI の同期、その他のタスクの実行が必要な場合に便利です。

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>ドキュメントが開くと読み込むアドインを構成する

次のコードは、ドキュメントを開く際に読み込み、実行を開始するアドインを構成します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> メソッド `setStartupBehavior` は非同期です。

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>ドキュメントを開く場合の読み込み時の動作が発生しなか、アドインを構成する

次のコードは、ドキュメントを開く際にアドインが起動しなく設定します。 代わりに、リボン ボタンを選択したり、作業ウィンドウを開くなど、何らかの方法でユーザーが操作を開始します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>現在の読み込み動作を取得する

現在のスタートアップ動作を確認するには、オブジェクトを返す次の関数を実行 `Office.StartupBehavior` します。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>ドキュメントを開く際にコードを実行する方法

ドキュメントを開く際に読み込むアドインが構成されている場合は、すぐに実行されます。 イベント `Office.initialize` ハンドラーが呼び出されます。 スタートアップ コードをイベント ハンドラー `Office.initialize` に `Office.onReady` 配置します。

次の Excel アドイン コードは、アクティブ ワークシートから変更イベントのイベント ハンドラーを登録する方法を示しています。 ドキュメントを開く際に読み込むアドインを構成する場合、このコードはドキュメントを開く際にイベント ハンドラーを登録します。 作業ウィンドウを開く前に変更イベントを処理できます。

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

次の PowerPoint アドイン コードは、PowerPoint ドキュメントから選択変更イベントのイベント ハンドラーを登録する方法を示しています。 ドキュメントを開く際に読み込むアドインを構成する場合、このコードはドキュメントを開く際にイベント ハンドラーを登録します。 作業ウィンドウを開く前に変更イベントを処理できます。

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

## <a name="see-also"></a>関連項目

- [共有 JavaScript Office使用する新しいアドインを構成する](configure-your-add-in-to-use-a-shared-runtime.md)
- [Excel カスタム関数と作業ウィンドウのチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Excel JavaScript API を使用してイベントを操作する](../excel/excel-add-ins-events.md)
