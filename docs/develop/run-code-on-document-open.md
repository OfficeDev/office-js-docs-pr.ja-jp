---
title: ドキュメントが開いたら、Office アドインでコードを実行する
description: ドキュメントが開いたら、Officeアドインでコードを実行する方法について学習します。
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: 20cd7a90f34c0141ca166119ceae92960a904595
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744079"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a>ドキュメントが開いたら、Office アドインでコードを実行する

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

ドキュメントが開Officeコードを読み込み、実行するアドインを構成できます。 これは、アドインが表示される前に、イベント ハンドラーの登録、作業ウィンドウのデータの事前読み込み、UI の同期、その他のタスクの実行が必要な場合に便利です。

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>ドキュメントが開くと読み込むアドインを構成する

次のコードは、ドキュメントを開いたときに読み込み、実行を開始するアドインを構成します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> メソッド `setStartupBehavior` は非同期です。

## <a name="place-startup-code-in-officeinitialize"></a>スタートアップ コードを Office.initialize に配置する

ドキュメントが開いているときに読み込むアドインが構成されている場合、すぐに実行されます。 イベント `Office.initialize` ハンドラーが呼び出されます。 スタートアップ コードをイベント ハンドラーに`Office.initialize``Office.onReady`配置します。

次のExcelアドイン コードは、アクティブなワークシートから変更イベントのイベント ハンドラーを登録する方法を示しています。 ドキュメントを開く際に読み込むアドインを構成すると、ドキュメントを開く際にイベント ハンドラーが登録されます。 作業ウィンドウを開く前に変更イベントを処理できます。

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

次のPowerPointアドイン コードは、ドキュメントから選択変更イベントのイベント ハンドラーを登録するPowerPointします。 ドキュメントを開く際に読み込むアドインを構成すると、ドキュメントを開く際にイベント ハンドラーが登録されます。 作業ウィンドウを開く前に変更イベントを処理できます。

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

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>ドキュメントを開く時に読み込み動作を行う必要がないアドインを構成する

次のコードでは、ドキュメントを開いたときにアドインを起動しなく設定します。 代わりに、リボン ボタンの選択や作業ウィンドウの開きなど、ユーザーが何らかの方法で操作を開始します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>現在の読み込み動作を取得する

現在の起動動作を確認するには、オブジェクトを返す次の関数を実行 `Office.StartupBehavior` します。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="see-also"></a>関連項目

- [Office アドインを構成して共有 JavaScript ランタイムを使用する](configure-your-add-in-to-use-a-shared-runtime.md)
- [カスタム関数と作業ウィンドウのチュートリアルExcelデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Excel JavaScript API を使用してイベントを操作する](../excel/excel-add-ins-events.md)
