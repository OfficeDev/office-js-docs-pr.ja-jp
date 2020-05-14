---
title: ドキュメントが開いたときに Excel アドインでコードを実行する
description: ドキュメントが開いたときに、Excel アドインでコードを実行します。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: 0a9090315a4ddca80e25a94092c779a3f3271087
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217950"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens"></a>ドキュメントが開いたときに Excel アドインでコードを実行する

ドキュメントが開かれるとすぐに、コードを読み込んで実行するように Excel アドインを構成することができます。 これは、アドインが表示される前に、イベントハンドラーの登録、作業ウィンドウのデータの事前読み込み、UI の同期、またはその他のタスクの実行が必要な場合に便利です。

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a>ドキュメントが開いたときに読み込まれるようにアドインを構成する

次のコードは、ドキュメントが開かれたときに読み込み、実行を開始するようにアドインを構成します。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> `setStartupBehavior`メソッドは非同期です。

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a>ドキュメントを開くときに読み込み動作を行わないようにアドインを構成する

次のコードは、ドキュメントが開かれたときに開始しないようにアドインを構成します。 代わりに、ユーザーが何らかの方法 (リボンボタンを選択したときや作業ウィンドウを開いたときなど) に実行されます。

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a>現在の読み込み動作を取得する

現在のスタートアップ動作を確認するには、次の関数を実行します。この関数は、Office の StartupBehavior オブジェクトを返します。

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a>ドキュメントが開いたときにコードを実行する方法

アドインがドキュメントを開いたときに読み込むように構成すると、すぐに実行されます。 `Office.initialize`イベントハンドラーが呼び出されます。 スタートアップコードを `Office.initialize` イベントハンドラーに配置します。

次のコードは、作業中のワークシートから変更イベントのイベントハンドラーを登録する方法を示しています。 アドインをドキュメントを開いたときに読み込むように構成した場合、このコードは、ドキュメントが開かれたときにイベントハンドラーを登録します。 作業ウィンドウを開く前に、変更イベントを処理することができます。


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
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

## <a name="see-also"></a>関連項目

- [Excel カスタム関数と作業ウィンドウチュートリアルの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)