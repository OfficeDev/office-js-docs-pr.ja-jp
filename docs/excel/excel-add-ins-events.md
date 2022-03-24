---
title: Excel JavaScript API を使用してイベントを操作する
description: JavaScript オブジェクトのイベントExcel一覧です。 これには、イベント ハンドラーと関連付けられたパターンの使用に関する情報が含まれます。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c15beba846fc5348143b63dfb07321b6dad01ea2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745013"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>Excel JavaScript API を使用してイベントを操作する

この記事では、Excel のイベント操作に関連する重要な概念について説明します。また、Excel JavaScript API を使用したイベント ハンドラーの登録、イベントの処理、およびイベント ハンドラーの削除の方法を示すコード例も提供します。

## <a name="events-in-excel"></a>Excel のイベント

Excel ブックで特定の種類の変更が発生するたびに、イベント通知がトリガーされます。 Excel JavaScript API を使用すると、イベント ハンドラーを登録できます。このハンドラーによって、特定のイベントが発生したときに、アドインで目的の関数を自動的に実行できるようになります。 現時点でサポートされているイベントは次のとおりです。

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onActivated` | オブジェクトがアクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)、[**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | ブックがアクティブ化されると発生します。 | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | オブジェクトがコレクションに追加されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | ブックで `autoSave` の設定が変更されると発生します。 | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | ワークシートの計算が完了したとき (あるいはコレクションのすべてのワークシートが完了したとき) に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | 個々のセルまたはコメントのデータが変更された場合に発生します。 | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) |
| `onColumnSorted` | 1 つ以上の列を並べ替えたときに発生します。 これは、左から右に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | バインド内でデータまたは書式設定が変更されるときに発生します。 | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | オブジェクトが非アクティブ化されたときに発生します。 | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)、[**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | オブジェクトがコレクションから削除されたときに発生します。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | ワークシートで書式設定が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | 数式が変更された場合に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onProtectionChanged` | ワークシートの保護状態が変更された場合に発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | 特定のワークシート上の行非表示状態が変更されたときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | 1 つ以上の行を並べ替えたときに発生します。 これは、上から下に並べ替えを実行したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | アクティブなセルまたは選択範囲が変更されたときに発生します。 | [**バインド**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member)、[**テーブル、**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)[**ブック、**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member)[**ワークシート、**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | ドキュメント内の設定が変更されるときに発生します。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | ワークシートで左クリック / タップされたアクションが発生したときに発生します。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |

### <a name="events-in-preview"></a>プレビューでのイベント

> [!NOTE]
> 次のイベントは現在、パブリック プレビューでのみ利用できます。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| イベント | 説明 | サポートされているオブジェクト |
|:---------------|:-------------|:-----------|
| `onFiltered` | フィルターがオブジェクトに適用されたときに発生します。 | [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

### <a name="event-triggers"></a>イベント トリガー

Excel ブックのイベントは、次の事項でトリガーできます。

- ブックを変更する Excel ユーザー インターフェイス (UI) からのユーザー操作
- ブックを変更する Office アドイン (JavaScript) コード
- ブックを変更する VBA アドイン (マクロ) コード

Excel の既定の動作に準拠する変更により、それに対応するブックのイベントがトリガーされます。

### <a name="lifecycle-of-an-event-handler"></a>イベント ハンドラーのライフサイクル

アドインがイベント ハンドラーを登録すると、そのイベント ハンドラーが作成されます。 アドインがイベント ハンドラーを登録解除するか、アドインが更新、再読み込み、または閉じられると、イベント ハンドラーは破棄されます。 イベント ハンドラーは Excel ファイルの一部として保持されず、また Excel on the web のセッション間でも保持されません。

> [!CAUTION]
> イベントが登録されているオブジェクト (`onChanged` イベントが登録されているテーブルなど) が削除されると、イベント ハンドラーはトリガーされませんが、アドインまたは Excel セッションが更新または閉じるまではメモリで維持されます。

### <a name="events-and-coauthoring"></a>イベントと共同編集

[共同編集機能](co-authoring-in-excel-add-ins.md)により、複数のユーザーが連携して同じ Excel ブックを同時に編集できるようになります。共同編集でトリガーできるイベント (`onChanged` など) の場合、対応する **Event** オブジェクトには **source** プロパティが含まれるようになります。このプロパティは、イベントが現在のユーザーによってローカルにトリガーされた (`event.source = Local`) ものか、リモートの共同作成者によってトリガーされた (`event.source = Remote`) ものかを示します。

## <a name="register-an-event-handler"></a>イベント ハンドラーの登録

次のコード例では、ワークシートの `onChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録します。 このコードでは、そのワークシートでデータが変更されたときに、`handleChange` 関数を実行するように指定しています。

```js
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>イベントの処理

前の例で示したように、イベント ハンドラーの登録時には、特定のイベントが発生したときに実行する関数を指定します。 その関数は、目的のシナリオに必要なアクションを実行するように設計できます。 次のコード例は、イベントに関する情報を単にコンソールに出力するイベント ハンドラー関数を示しています。

```js
async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>イベント ハンドラーを削除する

次のコード例では、ワークシートの `onSelectionChanged` イベントに対応するイベント ハンドラーを **Sample** という名前で登録して、そのイベントの発生時に実行される `handleSelectionChange` 関数を定義しています。 また、そのイベント ハンドラーを削除するために、後から呼び出すことができる `remove()` 関数も定義しています。 イベント ハンドラーを削除 `RequestContext` するには、イベント ハンドラーの作成に使用する必要があります。

```js
let eventResult;

async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  await Excel.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## <a name="enable-and-disable-events"></a>イベントの有効化と無効化

イベントを無効にすると、アドインのパフォーマンスが向上する可能性があります。
たとえば、アプリがイベントを受信する必要がないことや、複数エンティティの一括編集を実行中にイベントを無視できるすることがあります。

イベントは[ランタイム](/javascript/api/excel/excel.runtime) レベルで有効または無効にできます。
`enableEvents` プロパティは、イベントが発生したかどうかと、イベント ハンドラーがアクティブになったかどうかを判別します。

次のコード サンプルは、イベントのオンとオフを切り替える方法を示しています。

```js
await Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();

    let eventBoolean = !context.runtime.enableEvents;
    context.runtime.enableEvents = eventBoolean;
    if (eventBoolean) {
        console.log("Events are currently on.");
    } else {
        console.log("Events are currently off.");
    }
    
    await context.sync();
});
```

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
