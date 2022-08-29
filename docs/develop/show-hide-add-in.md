---
title: Office アドインの作業ウィンドウを表示または非表示にする
description: アドインの継続的な実行中に、プログラムによってアドインのユーザー インターフェイスを非表示または表示する方法について説明します。
ms.date: 08/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8122282414fcc9472fc300acd07da354d5a282f0
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422890"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Office アドインの作業ウィンドウを表示または非表示にする

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

メソッドを呼び出すことで、Office アドインの作業ウィンドウを `Office.addin.showAsTaskpane()` 表示できます。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

前のコードでは、 **CurrentQuarterSales** という名前の Excel ワークシートがあるシナリオを前提としています。 アドインは、このワークシートがアクティブ化されるたびに作業ウィンドウを表示します。 このメソッド `onCurrentQuarter` は、ワークシートに登録されている [Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) イベントのハンドラーです。

また、メソッドを呼び出して作業ウィンドウを `Office.addin.hide()` 非表示にすることもできます。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前のコードは、 [Office.Worksheet.onDeactivated イベントに](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) 登録されているハンドラーです。

## <a name="additional-details-on-showing-the-task-pane"></a>作業ウィンドウの表示に関するその他の詳細

呼び出 `Office.addin.showAsTaskpane()`すと、作業ウィンドウのリソース ID (`resid`) 値として割り当てたファイルが作業ウィンドウに表示されます。 この`resid`値は、**manifest.xml** ファイルを開いて要素内`<Action xsi:type="ShowTaskpane">`を **\<SourceLocation\>** 特定することで、割り当てまたは変更できます。
(詳細については、「 [共有ランタイムを使用するように Office アドインを構成する](configure-your-add-in-to-use-a-shared-runtime.md) 」を参照してください)。

非同期メソッドであるため `Office.addin.showAsTaskpane()` 、メソッドが完了するまでコードの実行は続行されます。 使用している JavaScript 構文に応じて、キーワードまたは`then()`メソッドで`await`この完了を待ちます。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

メソッドを`showAsTaskpane()``hide()`使用するには、アドインで[共有ランタイム](../testing/runtimes.md#shared-runtime)を使用する必要があります。 詳細については、「 [共有ランタイムを使用するように Office アドインを構成する](configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="preservation-of-state-and-event-listeners"></a>状態リスナーとイベント リスナーの保持

およびメソッドは `hide()` 、 `showAsTaskpane()` 作業ウィンドウの *可視性* のみを変更します。 アンロードまたは再読み込み (またはその状態の再初期化) は行われません。

次のシナリオを考慮してください。作業ウィンドウはタブで設計されています。 アドインが最初に起動されると、[ **ホーム** ] タブが開きます。 ユーザーが **[設定]** タブを開き、後で作業ウィンドウのコードが何らかのイベントに応答して呼び出 `hide()` されたとします。 別のイベントに応答して、さらに後のコード呼び出し `showAsTaskpane()` 。 作業ウィンドウが再表示され、[ **設定]** タブは引き続き選択されます。

![[ホーム]、[設定]、[お気に入り]、および [アカウント] という 4 つのタブがある作業ウィンドウ。](../images/TaskpaneWithTabs.png)

さらに、作業ウィンドウに登録されているイベント リスナーは、作業ウィンドウが非表示になっている場合でも引き続き実行されます。

次のシナリオを考慮してください。作業ウィンドウには、Excel `Worksheet.onActivated` の登録済みハンドラーと`Worksheet.onDeactivated`**、Sheet1** という名前のシートのイベントがあります。 アクティブ化されたハンドラーにより、作業ウィンドウに緑色のドットが表示されます。 非アクティブ化されたハンドラーは、ドットの赤 (既定の状態) を回します。 次に、**Sheet1** がアクティブ化されておらず、ドットが赤の場合にコードが呼び出`hide()`されるとします。 作業ウィンドウが非表示になっている間、 **シート 1** がアクティブ化されます。 後のコードは、いくつかのイベントに応答して呼び出されます `showAsTaskpane()` 。 作業ウィンドウが開くと、作業ウィンドウが非表示であってもイベント リスナーとハンドラーが実行されたため、ドットは緑色になります。

## <a name="handle-the-visibility-changed-event"></a>可視性が変更されたイベントを処理する

コードで作業ウィンドウ`showAsTaskpane()``hide()`の表示を変更すると、Office によってイベントが`VisibilityModeChanged`トリガーされます。 このイベントを処理すると便利です。 たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されたとします。 作業ウィンドウが非表示の間に新しいワークシートが追加された場合、作業ウィンドウを表示することは、それ自体では、新しいワークシート名をリストに追加しません。 ただし、次の`VisibilityModeChanged`コード例に示すように、コードはイベントに応答して[、Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) コレクション内のすべてのワークシートの [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) プロパティを再読み込みできます。

イベントのハンドラーを登録するには、ほとんどの Office JavaScript コンテキストと同様に "ハンドラーの追加" メソッドを使用しません。 代わりに、ハンドラー ( [Office.addin.onVisibilityModeChanged) を](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1))渡す特別な関数があります。 次に例を示します。 プロパティは `args.visibilityMode` [VisibilityMode](/javascript/api/office/office.visibilitymode) 型であることに注意してください。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode == "Taskpane") {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

この関数は、ハンドラーの *登録を解除* する別の関数を返します。 単純な堅牢な例を次に示します。

```javascript
const removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode == "Taskpane") {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

このメソッドは `onVisibilityModeChanged` 非同期であり、promise を返します。つまり、 **コードは、登録解除** ハンドラーを呼び出す前に、Promise の履行を待機する必要があることを意味します。

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
const removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode == "Taskpane") {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

登録解除関数も非同期であり、Promise を返します。 そのため、登録解除が完了するまで実行しないコードがある場合は、登録解除関数によって返される Promise を待機する必要があります。

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>関連項目

- [共有ランタイムを使用するように Office アドインを構成する](configure-your-add-in-to-use-a-shared-runtime.md)
- [ドキュメントが開いたら、Office アドインでコードを実行する](run-code-on-document-open.md)
