---
title: 共有ランタイムで Office アドインを表示または非表示にする
description: 連続して実行している間にプログラムでアドインのユーザーインターフェイスを表示または非表示にする方法について説明します。
ms.date: 05/17/2020
localization_priority: Normal
ms.openlocfilehash: e49c47c86a986c85ad12e09666b7ac2fb5411322
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275715"
---
# <a name="show-or-hide-an-office-add-in-in-a-shared-runtime"></a>共有ランタイムで Office アドインを表示または非表示にする

Office アドインには、次のいずれかの部分を含めることができます。

- 作業ウィンドウ
- UI レス関数ファイル (作業ウィンドウや他のユーザーインターフェイス要素を使用しないカスタム関数)
- Excel カスタム関数

既定では、各パーツは独自の独立した JavaScript ランタイムで実行され、独自のグローバルオブジェクトとグローバル変数を持ちます。

2つ以上のパーツを含むアドインは、共通の JavaScript ランタイムを共有できます。 この共有ランタイム機能を使用すると、アドインの実行中に作業ウィンドウを非表示にし、再び開くことができる新しい Api が有効になります。

## <a name="configure-an-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するようにアドインを構成する

共有ランタイムを使用するようにアドインを構成するには、「[共有ランタイムを使用するように Office アドインを構成](configure-your-add-in-to-use-a-shared-runtime.md)する」を参照してください。

## <a name="show-and-hide-the-task-pane"></a>作業ウィンドウを表示または非表示にする

新しい Api は `Office.addin` プロパティにあります。 作業ウィンドウを表示するには、コードを呼び出し `Office.addin.showAsTaskpane()` ます。 Office は、作業ウィンドウのリソース ID () に割り当てたページを作業ウィンドウに表示し `resid` ます。 これは、 `resid` マニフェスト内のにに割り当てられたです `<SourceLocation>` `<Action xsi:type="ShowTaskpane">` 。 (「[共有ランタイムを使用するために Office アドインを構成する」を](configure-your-add-in-to-use-a-shared-runtime.md)参照してください)。

これは非同期メソッドなので、完了するまで後続のコードが実行されないように、コードで待機する必要があります。 この完了は `await` `then()` 、使用している JavaScript 構文に応じて、キーワードまたはメソッドのいずれかを使用して待機します。 次の例では、 **CurrentQuarterSales**という名前の Excel ワークシートが存在することを前提としています。 このワークシートがアクティブになると、アドインによって作業ウィンドウが表示されるようになります。 このメソッド `onCurrentQuarter` は、ワークシートに登録されている、 [onactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview#onactivated)イベントのハンドラーです。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

作業ウィンドウを非表示にするには、コードを呼び出し `Office.addin.hide()` ます。 次の例は、 [onDeactivated アクティブ](/javascript/api/excel/excel.worksheet?view=excel-js-preview#ondeactivated)化イベントに登録されているハンドラーです。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

### <a name="preservation-of-state-and-event-listeners"></a>状態およびイベントリスナーの保持

`hide()`メソッドと `showAsTaskpane()` メソッドは、作業ウィンドウの*表示状態*のみを変更します。 アンロードまたは再ロードしたり、その状態を再初期化したりすることはありません。

次のシナリオを考えてみます。作業ウィンドウは、タブで設計されています。 [**ホーム**] タブは、アドインを最初に起動したときに開かれています。 ユーザーが [**設定**] タブを開き、後で、あるイベントに応答して、作業ウィンドウの呼び出しのコードを開くとし `hide()` ます。 他のイベントに応答して、後でコードを呼び出すこと `showAsTaskpane()` ができます。 作業ウィンドウが再度表示され、[**設定**] タブが選択されたままになります。

![[ホーム]、[設定]、[お気に入り]、および [アカウント] というラベルの付いた4つのタブがある作業ウィンドウのスクリーンショット。](../images/TaskpaneWithTabs.png)

さらに、作業ウィンドウに登録されているイベントリスナーは、作業ウィンドウが非表示になっている場合でも、引き続き実行されます。

次のシナリオを考えます。この作業ウィンドウには、 `Worksheet.onActivated` Sheet1 というシートの Excel およびイベントのハンドラーが登録されてい `Worksheet.onDeactivated` ます。 **Sheet1** アクティブ化されたハンドラーによって、作業ウィンドウに緑の点が表示されます。 非アクティブ化されたハンドラーは、ドット red (これは既定の状態) をオフにします。 `hide()` **Sheet1**がアクティブ化されておらず、ドットが赤の場合は、コードが呼び出されるとします。 作業ウィンドウは非表示になっていますが、 **Sheet1**がアクティブになります。 `showAsTaskpane()`イベントに応答して、後でコードを呼び出すことができます。 作業ウィンドウが開くと、その作業ウィンドウが非表示になっているにもかかわらず、イベントリスナーとハンドラーが実行されるため、ドットは緑になります。

### <a name="handle-visibility-changed-event"></a>可視性の変更イベントを処理する

コードによって作業ウィンドウの表示がまたはに変更されると、Office によって `showAsTaskpane()` `hide()` イベントがトリガーさ `VisibilityModeChanged` れます。 このイベントを処理すると便利な場合があります。 たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されているとします。 作業ウィンドウが非表示になっているときに新しいワークシートが追加されても、その作業ウィンドウが表示されないようにするには、リストに新しいワークシート名を追加します。 しかし、 `VisibilityModeChanged` 以下のコード例に示されているように、コードでイベントに応答して、 [Worksheet.name](/javascript/api/excel/excel.worksheet#name)コレクション内のすべてのワークシートのプロパティを再読み込みすることが[できます。](/javascript/api/excel/excel.workbook#worksheets)

イベントのハンドラーを登録するには、ほとんどの Office JavaScript コンテキストでの "add handler" メソッドは使用しません。 代わりに、ハンドラーを渡すための特殊な関数[onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)が用意されています。 次に例を示します。 このプロパティの `args.visibilityMode` 型は[VisibilityMode](/javascript/api/office/office.visibilitymode)であることに注意してください。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

この関数は、ハンドラーを*deregisters*する別の関数を返します。 この例は、単純ですが、堅牢ではありません。

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

この `onVisibilityModeChanged` メソッドは非同期です。つまり、コードから返される*登録*解除ハンドラーを呼び出す場合は、登録解除ハンドラーを呼び出す前に、が完了していることを `onVisibilityModeChanged` 確認する必要があり `onVisibilityModeChanged` ます。 そのための1つの方法は、 `await` 次の例のように、メソッド呼び出しでキーワードを使用することです。

```javascript
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

ES2015 の JavaScript のみを使用する場合は、次の例に示すように、コードでメソッドを使用して、 `then` 返された Promise オブジェクトが解決されるまで待機し、返された関数をグローバル変数に代入することができます。

```javascript
var removeVisibilityModeHandler;

Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
}).then(function(removeHandler) {
        removeVisibilityModeHandler = removeHandler;
    });

// In some later code path, deregister with:
removeVisibilityModeHandler();
```

登録解除関数は、それ自体が非同期です。 そのため、登録解除の完了後に実行してはならないコードがある場合は、次の例に示すように、登録解除関数を `await` キーワードまたはメソッドで待機する必要があり `then` ます。

ハンドラーを登録解除するには、次のようにします。

```javascript
await removeVisibilityModeHandler();
// subsequent code here

// or use pre-ES2015 syntax:
removeVisibilityModeHandler().then(function () {
        // subsequent code here
    })
```
