---
title: アドインの作業ウィンドウを表示Office非表示にする
description: アドインが継続的に実行されている間に、プログラムによってアドインのユーザー インターフェイスを非表示または表示する方法について説明します。
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 20db609a3a6ded5624391f705dab1ad6b8f6e043
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789251"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>アドインの作業ウィンドウを表示Office非表示にする

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

関数を呼び出すことによって、Officeアドインの作業ウィンドウを表示 `Office.addin.showAsTaskpane()` できます。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

前のコードでは **、CurrentQuarterSales** という名前の Excel ワークシートがあるシナリオを想定しています。 このワークシートがアクティブ化されるたびに、アドインによって作業ウィンドウが表示されます。 このメソッド `onCurrentQuarter` は、ワークシートに登録 [されている Office.Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onactivated) イベントのハンドラーです。

関数を呼び出して作業ウィンドウを非表示 `Office.addin.hide()` にすることもできます。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前のコードは [、Office.Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#ondeactivated) イベントに登録されているハンドラーです。

## <a name="additional-details-on-showing-the-task-pane"></a>作業ウィンドウの表示に関するその他の詳細

呼び出しOffice作業ウィンドウのリソース ID ( ) 値として割り当てたファイルが作業ウィンドウ `Office.addin.showAsTaskpane()` `resid` に表示されます。 この `resid` 値は、ファイルを開き、要素内 **manifest.xml** することで、割り当 `<SourceLocation>` てまたは変更 `<Action xsi:type="ShowTaskpane">` できます。
(詳 [しくは、「共有ランタイムOffice使うアドイン](configure-your-add-in-to-use-a-shared-runtime.md) の構成」をご覧ください)。

非同期 `Office.addin.showAsTaskpane()` メソッドの場合、コードは関数が完了するまで実行を続行します。 使用している JavaScript 構文に応じて、キーワードまたはメソッドを使用してこの完了 `await` `then()` を待ちます。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

これらのメソッドを `showAsTaskpane()` `hide()` 使用するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「共有ランタイム [を使用Officeアドインを構成する」を参照してください](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="preservation-of-state-and-event-listeners"></a>状態リスナーとイベント リスナーの保持

The `hide()` and `showAsTaskpane()` methods only change the *visibility* of the task pane. アンロードまたは再読み込み (または状態の再初期化) は行されません。

次のシナリオを考えます。作業ウィンドウはタブで設計されています。 アドイン **が** 最初に起動すると、[ホーム] タブが開きます。 ユーザーが [設定] タブを開き、後で何らかのイベントに応答して作業ウィンドウの呼び出し `hide()` をコード化したとします。 別のイベントへの応答 `showAsTaskpane()` として、後でコードが呼び出されます。 作業ウィンドウが再び表示され、[設定] **タブは** 引き続き選択されます。

![[ホーム]、[設定]、[お気に入り]、および [アカウント] というラベルの付いた 4 つのタブがある作業ウィンドウのスクリーンショット。](../images/TaskpaneWithTabs.png)

また、作業ウィンドウに登録されているイベント リスナーは、作業ウィンドウが非表示の場合でも引き続き実行されます。

次のシナリオについて考えます。作業ウィンドウには、Excel の登録されたハンドラーと `Worksheet.onActivated` `Worksheet.onDeactivated` **、Sheet1** という名前のシートのイベントがあります。 アクティブ化されたハンドラーによって、作業ウィンドウに緑色のドットが表示されます。 非アクティブ化されたハンドラーは、ドットを赤色 (既定の状態) に変更します。 次に、シート `hide()` **1** がアクティブ化されていないときに、ドットが赤のときにコードが呼び出されるとします。 作業ウィンドウが非表示の間、 **シート 1 が** アクティブになります。 以降のコードは、 `showAsTaskpane()` 何らかのイベントに応答して呼び出します。 作業ウィンドウが開くと、作業ウィンドウが非表示でもイベント リスナーとハンドラーが実行されたため、ドットは緑色になります。

## <a name="handle-the-visibility-changed-event"></a>可視性の変更イベントを処理する

コードで作業ウィンドウの表示/非表示を変更すると、イベントOffice `showAsTaskpane()` `hide()` トリガー `VisibilityModeChanged` されます。 このイベントを処理すると便利です。 たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されたとします。 作業ウィンドウが非表示の間に新しいワークシートが追加された場合、作業ウィンドウを表示すると、それ自体は新しいワークシート名をリストに追加しません。 ただし、次のコード例に示すように、コードはイベントに応答して `VisibilityModeChanged` [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets)コレクション内のすべてのワークシートの[Worksheet.name](/javascript/api/excel/excel.worksheet#name)プロパティを再読み込みできます。

イベントのハンドラーを登録するには、JavaScript コンテキストのほとんどの場合のように、"ハンドラーの追加" メソッドOffice使用します。 代わりに、ハンドラーを渡す特別な関数 [Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onvisibilitymodechanged-listener-)があります。 次に例を示します。 プロパティは `args.visibilityMode` [VisibilityMode 型です](/javascript/api/office/office.visibilitymode)。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

この関数は、ハンドラーを登録解除 *する別の関数を* 返します。 次に、シンプルですが堅牢ではない例を示します。

```javascript
var removeVisibilityModeHandler =
    Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });


// In some later code path, deregister with:
removeVisibilityModeHandler();
```

メソッドは非同期であり、promise を返します。つまり、登録解除ハンドラーを呼び出す前に、コードで promise のフルフィルメントを待 `onVisibilityModeChanged` **つ必要** があります。

```javascript
// await the promise from onVisibilityModeChanged and assign
// the returned deregister handler to removeVisibilityModeHandler.
var removeVisibilityModeHandler =
    await Office.addin.onVisibilityModeChanged(function(args) {
        if (args.visibilityMode = "Taskpane"); {
            // Code that runs whenever the task pane is made visible.
        }
    });
```

登録解除関数も非同期で、promise を返します。 したがって、登録解除が完了するまで実行しないコードがある場合は、登録解除関数によって返される promise を待つ必要があります。

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>関連項目

- [共有 JavaScript Office使用する新しいアドインを構成する](configure-your-add-in-to-use-a-shared-runtime.md)
- [ドキュメントが開Officeアドインでコードを実行する](run-code-on-document-open.md)
