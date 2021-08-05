---
title: Office アドインの作業ウィンドウを表示または非表示にする
description: 継続的に実行されている間に、アドインのユーザー インターフェイスをプログラムで非表示または表示する方法について説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: b2f0efa95f4ce71fc73d9834cfc165cfdd85dc8f
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773756"
---
# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Office アドインの作業ウィンドウを表示または非表示にする

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

前のコードでは **、CurrentQuarterSales** という名前のワークシートExcelシナリオを想定しています。 このワークシートがアクティブ化されるたびに、アドインによって作業ウィンドウが表示されます。 メソッドは `onCurrentQuarter` 、メソッドのハンドラー [Office。ワークシートに登録されている Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onActivated)イベント。

関数を呼び出して作業ウィンドウを非表示 `Office.addin.hide()` にすることもできます。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前のコードは、アプリケーションに登録されているハンドラー [Office。Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#onDeactivated)イベント。

## <a name="additional-details-on-showing-the-task-pane"></a>作業ウィンドウの表示に関するその他の詳細

呼び出す場合Office作業ウィンドウのリソース ID ( ) 値として割り当てたファイル `Office.addin.showAsTaskpane()` `resid` が作業ウィンドウに表示されます。 この `resid` 値を割り当てたり変更したりするには、manifest.xmlファイルを開き、要素 `<SourceLocation>` 内を検索 `<Action xsi:type="ShowTaskpane">` します。
(詳細[については、「Office共有ランタイムを使用するアドインの構成」](configure-your-add-in-to-use-a-shared-runtime.md)を参照してください。

非同期 `Office.addin.showAsTaskpane()` メソッドであるから、関数が完了するまでコードは実行を続ける。 使用している JavaScript 構文に応じて、キーワードまたはメソッドでこの完了 `await` `then()` を待ちます。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

and メソッド `showAsTaskpane()` を `hide()` 使用するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「共有ランタイム[を使用Officeアドインを構成する」を参照してください](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="preservation-of-state-and-event-listeners"></a>状態リスナーとイベント リスナーの保持

and `hide()` メソッド `showAsTaskpane()` は、作業ウィンドウ *の表示* 設定のみを変更します。 アンロードまたは再読み込み (または状態の再初期化) は行ないます。

次のシナリオについて考えます。作業ウィンドウはタブで設計されています。 アドイン **が** 最初に起動すると、[ホーム] タブが開きます。 たとえば、ユーザーが **[設定]** タブを開き、後で作業ウィンドウ内のコードが何らかのイベントに応答 `hide()` して呼び出されたとします。 別のイベントに応答 `showAsTaskpane()` して以降のコード呼び出し。 作業ウィンドウが再表示され、引き **続き**[設定] タブが選択されます。

![[ホーム]、[お気に入り]、および [アカウント] という 4 つのタブ設定作業ウィンドウのスクリーンショットです。](../images/TaskpaneWithTabs.png)

さらに、作業ウィンドウに登録されているイベント リスナーは、作業ウィンドウが非表示の場合でも引き続き実行されます。

次のシナリオを検討してください。 作業ウィンドウには、Sheet1 という名前のシートのExcelイベントの登録 `Worksheet.onActivated` `Worksheet.onDeactivated` されたハンドラー **があります**。 アクティブ化されたハンドラーによって、作業ウィンドウに緑色のドットが表示されます。 非アクティブ化されたハンドラーは、ドットを赤色 (既定の状態) に変える。 次に、シート `hide()` **1** がアクティブ化されていないときにコードが呼び出され、ドットが赤になっているとします。 作業ウィンドウが非表示の間、 **シート 1 が** アクティブになります。 以降のコードは `showAsTaskpane()` 、いくつかのイベントに応答して呼び出します。 作業ウィンドウが開くと、作業ウィンドウが非表示の場合でもイベント リスナーとハンドラーが実行されたため、ドットは緑色になります。

## <a name="handle-the-visibility-changed-event"></a>表示が変更されたイベントを処理する

コードで作業ウィンドウの表示設定を変更すると、イベント `showAsTaskpane()` Office `hide()` トリガー `VisibilityModeChanged` されます。 このイベントを処理すると便利です。 たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されたとします。 作業ウィンドウが非表示の状態で新しいワークシートが追加された場合、作業ウィンドウを表示すると、それ自体で新しいワークシート名がリストに追加されません。 ただし、以下のコード例に示すように、コードはイベントに応答して `VisibilityModeChanged` [Workbook.worksheets](/javascript/api/excel/excel.workbook#worksheets)コレクション内のすべてのワークシートの[Worksheet.name](/javascript/api/excel/excel.worksheet#name)プロパティを再読み込みできます。

イベントのハンドラーを登録するには、ほとんどの JavaScript コンテキストと同様に、"add handler" メソッドOffice使用します。 代わりに、ハンドラーを渡す特別な関数があります[:Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#onVisibilityModeChanged_listener_). 次に例を示します。 プロパティの種類 `args.visibilityMode` は [VisibilityMode です](/javascript/api/office/office.visibilitymode)。

```javascript
Office.addin.onVisibilityModeChanged(function(args) {
    if (args.visibilityMode = "Taskpane"); {
        // Code that runs whenever the task pane is made visible.
        // For example, an Excel.run() that loads the names of
        // all worksheets and passes them to the task pane UI.
    }
});
```

この関数は、ハンドラーを登録解除 *する別の関数を* 返します。 ここでは、単純ですが堅牢ではない例を示します。

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

このメソッドは非同期であり、約束を返します。つまり、コードは登録解除ハンドラーを呼び出す前に、約束の履行を待つ `onVisibilityModeChanged` **必要** があります。

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

登録解除関数も非同期であり、約束を返します。 したがって、登録解除が完了するまで実行しないコードがある場合は、登録解除関数によって返される約束を待つ必要があります。

```javascript
// await the promise from the deregister handler before continuing
await removeVisibilityModeHandler();
// subsequent code here
```

## <a name="see-also"></a>関連項目

- [Office アドインを構成して共有 JavaScript ランタイムを使用する](configure-your-add-in-to-use-a-shared-runtime.md)
- [ドキュメントが開いたら、Office アドインでコードを実行する](run-code-on-document-open.md)
