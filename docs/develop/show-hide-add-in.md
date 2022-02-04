---
title: Office アドインの作業ウィンドウを表示または非表示にする
description: 継続的に実行されている間に、アドインのユーザー インターフェイスをプログラムで非表示または表示する方法について説明します。
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# <a name="show-or-hide-the-task-pane-of-your-office-add-in"></a>Office アドインの作業ウィンドウを表示または非表示にする

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

関数を呼び出すことによって、Officeアドインの作業ウィンドウを表示`Office.addin.showAsTaskpane()`できます。

```javascript
function onCurrentQuarter() {
    Office.addin.showAsTaskpane()
    .then(function() {
        // Code that enables task pane UI elements for
        // working with the current quarter.
    });
}
```

前のコードでは、**CurrentQuarterSales** という名前のExcelシナリオを想定しています。 このワークシートがアクティブ化されるたびに、アドインによって作業ウィンドウが表示されます。 メソッドは`onCurrentQuarter`、メソッドのハンドラー [Office。ワークシートに登録されている Worksheet.onActivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-onactivated-member) イベント。

関数を呼び出して作業ウィンドウを非表示に `Office.addin.hide()` することもできます。

```javascript
function onCurrentQuarterDeactivated() {
    Office.addin.hide();
}
```

前のコードは、アプリケーションに登録されている[ハンドラー Office。Worksheet.onDeactivated](/javascript/api/excel/excel.worksheet?view=excel-js-preview&preserve-view=true#excel-excel-worksheet-ondeactivated-member) イベント。

## <a name="additional-details-on-showing-the-task-pane"></a>作業ウィンドウの表示に関するその他の詳細

呼び出す`Office.addin.showAsTaskpane()`場合Office作業ウィンドウのリソース ID (`resid`) 値として割り当てたファイルが作業ウィンドウに表示されます。 この `resid` 値を割り当てたり変更したりするには、 **manifest.xmlファイル** `<SourceLocation>` を開き、要素内を検索 `<Action xsi:type="ShowTaskpane">` します。
(詳細[については、「Office共有](configure-your-add-in-to-use-a-shared-runtime.md)ランタイムを使用するアドインの構成」を参照してください。

非同期 `Office.addin.showAsTaskpane()` メソッドであるから、関数が完了するまでコードは実行を続ける。 使用している JavaScript 構文`await``then()`に応じて、キーワードまたはメソッドでこの完了を待ちます。

## <a name="configure-your-add-in-to-use-the-shared-runtime"></a>共有ランタイムを使用するアドインを構成する

and メソッドを `showAsTaskpane()` 使用 `hide()` するには、アドインで共有ランタイムを使用する必要があります。 詳細については、「共有ランタイム[を使用Officeアドインを構成する」を参照してください](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="preservation-of-state-and-event-listeners"></a>状態リスナーとイベント リスナーの保持

and `hide()` メソッド `showAsTaskpane()` は、作業ウィンドウ *の表示* 設定のみを変更します。 アンロードまたは再読み込み (または状態の再初期化) は行ないます。

次のシナリオについて考えます。作業ウィンドウはタブで設計されています。 アドイン **が** 最初に起動すると、[ホーム] タブが開きます。 たとえば、ユーザーが [**設定]** `hide()` タブを開き、後で作業ウィンドウ内のコードが何らかのイベントに応答して呼び出されたとします。 別のイベントに応答して `showAsTaskpane()` 以降のコード呼び出し。 作業ウィンドウが再表示され、引き **続** き [設定] タブが選択されます。

![[ホーム]、[お気に入り]、および [アカウント] という 4 つのタブ設定作業ウィンドウのスクリーンショットです。](../images/TaskpaneWithTabs.png)

さらに、作業ウィンドウに登録されているイベント リスナーは、作業ウィンドウが非表示の場合でも引き続き実行されます。

次のシナリオを検討してください。 作業`Worksheet.onActivated`ウィンドウには、シート 1 という名前のシートExcelイベント`Worksheet.onDeactivated`の登録されたハンドラー **があります**。 アクティブ化されたハンドラーによって、作業ウィンドウに緑色のドットが表示されます。 非アクティブ化されたハンドラーは、ドットを赤色 (既定の状態) に変える。 次に、**シート 1** が`hide()`アクティブ化されていないときにコードが呼び出され、ドットが赤になっているとします。 作業ウィンドウが非表示の間、 **シート 1 が** アクティブになります。 以降のコードは、 `showAsTaskpane()` いくつかのイベントに応答して呼び出します。 作業ウィンドウが開くと、作業ウィンドウが非表示の場合でもイベント リスナーとハンドラーが実行されたため、ドットは緑色になります。

## <a name="handle-the-visibility-changed-event"></a>表示が変更されたイベントを処理する

コードで作業ウィンドウの表示`showAsTaskpane()``hide()``VisibilityModeChanged`設定を変更すると、イベントOfficeトリガーされます。 このイベントを処理すると便利です。 たとえば、作業ウィンドウにブック内のすべてのシートの一覧が表示されたとします。 作業ウィンドウが非表示の状態で新しいワークシートが追加された場合、作業ウィンドウを表示すると、それ自体で新しいワークシート名がリストに追加されません。 ただし、以下`VisibilityModeChanged`のコード例に示すように、コードはイベントに応答して [Workbook.worksheets](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member) コレクション内のすべてのワークシートの [Worksheet.name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member) プロパティを再読み込みできます。

イベントのハンドラーを登録するには、ほとんどの JavaScript コンテキストと同様に、"add handler" メソッドOffice使用します。 代わりに、ハンドラーを渡す特別な関数があります:[Office.addin.onVisibilityModeChanged](/javascript/api/office/office.addin#office-office-addin-onvisibilitymodechanged-member(1))。 次に例を示します。 プロパティは [VisibilityMode](/javascript/api/office/office.visibilitymode)`args.visibilityMode` 型です。

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

この `onVisibilityModeChanged` メソッドは非同期であり、約束を返します。つまり、コードは登録解除ハンドラーを呼び出す前に、約束の履行を待つ **必要** があります。

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
