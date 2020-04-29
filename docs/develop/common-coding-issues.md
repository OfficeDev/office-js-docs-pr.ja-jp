---
title: 一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: dea879899dce2e957d34f2eb8e7498d4fdb868c0
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930317"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a>一般的な問題と予期しないプラットフォームの動作に関するコーディングガイダンス

この記事では、予期しない動作が発生するか、必要な結果を得るために特定のコーディングパターンが必要になる可能性がある Office JavaScript API の側面について説明します。 このリストに含まれる問題が発生した場合は、記事の下部にあるフィードバックフォームを使用してお知らせください。

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>一般的な Api と Outlook Api は、約束に基づくものではありません

[共通 api](/javascript/api/office) (特定の Office ホストに縛られていないもの) と[Outlook api](/javascript/api/outlook)では、コールバックベースのプログラミングモデルが使用されます。 基になる Office ドキュメントと対話するには、操作が完了したときに実行されるコールバックを指定する非同期の読み取りまたは書き込みの呼び出しが必要です。 このパターンの例については、「 [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)」を参照してください。

これらの共通 API および Outlook API メソッドは、[約束](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)を返しません。 そのため、[待機](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)を使用して、非同期操作が完了するまで実行を一時停止することはできません。 振る舞いが必要`await`な場合は、明示的に作成した約束でメソッドの呼び出しをラップすることができます。

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> 参照ドキュメントには、 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)の Promise ラップによる実装が含まれています。

## <a name="some-properties-cannot-be-set-directly"></a>一部のプロパティは直接設定できません

> [!NOTE]
> このセクションは、Excel および Word のホスト固有の Api にのみ適用されます。

書き込み可能であっても、一部のプロパティを設定することはできません。 これらのプロパティは、1つのオブジェクトとして設定する必要がある親プロパティの一部です。 これは、親プロパティが特定の論理的な関係を持つサブプロパティに依存しているためです。 このような親プロパティは、オブジェクトの個々のサブプロパティを設定するのではなく、オブジェクトリテラル表記を使用して設定し、オブジェクト全体を設定する必要があります。 この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。 この`zoom`プロパティは、次に示すように、1つの[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクトで設定する必要があります。

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

前の例では、値`zoom` `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることはでき***ません***。 が読み込まれてい`zoom`ないため、このステートメントはエラーをスローします。 ロードさ`zoom`れた場合でも、スケールのセットは有効になりません。 すべての`zoom`コンテキスト操作が行われ、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。

この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)とは異なります。 の`format`プロパティは、次に示すように、object ナビゲーションを使用して設定できます。

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

読み取り専用修飾子をチェックすることによって、そのサブプロパティを直接設定できないプロパティを識別できます。 読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。 書き込み可能な`PageLayout.zoom`プロパティは、そのレベルのオブジェクトで設定する必要があります。 概要:

- 読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。
- 書き込み可能なプロパティ: ナビゲーションを使用してサブプロパティを設定することはできません (最初の親オブジェクトの割り当ての一部として設定する必要があります)。

## <a name="excel-data-transfer-limits"></a>Excel データ転送の制限

Excel アドインを作成している場合は、ブックを操作するときに以下のサイズ制限に注意してください。

- Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。 その制限を超えると、`RichAPI.Error` がスローされます。
- 範囲は、取得操作に500万のセルに制限されます。

ユーザー入力がこれらの制限を超えていることが予想される場合は、 `context.sync()`必ずデータを確認してから、を呼び出してください。 必要に応じて、操作を小さな部分に分割します。 各サブ操作を`context.sync()`呼び出して、それらの操作が再度一括されないようにしてください。

これらの制限は、通常、大きな範囲を超えています。 アドインでは、範囲内のセルを戦略的に更新するために[Rangeareas](/javascript/api/excel/excel.rangeareas)を使用できる場合があります。 詳細については、「 [Excel アドインで複数の範囲を同時に操作](../excel/excel-add-ins-multiple-ranges.md)する」を参照してください。

## <a name="setting-read-only-properties"></a>読み取り専用プロパティの設定

Office JS の[TypeScript 定義](referencing-the-javascript-api-for-office-library-from-its-cdn.md)は、読み取り専用のオブジェクトプロパティを指定します。 読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。 次の例では、誤って読み取り専用プロパティ[Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a>イベントハンドラーの削除

イベントハンドラーは、追加したもの`RequestContext`と同じものを使用して削除する必要があります。 実行中にアドインでイベントハンドラーを削除する必要がある場合は、ハンドラーを追加するために使用されるコンテキストオブジェクトを格納する必要があります。

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a>Internet Explorer のサポート

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a>関連項目

- [Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。
- [スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。 スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。
- [UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。
