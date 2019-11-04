---
title: 一般的なコーディングの問題と予期しないプラットフォームの動作
description: 開発者がよく遭遇する Office JavaScript API プラットフォームの問題の一覧です。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924795"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>一般的なコーディングの問題と予期しないプラットフォームの動作

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

## <a name="some-properties-must-be-set-with-json-structs"></a>一部のプロパティは、JSON 構造体で設定する必要があります。

> [!NOTE]
> このセクションは、Excel および Word のホスト固有の Api にのみ適用されます。

一部のプロパティは、個々のサブプロパティを設定するのではなく、JSON 構造体として設定する必要があります。 この例の1つは、 [PageLayout](/javascript/api/excel/excel.pagelayout)にあります。 この`zoom`プロパティは、次に示すように、1つの[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)オブジェクトで設定する必要があります。

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

前の例では、値`zoom` `sheet.pageLayout.zoom.scale = 200;`を直接割り当てることはでき***ません***。 が読み込まれてい`zoom`ないため、このステートメントはエラーをスローします。 ロードさ`zoom`れた場合でも、スケールのセットは有効になりません。 すべての`zoom`コンテキスト操作が行われ、アドイン内のプロキシオブジェクトが更新され、ローカルに設定された値が上書きされます。

この動作は、[範囲形式](/javascript/api/excel/excel.range#format)などの[ナビゲーションプロパティ](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)とは異なります。 の`format`プロパティは、次に示すように、object ナビゲーションを使用して設定できます。

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

読み取り専用修飾子をチェックすることで、そのサブプロパティを JSON 構造体で設定する必要があるプロパティを識別できます。 読み取り専用のプロパティは、読み取り専用でないサブプロパティを直接設定することができます。 書き込み可能な`PageLayout.zoom`プロパティは、JSON 構造体で設定する必要があります。 概要:

- 読み取り専用プロパティ: サブプロパティは、ナビゲーションを使用して設定できます。
- 書き込み可能なプロパティ: サブプロパティは JSON 構造体で設定する必要があります (ナビゲーションで設定することはできません)。

## <a name="excel-range-limits"></a>Excel 範囲制限

範囲を使用する Excel アドインを作成している場合は、次のサイズ制限に注意してください。

- Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。 その制限を超えると、`RichAPI.Error` がスローされます。
- 範囲は、設定操作では500万のセルに制限されます。

ユーザー入力がこれらの制限を超えていることが予想される場合は、データをチェックして、範囲を複数のオブジェクトに分割します。 また、複数`context.sync()`の呼び出しを送信して、より小さな範囲の操作が同時に一括されないようにする必要もあります。

アドインでは、範囲内のセルを戦略的に更新するために[Rangeareas](/javascript/api/excel/excel.rangeareas)を使用できる場合があります。 詳細については、「 [Excel アドインで複数の範囲を同時に操作](../excel/excel-add-ins-multiple-ranges.md)する」を参照してください。

## <a name="setting-read-only-properties"></a>読み取り専用プロパティの設定

Office JS の[TypeScript 定義](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)は、読み取り専用のオブジェクトプロパティを指定します。 読み取り専用プロパティを設定しようとすると、エラーがスローされずに書き込み操作が失敗します。 次の例では、誤って読み取り専用プロパティ[Chart.id](/javascript/api/excel/excel.chart#id)を設定しようとしています。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>関連項目

- [Officedev/office-js](https://github.com/OfficeDev/office-js/issues): office アドインプラットフォームおよび JavaScript api の問題を報告および表示する場所です。
- [スタックオーバーフロー](https://stackoverflow.com/questions/tagged/office-js): Office JavaScript api に関するプログラミング上の問題を確認および表示する場所です。 スタックオーバーフローに投稿するときには、必ず "office-js" タグを質問に適用してください。
- [UserVoice](https://officespdev.uservoice.com/): office アドインプラットフォームおよび Office JavaScript api の新機能を提案する場所です。
