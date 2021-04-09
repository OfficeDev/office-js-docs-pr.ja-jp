---
title: Excel JavaScript API を使用して日付を使用する
description: 日付をMoment-MSDateには、Excel JavaScript API を使用してプラグインを使用します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3f59e5daad042541bd933fb4e644d40f27a6e5e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652934"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>Excel JavaScript API とプラグインを使用して日付をMoment-MSDateする

この記事では、Excel JavaScript API と [Moment-MSDate](https://www.npmjs.com/package/moment-msdate)プラグインを使用して日付を処理する方法を示すコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、Excel.Range クラスを参照してください](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>日付をMoment-MSDateするには、このプラグインを使用する

[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。 [Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。 これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。

次のコードは **、B4** の範囲をモーメントのタイムスタンプに設定する方法を示しています。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

次のコード サンプルは、セルから日付を取得し、その日付を他の形式に変換する同様の `Moment` 手法を示しています。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

ユーザーが読み取り可能な形式で日付を表示するには、アドインで範囲の書式を設定する必要があります。 たとえば `"[$-409]m/d/yy h:mm AM/PM;@"` 、「12/3/18 3:57 PM」と表示されます。 日付と時刻の形式の詳細については、「数値書式のカスタマイズに関するガイドライン」の「日付と時刻の書式に関するガイドライン」 [を参照](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) してください。


## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用してセルを使用する](excel-add-ins-cells.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
