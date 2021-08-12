---
title: JavaScript API を使用して日付Excelする
description: 日付をMoment-MSDateするには、JavaScript API Excelプラグインを使用します。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdfc39f12b3374d9903156b1ba71a9bbd4f296735f0ed41dac56d62243058c1d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084733"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>JavaScript API と Excelプラグインを使用して日付Moment-MSDate作業する

この記事では、JavaScript API と[Moment-MSDate](https://www.npmjs.com/package/moment-msdate)プラグインを使用して日付Excelする方法を示すコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの完全な一覧については、 `Range` 次の[Excel。Range クラス](/javascript/api/excel/excel.range)。

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

- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
