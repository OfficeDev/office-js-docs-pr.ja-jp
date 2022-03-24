---
title: JavaScript API を使用して日付Excelする
description: 日付をMoment-MSDateするには、JavaScript API Excelプラグインを使用します。
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7ca6e0eacab7aab0308b2e397f313a8e07b59777
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745074"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>JavaScript API と Excelプラグインを使用して日付Moment-MSDate作業する

この記事では、JavaScript API と [Moment-MSDate](https://www.npmjs.com/package/moment-msdate) プラグインを使用して日付をExcelする方法を示すコード サンプルを提供します。 オブジェクトがサポートするプロパティとメソッドの`Range`完全な一覧については、次のExcel[。Range クラス](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>日付をMoment-MSDateする場合は、このプラグインを使用します。

[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。 [Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。 これは、[NOW 関数](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。

次のコードは、 **B4** の範囲をモーメントのタイムスタンプに設定する方法を示しています。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let now = Date.now();
    let nowMoment = moment(now);
    let nowMS = nowMoment.toOADate();

    let dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    await context.sync();
});
```

次のコード サンプルは、 `Moment` セルから日付を取得し、その日付を他の形式に変換する同様の手法を示しています。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let dateRange = sheet.getRange("B4");
    dateRange.load("values");

    await context.sync();

    let nowMS = dateRange.values[0][0];

    // Log the date as a moment.
    let nowMoment = moment.fromOADate(nowMS);
    console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

    // Log the date as a UNIX-style timestamp.
    let now = nowMoment.unix();
    console.log(`get (timestamp): ${now}`);
});
```

ユーザーが読み取り可能な形式で日付を表示するには、アドインで範囲の書式を設定する必要があります。 たとえば、「 `"[$-409]m/d/yy h:mm AM/PM;@"` 12/3/18 3:57 PM」と表示されます。 日付と時刻の形式の詳細については、「数値書式のカスタマイズに関するガイドライン」の「日付と時刻の書式に関するガイドライン」 [を参照](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) してください。

## <a name="see-also"></a>関連項目

- [JavaScript API を使用してセルExcelする](excel-add-ins-cells.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)
