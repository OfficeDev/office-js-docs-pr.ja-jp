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
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a><span data-ttu-id="c0599-103">Excel JavaScript API とプラグインを使用して日付をMoment-MSDateする</span><span class="sxs-lookup"><span data-stu-id="c0599-103">Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in</span></span>

<span data-ttu-id="c0599-104">この記事では、Excel JavaScript API と [Moment-MSDate](https://www.npmjs.com/package/moment-msdate)プラグインを使用して日付を処理する方法を示すコード サンプルを提供します。</span><span class="sxs-lookup"><span data-stu-id="c0599-104">This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate).</span></span> <span data-ttu-id="c0599-105">オブジェクトがサポートするプロパティとメソッドの完全な一覧については `Range` [、Excel.Range クラスを参照してください](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="c0599-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a><span data-ttu-id="c0599-106">日付をMoment-MSDateするには、このプラグインを使用する</span><span class="sxs-lookup"><span data-stu-id="c0599-106">Use the Moment-MSDate plug-in to work with dates</span></span>

<span data-ttu-id="c0599-107">[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="c0599-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="c0599-108">[Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。</span><span class="sxs-lookup"><span data-stu-id="c0599-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="c0599-109">これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。</span><span class="sxs-lookup"><span data-stu-id="c0599-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="c0599-110">次のコードは **、B4** の範囲をモーメントのタイムスタンプに設定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c0599-110">The following code shows how to set the range at **B4** to a moment's timestamp.</span></span>

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

<span data-ttu-id="c0599-111">次のコード サンプルは、セルから日付を取得し、その日付を他の形式に変換する同様の `Moment` 手法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c0599-111">The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.</span></span>

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

<span data-ttu-id="c0599-112">ユーザーが読み取り可能な形式で日付を表示するには、アドインで範囲の書式を設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c0599-112">Your add-in has to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="c0599-113">たとえば `"[$-409]m/d/yy h:mm AM/PM;@"` 、「12/3/18 3:57 PM」と表示されます。</span><span class="sxs-lookup"><span data-stu-id="c0599-113">For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM".</span></span> <span data-ttu-id="c0599-114">日付と時刻の形式の詳細については、「数値書式のカスタマイズに関するガイドライン」の「日付と時刻の書式に関するガイドライン」 [を参照](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) してください。</span><span class="sxs-lookup"><span data-stu-id="c0599-114">For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>


## <a name="see-also"></a><span data-ttu-id="c0599-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="c0599-115">See also</span></span>

- [<span data-ttu-id="c0599-116">Excel JavaScript API を使用してセルを使用する</span><span class="sxs-lookup"><span data-stu-id="c0599-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="c0599-117">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="c0599-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c0599-118">Excel アドインで複数の範囲を同時に操作する</span><span class="sxs-lookup"><span data-stu-id="c0599-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
