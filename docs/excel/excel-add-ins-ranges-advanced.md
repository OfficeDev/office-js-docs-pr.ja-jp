---
title: Excel JavaScript API を使用して範囲を操作する (高度)
description: ''
ms.date: 12/18/2018
ms.openlocfilehash: 6d3da1e7eff4e61ae1b88213d0b432581d8f6a8a
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383240"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Excel JavaScript API を使用して範囲を操作する (高度)

この記事は、「[Excel JavaScript API を使用して範囲を操作する (基本)](excel-add-ins-ranges.md)」の情報に基づいており、コード サンプルでは Excel JavaScript API を使って範囲のより高度なタスクを実行する方法を示します。 **Range** オブジェクトがサポートするプロパティとメソッドの完全な一覧については、「[Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)」を参照してください。

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Moment-MSDate プラグインを使用した日付の操作

[Moment JavaScript ライブラリ](https://momentjs.com/)により、日付とタイムスタンプが便利に使用できるようになります。 [Moment-MSDate プラグイン](https://www.npmjs.com/package/moment-msdate)は、日付と時刻の形式を Excel に適したものに変換します。 これは、[NOW 関数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)から返される形式と同じです。

次のコードは、範囲 **B4** に時刻のタイムスタンプを設定する方法を示しています。

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

これは、次の例に示すように、セルから日付を取得して、その日付を時刻などの形式に変換するのと同様の手法です。

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

アドインでは、わかりやすい形式で日付が表示されるように、範囲の書式を設定する必要があります。 たとえば、`"[$-409]m/d/yy h:mm AM/PM;@"` では時刻が "12/3/18 3:57 PM" のように表示されます。 日付と時刻の数値書式の詳細については、「[表示形式のカスタマイズに関するガイドラインを確認する](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)」の記事で「日付と時刻の表示に関するガイドライン」を参照してください。

## <a name="work-with-multiple-ranges-simultaneously-preview"></a>複数の範囲を同時に操作する (プレビュー)

> [!NOTE]
> 現在、`RangeAreas` オブジェクトは、パブリック プレビュー (ベータ版) でのみ利用できます。 この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。

`RangeAreas` オブジェクトを使用すると、アドインの操作を一度に複数の範囲で実行できます。 これらの範囲は、連続していても連続していなくても構いません。 `RangeAreas` については、「[Excel アドインで複数の範囲を同時に操作する](excel-add-ins-multiple-ranges.md)」にさらに詳しい説明があります。

## <a name="copy-and-paste-preview"></a>コピーと貼り付け (プレビュー)

> [!NOTE]
> 現在、`Range.copyFrom` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。 この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。

範囲の `copyFrom` 関数では、Excel UI のコピーと貼り付けの動作をレプリケートします。 `copyFrom` が呼び出される範囲オブジェクトがコピー先になります。
コピーされるソースは、範囲または範囲を表す文字列のアドレスとして渡されます。 次のコード サンプルでは、**A1:E1** のデータを **G1** で始まる範囲にコピーします (この貼り付けは **G1:K1** で終わります)。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` には、省略可能なパラメーターが 3 つあります。

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` では、ソースからコピー先にコピーされるデータを指定します。

- `"Formulas"` では、ソースのセルの数式が転送され、それらの数式の範囲の相対配置は保持されます。 任意の数式以外のエントリはそのままコピーされます。
- `"Values"` では、データ値と、数式の場合は数式の結果をコピーします。
- `"Formats"` では、フォント、色、およびその他の書式設定を含む、範囲の書式設定をコピーしますが、値はコピーしません。
- `"All"` (既定のオプション) では、データと書式設定の両方がコピーされます。見つかった場合、セルの数式は保持されます。

`skipBlanks` では、空白セルをコピー先にコピーするかどうかを設定します。 true の場合、`copyFrom` ではソースの範囲にある空白セルはスキップされます。
スキップされたセルでは、コピー先の範囲内の対応するセルにある既存のデータを上書きすることはありません。 既定値は false です。

`transpose` では、ソースの場所へのデータの行と列の入れ替えを行うかどうかを決定します。
行と列を入れ替える範囲は対角線で反転されるため、行 **1**、**2**、**3** が列 **A**、**B**、**C** になります。

次のコード サンプルと画像は、この動作をシンプルなシナリオで示しています。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*前の関数が実行される前。*

![範囲のコピー メソッドが実行される前の Excel のデータ](../images/excel-range-copyfrom-skipblanks-before.png)

*前の関数が実行された後。*

![範囲のコピー メソッドが実行された後の Excel のデータ](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a>重複を削除 (プレビュー)

> [!NOTE]
> 現在、Range オブジェクトの `removeDuplicates` 関数は、パブリック プレビュー (ベータ版) でのみ利用できます。 この機能を使用するには、Office.js CDN のベータ版のライブラリを使用する必要があります: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。
> TypeScript を使用している場合、または IntelliSense に TypeScript 型定義ファイルを使用するコード エディターを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。

Range オブジェクトの `removeDuplicates` 関数は、指定された列で重複するエントリを持つ行を削除します。 関数は、範囲の一番小さい値のインデックスから一番大きい値のインデックスへ向かって各行を移動します (上から下へ)。 任意の行で、指定された 1 つまたは複数の列が範囲より前に表示されている場合、その行は削除されます。 範囲にある削除された行の下の行が上に移動します。 `removeDuplicates` は、範囲外にあるセルの位置には影響しません。

`removeDuplicates` は、どの重複をチェックするかを示す列インデックスを表す `number[]` を受け取ります。 この配列は、0 から始まり、ワークシートではなく範囲を基準にしています。 この関数は、最初の行がヘッダーかどうかを指定するブール値のパラメーターも受け取ります。 **true** の場合、重複について考慮するとき最初の行は無視されます。 `removeDuplicates` 関数は、削除する行の数と、残りの一意の行の数を指定する `RemoveDuplicatesResult` オブジェクトを返します。

範囲の `removeDuplicates` 関数を使う場合、次の点に注意してください。

- `removeDuplicates` は、関数の結果ではなくセルの値を考慮します。 2 つの異なる関数が同じ結果として評価される場合、セルの値は重複と見なしません。
- 空のセルは、`removeDuplicates` に無視されることはありません。 空のセルの値は、その他の値と同様に扱われます。 つまり、範囲に含まれる空の行は `RemoveDuplicatesResult` に含まれることになります。

次の例では、最初の列に重複する値があるエントリを削除する方法を示します。

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*前の関数が実行される前。*

![範囲の重複を削除するメソッドが実行される前の Excel のデータ](../images/excel-ranges-remove-duplicates-before.png)

*前の関数が実行された後。*

![範囲の重複を削除するメソッドが実行された後の Excel のデータ](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用して範囲を操作する](excel-add-ins-ranges.md)
- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)