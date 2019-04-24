---
title: Excel の範囲にデータの入力規則を追加する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b0b2d886ceb9026ebe41414fed4ef8be1b59cc95
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449212"
---
# <a name="add-data-validation-to-excel-ranges"></a>Excel の範囲にデータの入力規則を追加する

Excel の JavaScript ライブラリには、ブック内の表、列、行、その他の範囲に自動のデータの入力規則をアドインで追加できる API が用意されています。 データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事を参照してください。

- [セルに対するデータの入力規則の適用](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [データの入力規則の詳細](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel でのデータの入力規則の説明と例](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>データの入力規則のプログラムによる制御

`Range.dataValidation` プロパティは [DataValidation](/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これは Excel でデータの入力規則をプログラムにより制御するためのエントリ ポイントになります。 `DataValidation` オブジェクトには 5 つのプロパティがあります。

- `rule` &#8212; 範囲の有効データの構成要素を定義します。 「[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)」を参照してください。
- `errorAlert` &#8212; ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、**情報提供**、**警告**、**停止** などです。 「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。
- `prompt` &#8212; ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。 「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。
- `ignoreBlanks` &#8212; データの入力規則を範囲内の空白セルに適用するかどうかを指定します。 既定値は `true` です。
- `type` &#8212; WholeNumber、Date、TextLength などの入力規則のタイプの読み取り専用 ID です。これは `rule` プロパティを設定すると間接的に設定されます。

> [!NOTE]
> プログラムによって追加されたデータの入力規則は、手動で追加したデータの入力規則と同様に動作します。 具体的に言うと、データの入力規則は、ユーザーがセルに値を直接入力した場合、またはブックの別の場所からセルをコピーして貼り付けたときに、**値**の貼り付けオプションを選択した場合にのみトリガーされます。 ユーザーがセルをコピーしてデータの入力規則のある範囲内に単に貼り付けた場合は、データの入力規則はトリガーされません。

## <a name="creating-validation-rules"></a>入力規則を作成する

範囲にデータの入力規則を追加するには、コードで `Range.dataValidation` にある `DataValidation` オブジェクトの `rule` プロパティを設定する必要があります。 これには 7 つの省略可能なプロパティを持つ [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) オブジェクトが必要です。 *`DataValidationRule` オブジェクトにはこれらのプロパティの 1 つのみを設定できます。* 設定したプロパティにより、入力規則のタイプが決まります。

### <a name="basic-and-datetime-validation-rule-types"></a>Basic および DateTime 入力規則のタイプ

最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則のタイプ) は、その値として [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) オブジェクトをとります。

- `wholeNumber` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と整数が必要です。
- `decimal` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と 10 進数が必要です。
- `textLength` &#8212; `BasicDataValidation`オブジェクトの入力規則の詳細をセルの値の *長さ* に適用します。

次に、入力規則を作成する例を示します。 このコードについては、次の点に注意してください。

- `operator` は二項演算子 "GreaterThan" です。 二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値は左側のオペランドになり、`formula1` で指定された値は右側のオペランドになります。 そのため、この規則では、0 より大きい整数のみが有効になります。 
- `formula1` はハードコーディングされた値です。 コーディングの時点でその正しい値がわからない場合は、その値に Excel の数式 (文字列) を使用することもできます。 たとえば、"=A3" や "=SUM(A4,B5)" を `formula1` の値にすることもできます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

その他の二項演算子のリストについては、「[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)」を参照してください。 

また、2 個の三項演算子、"Between" と "NotBetween" もあります。 これらを使用するには、省略可能な `formula2` プロパティを指定する必要があります。 `formula1` and `formula2`の値はバウンディング オペランドです。 ユーザーがセルに入力しようとする値は、第三の (評価済み) オペランドです。 "Between" 演算子の使用例を次に示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

次の 2 つのルール プロパティは、値として [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをとります。

- `date`
- `time`

`DateTimeDataValidation` オブジェクトは `BasicDataValidation` と同様に構成されています。つまり、プロパティ `formula1`、`formula2`、および `operator` があり、同じ方法で使用されます。 ただし、数式プロパティで数値が使用できない代わりに [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点が異なります。 次に、2018 年 4 月の最初の週の日付として有効な値を定義する例を示します。 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

### <a name="list-validation-rule-type"></a>リスト入力規則のタイプ

有限リストからの値のみを有効な値として指定するには、`DataValidationRule` オブジェクトの `list` プロパティを使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- "Names" という名前のワークシートがあり、"A1:A3" の範囲の値が名前になっていると仮定します。
- `source` プロパティは有効な値のリストを指定します。 文字列引数は名前を含む範囲を参照します。 "Sue, Ricky, Liz" など、カンマで区切られたリストを割り当てることもできます。 
- `inCellDropDown` プロパティは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。 `true` に設定した場合、ドロップダウンには `source` からの値のリストが表示されます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a>カスタムの入力規則のタイプ

カスタムの入力規則式を指定するには、`DataValidationRule` オブジェクトの `custom` プロパティを使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列がある、2 列のテーブルがあると仮定します。
- **Comments** 列の冗長性を軽減するために、アスリート名を含むデータを無効にします。
- `SEARCH(A2,B2)` は A2 内の文字列の開始位置 (B2 内の文字列での) を返します。 A2 が B2 に含まれていない場合は数値を返しません。 `ISNUMBER()` はブール値を返します。 そのため、`formula` プロパティは、**コメント**列の有効なデータが**アスリート名**列内の文字列を含まないデータであることを示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

## <a name="create-validation-error-alerts"></a>入力規則のエラー アラートを作成する

ユーザーがセルに無効なデータを入力しようとした際に表示される、カスタムのエラー アラートを作成できます。 次に簡単な例を示します。 このコードについては、次の点に注意してください。

- `style` プロパティは、ユーザーが情報アラート、警告、または "停止" アラートを取得するかどうかを決定します。 ユーザーによる無効なデータの追加を実際に防止するのは `Stop` のみです。 いずれにせよ、`Warning` と `Information` のポップアップには、ユーザーが無効なデータを入力できるオプションがあります。
- `showAlert` プロパティの既定値は `true` です。 つまり、`showAlert` を `false` に設定するか、カスタムのメッセージ、タイトル、スタイルを設定するカスタム アラートを作成しない限り、Excel ホストは (`Stop` タイプの) 汎用アラートをポップアップ表示します。 このコードでは、カスタムのメッセージとタイトルを設定します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };

    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

詳細については、「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。

## <a name="create-validation-prompts"></a>入力規則のプロンプトを作成する

ユーザーがデータの入力規則が適用されたセルの上でカーソルを動かすか、データの入力規則が適用されたセルを選択した場合に表示される、説明用のダイアログを作成できます。 例を次に示します。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };

    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

詳細については、「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。

## <a name="remove-data-validation-from-a-range"></a>範囲からデータの入力規則を削除する

範囲からデータの入力規則を削除するには、[Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) メソッドを呼び出します。

```js
myrange.dataValidation.clear()
```

消去する範囲とデータの入力規則を追加した範囲とが、まったく同じになる必要はありません。 同じでない場合は、2 つの範囲に重複するセルがあると、それらのセルのみが消去されます。 

> [!NOTE]
> 範囲からデータの入力規則を削除すると、ユーザーが手動で範囲に追加したデータの入力規則も削除されます。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [DataValidation Object (JavaScript API for Excel)](/javascript/api/excel/excel.datavalidation)
- [Range オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.range)
