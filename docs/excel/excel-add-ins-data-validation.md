---
title: Excel の範囲にデータの入力規則を追加する
description: JavaScript API Excelにより、ブック内のテーブル、列、行、その他の範囲に自動データ検証をアドインで追加する方法について説明します。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: f13448d7739a5bc437e674341753ddf672137ca2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744787"
---
# <a name="add-data-validation-to-excel-ranges"></a>Excel の範囲にデータの入力規則を追加する

Excel の JavaScript ライブラリには、ブック内の表、列、行、その他の範囲に自動のデータの入力規則をアドインで追加できる API が用意されています。 データ検証の概念と用語を理解するには、ユーザーが UI を使用してデータ検証を追加する方法に関する以下のExcelしてください。

- [セルに対するデータの入力規則の適用](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249)
- [データの入力規則の詳細](https://support.microsoft.com/office/f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel でのデータの入力規則の説明と例](https://support.microsoft.com/help/211485)

## <a name="programmatic-control-of-data-validation"></a>データの入力規則のプログラムによる制御

`Range.dataValidation` プロパティは [DataValidation](/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これは Excel でデータの入力規則をプログラムにより制御するためのエントリ ポイントになります。 `DataValidation` オブジェクトには 5 つのプロパティがあります。

- `rule` &#8212; 範囲の有効データの構成要素を定義します。 「[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)」を参照してください。
- `errorAlert` &#8212; ユーザーが無効なデータを入力した場合にエラーがポップアップするかどうかを指定し、警告テキスト、タイトル、およびスタイルを定義します。たとえば、 `information`、 、 `warning`と です `stop`。 「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。
- `prompt` &#8212; ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。 「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。
- `ignoreBlanks` &#8212; データの入力規則を範囲内の空白セルに適用するかどうかを指定します。 既定値は `true` です。
- `type` &#8212; WholeNumber、Date、TextLength などの入力規則のタイプの読み取り専用 ID です。これは `rule` プロパティを設定すると間接的に設定されます。

> [!NOTE]
> プログラムによって追加されたデータの入力規則は、手動で追加したデータの入力規則と同様に動作します。 具体的に言うと、データの入力規則は、ユーザーがセルに値を直接入力した場合、またはブックの別の場所からセルをコピーして貼り付けたときに、**値** の貼り付けオプションを選択した場合にのみトリガーされます。 ユーザーがセルをコピーしてデータの入力規則のある範囲内に単に貼り付けた場合は、データの入力規則はトリガーされません。

## <a name="creating-validation-rules"></a>入力規則を作成する

範囲にデータの入力規則を追加するには、コードで `Range.dataValidation` にある `DataValidation` オブジェクトの `rule` プロパティを設定する必要があります。 これには 7 つの省略可能なプロパティを持つ [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) オブジェクトが必要です。 *`DataValidationRule` オブジェクトにはこれらのプロパティの 1 つのみを設定できます。* 設定したプロパティにより、入力規則のタイプが決まります。

### <a name="basic-and-datetime-validation-rule-types"></a>Basic および DateTime 入力規則のタイプ

最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則のタイプ) は、その値として [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) オブジェクトをとります。

- `wholeNumber` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と整数が必要です。
- `decimal` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と 10 進数が必要です。
- `textLength` &#8212; `BasicDataValidation`オブジェクトの入力規則の詳細をセルの値の *長さ* に適用します。

次に、入力規則を作成する例を示します。 このコードについては、次の点に注意してください。

- は `operator` バイナリ演算子です `greaterThan`。 二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値は左側のオペランドになり、`formula1` で指定された値は右側のオペランドになります。 そのため、この規則では、0 より大きい整数のみが有効になります。
- `formula1` はハードコーディングされた値です。 コーディングの時点でその正しい値がわからない場合は、その値に Excel の数式 (文字列) を使用することもできます。 たとえば、"=A3" や "=SUM(A4,B5)" を `formula1` の値にすることもできます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: Excel.DataValidationOperator.greaterThan
            }
        };

    await context.sync();
});
```

その他の二項演算子のリストについては、「[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)」を参照してください。

また、2 つの 3 項演算子があります。 `between` `notBetween` これらを使用するには、省略可能な `formula2` プロパティを指定する必要があります。 `formula1` and `formula2`の値はバウンディング オペランドです。 ユーザーがセルに入力しようとする値は、第三の (評価済み) オペランドです。 "Between" 演算子の使用例を次に示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
              operator: Excel.DataValidationOperator.between
            }
        };

    await context.sync();
});
```

次の 2 つのルール プロパティは、値として [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをとります。

- `date`
- `time`

`DateTimeDataValidation` オブジェクトは `BasicDataValidation` と同様に構成されています。つまり、プロパティ `formula1`、`formula2`、および `operator` があり、同じ方法で使用されます。 ただし、数式プロパティで数値が使用できない代わりに [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点が異なります。 次に、有効な値を 2022 年 4 月の最初の週の日付として定義する例を示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            date: {
                formula1: "2022-04-01",
                formula2: "2022-04-08",
                operator: Excel.DataValidationOperator.between
            }
        };

    await context.sync();
});
```

### <a name="list-validation-rule-type"></a>リスト入力規則のタイプ

有限リストからの値のみを有効な値として指定するには、`DataValidationRule` オブジェクトの `list` プロパティを使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- "Names" という名前のワークシートがあり、"A1:A3" の範囲の値が名前になっていると仮定します。
- `source` プロパティは有効な値のリストを指定します。 文字列引数は名前を含む範囲を参照します。 "Sue, Ricky, Liz" など、カンマで区切られたリストを割り当てることもできます。
- `inCellDropDown` プロパティは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。 `true` に設定した場合、ドロップダウンには `source` からの値のリストが表示されます。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");   
    let nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    await context.sync();
})
```

### <a name="custom-validation-rule-type"></a>カスタムの入力規則のタイプ

カスタムの入力規則式を指定するには、`DataValidationRule` オブジェクトの `custom` プロパティを使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列がある、2 列のテーブルがあると仮定します。
- **Comments** 列の冗長性を軽減するために、アスリート名を含むデータを無効にします。
- `SEARCH(A2,B2)` は A2 内の文字列の開始位置 (B2 内の文字列での) を返します。 A2 が B2 に含まれていない場合は数値を返しません。 `ISNUMBER()` はブール値を返します。 そのため、`formula` プロパティは、**コメント** 列の有効なデータが **アスリート名** 列内の文字列を含まないデータであることを示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    await context.sync();
});
```

## <a name="create-validation-error-alerts"></a>入力規則のエラー アラートを作成する

ユーザーがセルに無効なデータを入力しようとした際に表示される、カスタムのエラー アラートを作成できます。 次に簡単な例を示します。 このコードについては、次の点に注意してください。

- `style` プロパティは、ユーザーが情報アラート、警告、または "停止" アラートを取得するかどうかを決定します。 ユーザーによる無効なデータの追加を実際に防止するのは `stop` のみです。 ポップアップには、ユーザーが`warning``information`無効なデータを入力できるオプションがあります。
- `showAlert` プロパティの既定値は `true` です。 つまり、Excelメッセージ、タイトル、およびスタイルに設定または設定するカスタム アラートを作成しない限り、一般的なアラート (`stop`種類) `false` `showAlert` がポップアップします。 このコードでは、カスタムのメッセージとタイトルを設定します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // The default is 'true'.
              style: Excel.DataValidationAlertStyle.stop,
            title: "Negative or Decimal Number Entered"
        };

    // Set range.dataValidation.rule and optionally .prompt here.

    await context.sync();
});
```

詳細については、「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。

## <a name="create-validation-prompts"></a>入力規則のプロンプトを作成する

ユーザーがデータの入力規則が適用されたセルの上でカーソルを動かすか、データの入力規則が適用されたセルを選択した場合に表示される、説明用のダイアログを作成できます。 次に例を示します。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");

    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // The default is 'false'.
            title: "Positive Whole Numbers Only."
        };

    // Set range.dataValidation.rule and optionally .errorAlert here.

    await context.sync();
});
```

詳細については、「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。

## <a name="remove-data-validation-from-a-range"></a>範囲からデータの入力規則を削除する

範囲からデータの入力規則を削除するには、[Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1)) メソッドを呼び出します。

```js
myrange.dataValidation.clear()
```

消去する範囲とデータの入力規則を追加した範囲とが、まったく同じになる必要はありません。 同じでない場合は、2 つの範囲に重複するセルがあると、それらのセルのみが消去されます。 

> [!NOTE]
> 範囲からデータの入力規則を削除すると、ユーザーが手動で範囲に追加したデータの入力規則も削除されます。

## <a name="see-also"></a>関連項目

- [Office アドインの Excel JavaScript オブジェクト モデル](excel-add-ins-core-concepts.md)
- [DataValidation Object (JavaScript API for Excel)](/javascript/api/excel/excel.datavalidation)
- [Range オブジェクト (JavaScript API for Excel)](/javascript/api/excel/excel.range)
