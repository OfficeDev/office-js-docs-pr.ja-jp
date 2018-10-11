---
title: Excel の範囲にデータの入力規則を追加する
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 9e3aba8d87e84405bb3e1ae35a8d35d60ce8e2b6
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459155"
---
# <a name="add-data-validation-to-excel-ranges"></a>Excel の範囲にデータの入力規則を追加する

Excel の JavaScript ライブラリには、ブック内の表、列、行、その他の範囲に自動のデータの入力規則をアドインで追加できる API が用意されています。データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事をご覧ください。

- [セルに対するデータの入力規則の適用](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [データの入力規則の詳細](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel でのデータの入力規則の説明と例](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>データの入力規則のプログラムによる制御

 `Range.dataValidation` プロパティは [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これは Excel でデータの入力規則をプログラムにより制御するためのエントリ ポイントとなります。`DataValidation` オブジェクトには 5 つのプロパティがあります。

- `rule` — 範囲の有効データの構成要素を定義します。「[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)」をご覧ください。
- `errorAlert` — ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、**[情報提供]**、 **[警告]**、**[停止]** などです。「[DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)」をご覧ください。
- `prompt` — ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」をご覧ください。
- `ignoreBlanks` — データの入力規則のルールを範囲内の空白セルに適用するかどうかを指定します。既定値は `true`です。
- `type` —  WholeNumber、Date、TextLength などの入力規則タイプの読み取り専用  ID です。これは `rule` プロパティを設定すると間接的に設定されます。

> [!NOTE]
> プログラムによって追加されたデータの入力規則は、手動で追加したデータの入力規則と同様に動作します。特に、データの入力規則は、ユーザーがセルに値を直接入力した場合、またはブックの別の場所からセルをコピーして貼り付けたときに 「**値**の 貼り付け」オプションを選んだ場合にのみトリガーされます。ユーザーがセルをコピーしてデータの入力規則のある範囲内に単に貼り付けた場合は、データの入力規則はトリガーされません。

## <a name="creating-validation-rules"></a>入力規則ルールを作成する

範囲にデータの入力規則を追加するには、コードで `Range.dataValidation` にある `DataValidation` オブジェクトの `rule`  プロパティを設定する必要があります 。これは、7 つの省略可能なプロパティを持つ [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) オブジェクトが必要です。 *これらのプロパティの  1 つのみが、任意の `DataValidationRule` オブジェクト内に存在することができます* 。含まれているプロパティにより、入力規則の種類が決定されます。

### <a name="basic-and-datetime-validation-rule-types"></a>Basic および DateTime 入力規則ルールのタイプ

最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則ルール タイプ) は、その値として [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) オブジェクトをとります。

- `wholeNumber` ー  `BasicDataValidation` オブジェクトで指定された他の任意の入力規則に加えて、整数が必要です。
- `decimal` ー  `BasicDataValidation` オブジェクトで指定された他の任意の入力規則に加えて、10進数が必要です。
- `textLength` ー  `BasicDataValidation` オブジェクトの入力規則の詳細を、セルの値の*長さ*に適用します。

ここに、入力規則を作成する例があります。このコードについては、以下に注意してください。

-  `operator` は二項演算子 "GreaterThan" です。二二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値は左側のオペランドになり、`formula1` で指定された値は右側のオペランドになります。ですのでこの規則は、0 より大きい整数のみが有ù効であると述べています。 
-  `formula1` はハードコーディングされた値です。コーディングの時点でその正しい値がわからない場合は、その値に Excel の数式 (文字列) を使用することもできます。たとえば、"=A3" や "=SUM(A4,B5)" を `formula1` の値とすることもできます。

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

その他の二項演算子のリストについては、「[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 」をご覧ください。 

また、2 個の三項演算子、"Between" と "NotBetween" もあります。これらを使用するには、省略可能な `formula2` プロパティを指定する必要があります。`formula1` と `formula2` の値はバウンディング オペランドです。ユーザーがセルに入力しようとする値は、第三の  (評価済み) オペランドです。次は "Between" 演算子の使用例です。

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

次の 2 つのルール プロパティは、その値として [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをとります。

- `date`
- `time`

 `DateTimeDataValidation` オブジェクトは `BasicDataValidation` と構成が似ています。つまり、プロパティ `formula1`、 `formula2`、`operator`が備わっており、同じ方法で使用されます。違いは、数式プロパティで数値は使えませんが、[ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点です。次に、2018年4 月の最初の週の日付として有効な値を定義する例を示します。 

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

### <a name="list-validation-rule-type"></a>リスト入力規則ルール タイプ

有限リストからの値のみが有効な値であるように指定するには、`list` オブジェクトの `DataValidationRule` プロパティを使用します。次に例を示します。このコードについては、以下にご注意ください。

- 「Names」という名前のワークシートがあり、"A1:A3" の範囲の値が名前であることを前提としています。
-  `source` プロパティは有効な値のリストを指定します。名前の範囲がそれに割り当てられています。コンマで区切られたリスト (たとえば "Sue, Ricky, Liz") を割り当てることもできます。 
-  `inCellDropDown` プロパティでは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。`true` に設定した場合 、ドロップダウンには `source` からの値のリストが表示されます。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a>カスタムの入力規則ルール タイプ

カスタムの入力規則式を指定するには、`custom` オブジェクトの `DataValidationRule` プロパティ を使用します。次に例を示します。このコードについては、以下にご注意ください。

- ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列をもつ 2 列のテーブルがあると仮定しています。
- **Comments** 列の冗長性を軽減するために、アスリート名を含むデータを無効にします。
- `SEARCH(A2,B2)` A2 内の文字列の、B2 内の文字列の中での開始位置を返します。A2 が B2 に含まれていない場合は数値を返しません。`ISNUMBER()` はブール値を返します。つまり、`formula` プロパティは、[**コメント**] 列の有効なデータは [**アスリート名**] 列内の文字列を含まないデータであると言っています。

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

## <a name="create-validation-error-alerts"></a>入力規則エラー アラートを作成する

ユーザーがセルに無効なデータを入力しようとすると表示されるカスタムのエラー アラートを作成できます。次に簡単な例を示します。このコードについては、以下にご注意ください。

-  `style` プロパティは、ユーザーが情報アラート、警告、または「停止」アラートを取得するかどうかを決定します。ユーザーが無効なデータを追加することを実際に防止するのは、`Stop` のみです。`Warning` と `Information` のポップアップには、設定にかかわらずユーザーが無効なデータを入力できるオプションがあります。
-  `showAlert` プロパティの既定値は `true` です。つまり、`Stop` を `showAlert`に設定するか、カスタムのメッセージ、タイトル、スタイルを設定するカスタム アラートを作成しない限り、Excel ホストは (`false` 型の) 汎用アラートをポップアップ表示します。このコードでは、カスタムのメッセージとタイトルを設定します。


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

詳細情報については、「 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert) 」をご覧ください。

## <a name="create-validation-prompts"></a>入力規則プロンプトを作成する

ユーザーがデータの入力規則が適用されたセルの上でカーソルを動かすか、またはそのようなセルを選択した場合に表示される説明用ダイアログを作成できます。以下はその例です。

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

詳細情報については、「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」をご覧ください。

## <a name="remove-data-validation-from-a-range"></a>範囲からデータ入力規則を削除する

範囲からデータの入力規則を削除するには、[Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) メソッドを呼び出します。

```js
myrange.dataValidation.clear()
```

消去する範囲は、データの入力規則を追加した範囲と正確に同じである必要はありません。同じでない場合は、2 つの範囲に重複するセルがある場合に、それらのセルが消去されます。 

> [!NOTE]
> 範囲からデータの入力規則を削除すると、ユーザーが手動で範囲に追加したデータの入力規則も削除されます。

## <a name="see-also"></a>関連項目

- [Excel の JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)
- [DataValidation オブジェクト (Excel の JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Range オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
