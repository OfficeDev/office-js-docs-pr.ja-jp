---
title: Excel の範囲にデータの入力規則を追加する
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: af965df4a1aece5b7f8d5ea89664519b576a4850
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925312"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>Excel 範囲にデータの入力規則を追加する (プレビュー)

> [!NOTE]
> データの入力規則 API はプレビューとして提供されていますが、使用するには Office JavaScript ライブラリのベータ版を読み込む必要があります。 URL は、 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js です。 TypeScript を使用している場合、またはコードエディタで Intellisense 用の TypeScript 型定義ファイルを使用している場合は、https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts を使用してください。

> [!NOTE]
> データの入力規則 API はプレビュー中ですが、この記事の API リファレンスへのリンクは機能しません。 その間、 [ドラフト Excel API リファレンス](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel)を使用することができます。

Excel JavaScript ライブラリには、ワークブックに、表、列、行、その他の範囲に自動データ入力規則をアドインで追加できる API が用意されています。 データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事をご覧ください。

- [セルにデータの入力規則を適用する](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [データの入力規則の詳細](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel でのデータの入力規則の説明と例](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>データの入力規則のプログラムによる制御

プロパティは、[データの入力規則](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これが Excel でデータの入力規則をプログラムにより制御するためのエントリポイントとなります。`Range.dataValidation` オブジェクトには、次のような 5つのプロパティがあります。`DataValidation`

- `rule` — 範囲の有効データの構成要素を定義します。 「[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)」を参照してください。
- `errorAlert` — ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、 **[情報提供]**、 **[警告]**、**[停止]** などです。 「[DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。
- `prompt` — ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。 「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」を参照してください。
- `ignoreBlanks` — データの入力規則のルールを範囲内の空白セルに適用するかどうかを指定します。 既定値は `true` です。
- `type` —  WholeNumber、Date、TextLength などの入力規則タイプの読み取り専用  ID です。これは `rule` プロパティが設定されると間接的に設定されます。

> [!NOTE]
> プログラムにより追加されたデータの入力規則は、手動で追加されたデータの入力規則と同様に動作します。 特に、データの入力規則は、ユーザーがセルに値を直接入力したり、ワークブックの別の場所からセルをコピーして貼り付けるときに 「**値**の 貼り付け」オプションを選んだりした場合にのみトリガーされます。 ユーザーがセルをコピーし、データの入力規則が設定された範囲に単にペーストする場合、入力規則はトリガーされません。

### <a name="creating-validation-rules"></a>入力規則ルールを作成する

範囲にデータの入力規則を追加するには、コードで `rule` にある `DataValidation` オブジェクトの `Range.dataValidation` プロパティを設定する必要があります 。 これは、7 つのオプション プロパティのある [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) オブジェクトを取得します。 *どの `DataValidationRule` オブジェクトでも、これらの特性が 1 つ以上表示されることはありません。* 含めるプロパティによって、入力規則のタイプが決まります。

#### <a name="basic-and-datetime-validation-rule-types"></a>Basic および DateTime 入力規則ルールのタイプ

最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則ルール タイプ) は、[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) オブジェクトをその値として取得します。

- `wholeNumber` ー  `BasicDataValidation` オブジェクトで指定された他の妥当性確認に加えて整数を必要とします。
- `decimal` ー  `BasicDataValidation` オブジェクトで指定された他の妥当性確認に加えて、10進数が必要です。
- `textLength` ー  `BasicDataValidation` オブジェクトの妥当性確認の詳細をセルの値の *長さ* に適用します。

次に、入力規則のルールを作成する例を示します。 このコードについては、次の点に注意してください。

- は二項演算子「GreaterThan」です。`operator` 二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値が左側のオペランドになり、 `formula1` で指定された値が右側のオペランドになります。 したがって、このルールでは、0 より大きな整数だけが有効です。 
- は、ハードコーディングされた数字です。`formula1` コード時にどの値にすべきかわからない場合は、値の Excel 式を (文字列として) 使用することもできます。 たとえば、「= A3」および「= SUM（A4、B5）」も、 `formula1`の値にできます。

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

他の二項演算子のリストについては、「[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 」を参照してください。 

また、「Between」と「NotBetween」の2つの三項演算子もあります。 これらを使用するには、オプションの `formula2` プロパティを指定する必要があります。 と `formula2` 値はバウンディング オペランドです。`formula1` ユーザーがセルに入力しようとする値は、3 番目の (評価された) オペランドです。 以下は「Between」演算子の使用例です。

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

次の 2 つのルール プロパティは、 [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをその値として取得します。

- `date`
- `time`

オブジェクトは `BasicDataValidation` と構成が似ています。つまり、プロパティ `formula1`、 `formula2`、`operator`が備わっており、同じ方法で使われます。`DateTimeDataValidation` 違うのは、数式プロパティで数値を使えませんが、 [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点です。 以下は、2018 年 4 月の第 1 週の日付として有効な値を定義する例です。 

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

#### <a name="list-validation-rule-type"></a>リスト入力規則ルール タイプ

有限リストからの値のみが有効な値であるように指定するには、`list` オブジェクトの `DataValidationRule` プロパティを使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- 「Names」という名前のワークシートがあり、「A1：A3」の範囲の値が名前であることが前提です。
- プロパティは、有効な値のリストを指定します。`source` 名前を含んだ範囲が割り当てられています。 コンマ区切りのリストを割り当てることもできます。たとえば、「Sue、Ricky、Liz」です。 
- プロパティでは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。`inCellDropDown` に設定されている場合 、ドロップダウンには `source` からの値のリストが表示されます。`true`

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

#### <a name="custom-validation-rule-type"></a>カスタムの入力規則ルール タイプ

カスタムの入力規則式を指定するには、`custom` オブジェクトの `DataValidationRule` プロパティ を使用します。 次に例を示します。 このコードについては、次の点に注意してください。

- ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列をもつ 2 列のテーブルがあると仮定します。
- **Comments** 列の冗長性を軽減するには、アスリート名を含むデータを無効にします。
- `SEARCH(A2,B2)` B2 の文字列での A2 の文字列の開始位置を返します。 B2 に A2 が含まれていない場合は、数値は返されません。 `ISNUMBER()` ブール値を返します。 したがって `formula` プロパティは、 **Comment** 列の有効なデータは、 **Athlete Name** 列の文字列を含まないデータであることを示します。

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

### <a name="create-validation-error-alerts"></a>入力規則エラー アラートを作成する

ユーザーがセルに無効なデータを入力しようとすると表示されるカスタムのエラー アラートを作成できます。 次に簡単な例を示します。 このコードについては、次の点に注意してください。

- プロパティは、ユーザーが情報アラート、警告、または「停止」アラートを取得するかどうかを決定します。`style` 実際のところ、ユーザーが無効なデータを追加できないようにするのは、`Stop` のみです。 と `Information` のポップアップには、設定にかかわらずユーザーが無効なデータを入力できるオプションがあります。`Warning`
- プロパティの既定値は `true` です。`showAlert` つまり、`showAlert` を `false` に設定するか、カスタムのメッセージ、タイトル、スタイルを設定するカスタム アラートを作成するかしない限り、Excel ホストは (`Stop` タイプの) 汎用アラートをポップアップ表示します。 このコードはカスタム メッセージとタイトルを設定します。


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

詳細情報については、「 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert) 」を参照してください。

### <a name="create-validation-prompts"></a>入力規則プロンプトを作成する

ユーザーがデータ入力規則が適用されたセルの上でカーソルを動かすか、またはこのようなセルを選択するかしたときに表示される説明用ダイアログを作成できます。 例を次に示します。

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

詳細情報については、「 [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt) 」を参照してください。

### <a name="remove-data-validation-from-a-range"></a>範囲からデータ入力規則を削除する

範囲からデータ入力規則を削除するには、[Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) メソッドを呼び出します。

```js
myrange.dataValidation.clear()
```

削除する範囲は、データ入力規則を追加した範囲とまったく同じ範囲でなくてもかまいません。 範囲が同じでない場合は、2 つの範囲でオーバーラップしているセルがあれば、そのようなセルのみが削除されます。 

> [!NOTE]
> 範囲からデータ入力規則を削除すると、ユーザーが手動で範囲に追加したデータ入力規則も削除されます。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API の中心概念](excel-add-ins-core-concepts.md)
- [DataValidation オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Range オブジェクト (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
