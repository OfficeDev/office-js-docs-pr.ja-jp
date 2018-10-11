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
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="d4c6d-102">Excel の範囲にデータの入力規則を追加する</span><span class="sxs-lookup"><span data-stu-id="d4c6d-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="d4c6d-p101">Excel の JavaScript ライブラリには、ブック内の表、列、行、その他の範囲に自動のデータの入力規則をアドインで追加できる API が用意されています。データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p101">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="d4c6d-105">セルに対するデータの入力規則の適用</span><span class="sxs-lookup"><span data-stu-id="d4c6d-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="d4c6d-106">データの入力規則の詳細</span><span class="sxs-lookup"><span data-stu-id="d4c6d-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="d4c6d-107">Excel でのデータの入力規則の説明と例</span><span class="sxs-lookup"><span data-stu-id="d4c6d-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="d4c6d-108">データの入力規則のプログラムによる制御</span><span class="sxs-lookup"><span data-stu-id="d4c6d-108">Programmatic control of data validation</span></span>

<span data-ttu-id="d4c6d-p102"> `Range.dataValidation` プロパティは [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これは Excel でデータの入力規則をプログラムにより制御するためのエントリ ポイントとなります。`DataValidation` オブジェクトには 5 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p102">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="d4c6d-p103">`rule` — 範囲の有効データの構成要素を定義します。「[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p103">`rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="d4c6d-p104">`errorAlert` — ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、**[情報提供]**、 **[警告]**、**[停止]** などです。「[DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p104">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="d4c6d-p105">`prompt` — ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p105">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="d4c6d-p106">`ignoreBlanks` — データの入力規則のルールを範囲内の空白セルに適用するかどうかを指定します。既定値は `true`です。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p106">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.</span></span>
- <span data-ttu-id="d4c6d-119">`type` —  WholeNumber、Date、TextLength などの入力規則タイプの読み取り専用  ID です。これは `rule` プロパティを設定すると間接的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="d4c6d-p107">プログラムによって追加されたデータの入力規則は、手動で追加したデータの入力規則と同様に動作します。特に、データの入力規則は、ユーザーがセルに値を直接入力した場合、またはブックの別の場所からセルをコピーして貼り付けたときに 「**値**の 貼り付け」オプションを選んだ場合にのみトリガーされます。ユーザーがセルをコピーしてデータの入力規則のある範囲内に単に貼り付けた場合は、データの入力規則はトリガーされません。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p107">Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="d4c6d-123">入力規則ルールを作成する</span><span class="sxs-lookup"><span data-stu-id="d4c6d-123">Creating validation rules</span></span>

<span data-ttu-id="d4c6d-p108">範囲にデータの入力規則を追加するには、コードで `Range.dataValidation` にある `DataValidation` オブジェクトの `rule`  プロパティを設定する必要があります 。これは、7 つの省略可能なプロパティを持つ [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) オブジェクトが必要です。 *これらのプロパティの  1 つのみが、任意の `DataValidationRule` オブジェクト内に存在することができます* 。含まれているプロパティにより、入力規則の種類が決定されます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p108">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="d4c6d-128">Basic および DateTime 入力規則ルールのタイプ</span><span class="sxs-lookup"><span data-stu-id="d4c6d-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="d4c6d-129">最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則ルール タイプ) は、その値として [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) オブジェクトをとります。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="d4c6d-130">`wholeNumber` ー  `BasicDataValidation` オブジェクトで指定された他の任意の入力規則に加えて、整数が必要です。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="d4c6d-131">`decimal` ー  `BasicDataValidation` オブジェクトで指定された他の任意の入力規則に加えて、10進数が必要です。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="d4c6d-132">`textLength` ー  `BasicDataValidation` オブジェクトの入力規則の詳細を、セルの値の*長さ*に適用します。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="d4c6d-p109">ここに、入力規則を作成する例があります。このコードについては、以下に注意してください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p109">Here is an example of creating a validation rule. Note the following about this code:</span></span>

- <span data-ttu-id="d4c6d-p110"> `operator` は二項演算子 "GreaterThan" です。二二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値は左側のオペランドになり、`formula1` で指定された値は右側のオペランドになります。ですのでこの規則は、0 より大きい整数のみが有ù効であると述べています。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p110">The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="d4c6d-p111"> `formula1` はハードコーディングされた値です。コーディングの時点でその正しい値がわからない場合は、その値に Excel の数式 (文字列) を使用することもできます。たとえば、"=A3" や "=SUM(A4,B5)" を `formula1` の値とすることもできます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p111">The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="d4c6d-141">その他の二項演算子のリストについては、「[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="d4c6d-p112">また、2 個の三項演算子、"Between" と "NotBetween" もあります。これらを使用するには、省略可能な `formula2` プロパティを指定する必要があります。`formula1` と `formula2` の値はバウンディング オペランドです。ユーザーがセルに入力しようとする値は、第三の  (評価済み) オペランドです。次は "Between" 演算子の使用例です。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p112">There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="d4c6d-147">次の 2 つのルール プロパティは、その値として [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをとります。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="d4c6d-p113"> `DateTimeDataValidation` オブジェクトは `BasicDataValidation` と構成が似ています。つまり、プロパティ `formula1\`、 `formula2\`、`operator\`が備わっており、同じ方法で使用されます。違いは、数式プロパティで数値は使えませんが、[ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点です。次に、2018年4 月の最初の週の日付として有効な値を定義する例を示します。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p113">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="d4c6d-151">リスト入力規則ルール タイプ</span><span class="sxs-lookup"><span data-stu-id="d4c6d-151">List validation rule type</span></span>

<span data-ttu-id="d4c6d-p114">有限リストからの値のみが有効な値であるように指定するには、`list` オブジェクトの `DataValidationRule` プロパティを使用します。次に例を示します。このコードについては、以下にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p114">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="d4c6d-155">「Names」という名前のワークシートがあり、"A1:A3" の範囲の値が名前であることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="d4c6d-p115"> `source` プロパティは有効な値のリストを指定します。名前の範囲がそれに割り当てられています。コンマで区切られたリスト (たとえば "Sue, Ricky, Liz") を割り当てることもできます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p115">The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="d4c6d-p116"> `inCellDropDown` プロパティでは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。`true` に設定した場合 、ドロップダウンには `source` からの値のリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p116">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="d4c6d-161">カスタムの入力規則ルール タイプ</span><span class="sxs-lookup"><span data-stu-id="d4c6d-161">Custom validation rule type</span></span>

<span data-ttu-id="d4c6d-p117">カスタムの入力規則式を指定するには、`custom` オブジェクトの `DataValidationRule` プロパティ を使用します。次に例を示します。このコードについては、以下にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p117">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="d4c6d-165">ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列をもつ 2 列のテーブルがあると仮定しています。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="d4c6d-166">**Comments** 列の冗長性を軽減するために、アスリート名を含むデータを無効にします。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="d4c6d-p118">`SEARCH(A2,B2)` A2 内の文字列の、B2 内の文字列の中での開始位置を返します。A2 が B2 に含まれていない場合は数値を返しません。`ISNUMBER()` はブール値を返します。つまり、`formula` プロパティは、[**コメント**] 列の有効なデータは [**アスリート名**] 列内の文字列を含まないデータであると言っています。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p118">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="d4c6d-171">入力規則エラー アラートを作成する</span><span class="sxs-lookup"><span data-stu-id="d4c6d-171">Create validation error alerts</span></span>

<span data-ttu-id="d4c6d-p119">ユーザーがセルに無効なデータを入力しようとすると表示されるカスタムのエラー アラートを作成できます。次に簡単な例を示します。このコードについては、以下にご注意ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p119">You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:</span></span>

- <span data-ttu-id="d4c6d-p120"> `style` プロパティは、ユーザーが情報アラート、警告、または「停止」アラートを取得するかどうかを決定します。ユーザーが無効なデータを追加することを実際に防止するのは、`Stop` のみです。`Warning` と `Information` のポップアップには、設定にかかわらずユーザーが無効なデータを入力できるオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p120">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` actually prevents the user from adding invalid data. The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="d4c6d-p121"> `showAlert` プロパティの既定値は `true` です。つまり、`Stop` を `showAlert\`に設定するか、カスタムのメッセージ、タイトル、スタイルを設定するカスタム アラートを作成しない限り、Excel ホストは (`false` 型の) 汎用アラートをポップアップ表示します。このコードでは、カスタムのメッセージとタイトルを設定します。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p121">The `showAlert` property defaults to `true`. This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.</span></span>


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

<span data-ttu-id="d4c6d-181">詳細情報については、「 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert) 」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="d4c6d-182">入力規則プロンプトを作成する</span><span class="sxs-lookup"><span data-stu-id="d4c6d-182">Create validation prompts</span></span>

<span data-ttu-id="d4c6d-p122">ユーザーがデータの入力規則が適用されたセルの上でカーソルを動かすか、またはそのようなセルを選択した場合に表示される説明用ダイアログを作成できます。以下はその例です。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p122">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example:</span></span>

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

<span data-ttu-id="d4c6d-185">詳細情報については、「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="d4c6d-186">範囲からデータ入力規則を削除する</span><span class="sxs-lookup"><span data-stu-id="d4c6d-186">Remove data validation from a range</span></span>

<span data-ttu-id="d4c6d-187">範囲からデータの入力規則を削除するには、[Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="d4c6d-p123">消去する範囲は、データの入力規則を追加した範囲と正確に同じである必要はありません。同じでない場合は、2 つの範囲に重複するセルがある場合に、それらのセルが消去されます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-p123">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="d4c6d-190">範囲からデータの入力規則を削除すると、ユーザーが手動で範囲に追加したデータの入力規則も削除されます。</span><span class="sxs-lookup"><span data-stu-id="d4c6d-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="d4c6d-191">関連項目</span><span class="sxs-lookup"><span data-stu-id="d4c6d-191">See also</span></span>

- [<span data-ttu-id="d4c6d-192">Excel の JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="d4c6d-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d4c6d-193">DataValidation オブジェクト (Excel の JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="d4c6d-193">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="d4c6d-194">Range オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="d4c6d-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
