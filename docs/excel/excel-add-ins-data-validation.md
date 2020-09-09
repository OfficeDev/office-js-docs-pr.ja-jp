---
title: Excel の範囲にデータの入力規則を追加する
description: Excel JavaScript Api を使用して、ブック内のテーブル、列、行、およびその他の範囲に自動的なデータの入力規則を追加する方法について説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ce792e36f9ad24eb4b26e2034c59063d65940be4
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408552"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="2ee89-103">Excel の範囲にデータの入力規則を追加する</span><span class="sxs-lookup"><span data-stu-id="2ee89-103">Add data validation to Excel ranges</span></span>

<span data-ttu-id="2ee89-104">Excel の JavaScript ライブラリには、ブック内の表、列、行、その他の範囲に自動のデータの入力規則をアドインで追加できる API が用意されています。</span><span class="sxs-lookup"><span data-stu-id="2ee89-104">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="2ee89-105">データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-105">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="2ee89-106">セルに対するデータの入力規則の適用</span><span class="sxs-lookup"><span data-stu-id="2ee89-106">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="2ee89-107">データの入力規則の詳細</span><span class="sxs-lookup"><span data-stu-id="2ee89-107">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="2ee89-108">Excel でのデータの入力規則の説明と例</span><span class="sxs-lookup"><span data-stu-id="2ee89-108">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="2ee89-109">データの入力規則のプログラムによる制御</span><span class="sxs-lookup"><span data-stu-id="2ee89-109">Programmatic control of data validation</span></span>

<span data-ttu-id="2ee89-110">`Range.dataValidation` プロパティは [DataValidation](/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これは Excel でデータの入力規則をプログラムにより制御するためのエントリ ポイントになります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-110">The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="2ee89-111">`DataValidation` オブジェクトには 5 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-111">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="2ee89-112">`rule` &#8212; 範囲の有効データの構成要素を定義します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-112">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="2ee89-113">「[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-113">See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="2ee89-114">`errorAlert` &#8212; ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、**情報提供**、**警告**、**停止** などです。</span><span class="sxs-lookup"><span data-stu-id="2ee89-114">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="2ee89-115">「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-115">See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="2ee89-116">`prompt` &#8212; ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-116">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="2ee89-117">「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-117">See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="2ee89-118">`ignoreBlanks` &#8212; データの入力規則を範囲内の空白セルに適用するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-118">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="2ee89-119">既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-119">Defaults to `true`.</span></span>
- <span data-ttu-id="2ee89-120">`type` &#8212; WholeNumber、Date、TextLength などの入力規則のタイプの読み取り専用 ID です。これは `rule` プロパティを設定すると間接的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-120">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="2ee89-121">プログラムによって追加されたデータの入力規則は、手動で追加したデータの入力規則と同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-121">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="2ee89-122">具体的に言うと、データの入力規則は、ユーザーがセルに値を直接入力した場合、またはブックの別の場所からセルをコピーして貼り付けたときに、**値**の貼り付けオプションを選択した場合にのみトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-122">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="2ee89-123">ユーザーがセルをコピーしてデータの入力規則のある範囲内に単に貼り付けた場合は、データの入力規則はトリガーされません。</span><span class="sxs-lookup"><span data-stu-id="2ee89-123">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="2ee89-124">入力規則を作成する</span><span class="sxs-lookup"><span data-stu-id="2ee89-124">Creating validation rules</span></span>

<span data-ttu-id="2ee89-125">範囲にデータの入力規則を追加するには、コードで `Range.dataValidation` にある `DataValidation` オブジェクトの `rule` プロパティを設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-125">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="2ee89-126">これには 7 つの省略可能なプロパティを持つ [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) オブジェクトが必要です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-126">This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="2ee89-127">*`DataValidationRule` オブジェクトにはこれらのプロパティの 1 つのみを設定できます。*</span><span class="sxs-lookup"><span data-stu-id="2ee89-127">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="2ee89-128">設定したプロパティにより、入力規則のタイプが決まります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-128">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="2ee89-129">Basic および DateTime 入力規則のタイプ</span><span class="sxs-lookup"><span data-stu-id="2ee89-129">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="2ee89-130">最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則のタイプ) は、その値として [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) オブジェクトをとります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-130">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="2ee89-131">`wholeNumber` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と整数が必要です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-131">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="2ee89-132">`decimal` &#8212; `BasicDataValidation` オブジェクトで指定された他の任意の入力規則と 10 進数が必要です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-132">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="2ee89-133">`textLength` &#8212; `BasicDataValidation`オブジェクトの入力規則の詳細をセルの値の *長さ* に適用します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-133">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="2ee89-134">次に、入力規則を作成する例を示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-134">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="2ee89-135">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-135">Note the following about this code:</span></span>

- <span data-ttu-id="2ee89-136">`operator` は二項演算子 "GreaterThan" です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-136">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="2ee89-137">二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値は左側のオペランドになり、`formula1` で指定された値は右側のオペランドになります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-137">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="2ee89-138">そのため、この規則では、0 より大きい整数のみが有効になります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-138">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="2ee89-139">`formula1` はハードコーディングされた値です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-139">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="2ee89-140">コーディングの時点でその正しい値がわからない場合は、その値に Excel の数式 (文字列) を使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-140">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="2ee89-141">たとえば、"=A3" や "=SUM(A4,B5)" を `formula1` の値にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-141">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="2ee89-142">その他の二項演算子のリストについては、「[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-142">See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="2ee89-143">また、2 個の三項演算子、"Between" と "NotBetween" もあります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-143">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="2ee89-144">これらを使用するには、省略可能な `formula2` プロパティを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-144">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="2ee89-145">`formula1` and `formula2`の値はバウンディング オペランドです。</span><span class="sxs-lookup"><span data-stu-id="2ee89-145">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="2ee89-146">ユーザーがセルに入力しようとする値は、第三の (評価済み) オペランドです。</span><span class="sxs-lookup"><span data-stu-id="2ee89-146">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="2ee89-147">"Between" 演算子の使用例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-147">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="2ee89-148">次の 2 つのルール プロパティは、値として [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをとります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-148">The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="2ee89-149">`DateTimeDataValidation` オブジェクトは `BasicDataValidation` と同様に構成されています。つまり、プロパティ `formula1`、`formula2`、および `operator` があり、同じ方法で使用されます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-149">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="2ee89-150">ただし、数式プロパティで数値が使用できない代わりに [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点が異なります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-150">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="2ee89-151">次に、2018 年 4 月の最初の週の日付として有効な値を定義する例を示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-151">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="2ee89-152">リスト入力規則のタイプ</span><span class="sxs-lookup"><span data-stu-id="2ee89-152">List validation rule type</span></span>

<span data-ttu-id="2ee89-153">有限リストからの値のみを有効な値として指定するには、`DataValidationRule` オブジェクトの `list` プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-153">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="2ee89-154">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-154">The following is an example.</span></span> <span data-ttu-id="2ee89-155">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-155">Note the following about this code:</span></span>

- <span data-ttu-id="2ee89-156">"Names" という名前のワークシートがあり、"A1:A3" の範囲の値が名前になっていると仮定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-156">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="2ee89-157">`source` プロパティは有効な値のリストを指定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-157">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="2ee89-158">文字列引数は名前を含む範囲を参照します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-158">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="2ee89-159">"Sue, Ricky, Liz" など、カンマで区切られたリストを割り当てることもできます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-159">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="2ee89-160">`inCellDropDown` プロパティは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-160">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="2ee89-161">`true` に設定した場合、ドロップダウンには `source` からの値のリストが表示されます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-161">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="2ee89-162">カスタムの入力規則のタイプ</span><span class="sxs-lookup"><span data-stu-id="2ee89-162">Custom validation rule type</span></span>

<span data-ttu-id="2ee89-163">カスタムの入力規則式を指定するには、`DataValidationRule` オブジェクトの `custom` プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-163">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="2ee89-164">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-164">The following is an example.</span></span> <span data-ttu-id="2ee89-165">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-165">Note the following about this code:</span></span>

- <span data-ttu-id="2ee89-166">ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列がある、2 列のテーブルがあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-166">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="2ee89-167">**Comments** 列の冗長性を軽減するために、アスリート名を含むデータを無効にします。</span><span class="sxs-lookup"><span data-stu-id="2ee89-167">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="2ee89-168">`SEARCH(A2,B2)` は A2 内の文字列の開始位置 (B2 内の文字列での) を返します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-168">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="2ee89-169">A2 が B2 に含まれていない場合は数値を返しません。</span><span class="sxs-lookup"><span data-stu-id="2ee89-169">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="2ee89-170">`ISNUMBER()` はブール値を返します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-170">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="2ee89-171">そのため、`formula` プロパティは、**コメント**列の有効なデータが**アスリート名**列内の文字列を含まないデータであることを示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-171">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="2ee89-172">入力規則のエラー アラートを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee89-172">Create validation error alerts</span></span>

<span data-ttu-id="2ee89-173">ユーザーがセルに無効なデータを入力しようとした際に表示される、カスタムのエラー アラートを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-173">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="2ee89-174">次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-174">The following is a simple example.</span></span> <span data-ttu-id="2ee89-175">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-175">Note the following about this code:</span></span>

- <span data-ttu-id="2ee89-176">`style` プロパティは、ユーザーが情報アラート、警告、または "停止" アラートを取得するかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-176">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="2ee89-177">ユーザーによる無効なデータの追加を実際に防止するのは `Stop` のみです。</span><span class="sxs-lookup"><span data-stu-id="2ee89-177">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="2ee89-178">いずれにせよ、`Warning` と `Information` のポップアップには、ユーザーが無効なデータを入力できるオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="2ee89-178">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="2ee89-179">`showAlert` プロパティの既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="2ee89-179">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="2ee89-180">これは、カスタムの `Stop` `showAlert` `false` メッセージ、タイトル、およびスタイルを設定または設定するカスタム通知を作成しない限り、Excel は汎用的な通知 (種類の) をポップアップ表示することを意味します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-180">This means that Excel will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="2ee89-181">このコードでは、カスタムのメッセージとタイトルを設定します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-181">This code sets a custom message and title.</span></span>

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

<span data-ttu-id="2ee89-182">詳細については、「[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-182">For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="2ee89-183">入力規則のプロンプトを作成する</span><span class="sxs-lookup"><span data-stu-id="2ee89-183">Create validation prompts</span></span>

<span data-ttu-id="2ee89-184">ユーザーがデータの入力規則が適用されたセルの上でカーソルを動かすか、データの入力規則が適用されたセルを選択した場合に表示される、説明用のダイアログを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-184">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="2ee89-185">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-185">The following is an example:</span></span>

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

<span data-ttu-id="2ee89-186">詳細については、「[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2ee89-186">For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="2ee89-187">範囲からデータの入力規則を削除する</span><span class="sxs-lookup"><span data-stu-id="2ee89-187">Remove data validation from a range</span></span>

<span data-ttu-id="2ee89-188">範囲からデータの入力規則を削除するには、[Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="2ee89-188">To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="2ee89-189">消去する範囲とデータの入力規則を追加した範囲とが、まったく同じになる必要はありません。</span><span class="sxs-lookup"><span data-stu-id="2ee89-189">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="2ee89-190">同じでない場合は、2 つの範囲に重複するセルがあると、それらのセルのみが消去されます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-190">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="2ee89-191">範囲からデータの入力規則を削除すると、ユーザーが手動で範囲に追加したデータの入力規則も削除されます。</span><span class="sxs-lookup"><span data-stu-id="2ee89-191">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="2ee89-192">関連項目</span><span class="sxs-lookup"><span data-stu-id="2ee89-192">See also</span></span>

- [<span data-ttu-id="2ee89-193">Office アドインでの Excel JavaScript オブジェクトモデル</span><span class="sxs-lookup"><span data-stu-id="2ee89-193">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="2ee89-194">DataValidation Object (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="2ee89-194">DataValidation Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="2ee89-195">Range オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="2ee89-195">Range Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.range)
