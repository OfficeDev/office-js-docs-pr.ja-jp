---
title: Excel の範囲にデータの入力規則を追加する
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 7e545ccca01a12257f4083f19135a320b2693190
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967691"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="c7641-102">Excel 範囲にデータの入力規則を追加する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c7641-102">Add data validation to Excel ranges (Preview)</span></span>

<span data-ttu-id="c7641-103">Excel JavaScript ライブラリには、ワークブックに、表、列、行、その他の範囲に自動データ入力規則をアドインで追加できる API が用意されています。</span><span class="sxs-lookup"><span data-stu-id="c7641-103">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="c7641-104">データの入力規則の概念と用語を把握するには、ユーザーが Excel UI によってデータの入力規則を追加する方法に関する次の記事をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c7641-104">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="c7641-105">セルに対するデータの入力規則の適用</span><span class="sxs-lookup"><span data-stu-id="c7641-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="c7641-106">データの入力規則の詳細</span><span class="sxs-lookup"><span data-stu-id="c7641-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="c7641-107">Excel でのデータの入力規則の説明と例</span><span class="sxs-lookup"><span data-stu-id="c7641-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="c7641-108">データの入力規則のプログラムによる制御</span><span class="sxs-lookup"><span data-stu-id="c7641-108">Programmatic control of data validation</span></span>

<span data-ttu-id="c7641-109">プロパティは、[データの入力規則](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) オブジェクトを取得しますが、これが Excel でデータの入力規則をプログラムにより制御するためのエントリポイントとなります。`Range.dataValidation`</span><span class="sxs-lookup"><span data-stu-id="c7641-109">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="c7641-110">オブジェクトには、次のような 5つのプロパティがあります。`DataValidation`</span><span class="sxs-lookup"><span data-stu-id="c7641-110">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="c7641-111">`rule` — 範囲の有効データの構成要素を定義します。</span><span class="sxs-lookup"><span data-stu-id="c7641-111">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="c7641-112">「[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-112">See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="c7641-113">`errorAlert` — ユーザーが無効なデータを入力した場合にエラーがポップアップ表示されるかどうかを指定し、アラートのテキスト、タイトル、スタイルを定義します。たとえば、 **[情報提供]**、 **[警告]**、**[停止]** などです。</span><span class="sxs-lookup"><span data-stu-id="c7641-113">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="c7641-114">「[DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-114">See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="c7641-115">`prompt` — ユーザーが範囲の上にカーソルを動かすとダイアログが表示されるかどうかを指定し、表示されるダイアログ メッセージを定義します。</span><span class="sxs-lookup"><span data-stu-id="c7641-115">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="c7641-116">「[DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-116">See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="c7641-117">`ignoreBlanks` — データの入力規則のルールを範囲内の空白セルに適用するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="c7641-117">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="c7641-118">既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="c7641-118">Defaults to `true`.</span></span>
- <span data-ttu-id="c7641-119">`type` —  WholeNumber、Date、TextLength などの入力規則タイプの読み取り専用  ID です。これは `rule` プロパティが設定されると間接的に設定されます。</span><span class="sxs-lookup"><span data-stu-id="c7641-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="c7641-120">プログラムにより追加されたデータの入力規則は、手動で追加されたデータの入力規則と同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="c7641-120">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="c7641-121">特に、データの入力規則は、ユーザーがセルに値を直接入力したり、ワークブックの別の場所からセルをコピーして貼り付けるときに 「**値**の 貼り付け」オプションを選んだりした場合にのみトリガーされます。</span><span class="sxs-lookup"><span data-stu-id="c7641-121">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="c7641-122">ユーザーがセルをコピーし、データの入力規則が設定された範囲に単にペーストする場合、入力規則はトリガーされません。</span><span class="sxs-lookup"><span data-stu-id="c7641-122">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="c7641-123">入力規則ルールを作成する</span><span class="sxs-lookup"><span data-stu-id="c7641-123">Creating validation rules</span></span>

<span data-ttu-id="c7641-124">範囲にデータの入力規則を追加するには、コードで `rule` にある `DataValidation` オブジェクトの `Range.dataValidation` プロパティを設定する必要があります 。</span><span class="sxs-lookup"><span data-stu-id="c7641-124">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="c7641-125">これは、7 つのオプション プロパティのある [DataValidationRule](https://docs.microsoft.com/javascript/api/excel?view=office-js) オブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c7641-125">This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel?view=office-js) object which has seven optional properties.</span></span> <span data-ttu-id="c7641-126">*どの `DataValidationRule` オブジェクトでも、これらの特性が 1 つ以上表示されることはありません。*</span><span class="sxs-lookup"><span data-stu-id="c7641-126">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="c7641-127">含めるプロパティによって、入力規則のタイプが決まります。</span><span class="sxs-lookup"><span data-stu-id="c7641-127">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="c7641-128">Basic および DateTime 入力規則ルールのタイプ</span><span class="sxs-lookup"><span data-stu-id="c7641-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="c7641-129">最初の 3 つの `DataValidationRule` プロパティ (つまり、入力規則ルール タイプ) は、[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel) オブジェクトをその値として取得します。</span><span class="sxs-lookup"><span data-stu-id="c7641-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel) object as their value.</span></span>

- <span data-ttu-id="c7641-130">`wholeNumber` ー  `BasicDataValidation` オブジェクトで指定された他の妥当性確認に加えて整数を必要とします。</span><span class="sxs-lookup"><span data-stu-id="c7641-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="c7641-131">`decimal` ー  `BasicDataValidation` オブジェクトで指定された他の妥当性確認に加えて、10進数が必要です。</span><span class="sxs-lookup"><span data-stu-id="c7641-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="c7641-132">`textLength` ー  `BasicDataValidation` オブジェクトの妥当性確認の詳細をセルの値の *長さ* に適用します。</span><span class="sxs-lookup"><span data-stu-id="c7641-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="c7641-133">次に、入力規則のルールを作成する例を示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-133">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="c7641-134">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-134">Note the following about this code:</span></span>

- <span data-ttu-id="c7641-135">は二項演算子「GreaterThan」です。`operator`</span><span class="sxs-lookup"><span data-stu-id="c7641-135">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="c7641-136">二項演算子を使用する際は必ず、ユーザーがセルに入力しようとする値が左側のオペランドになり、 `formula1` で指定された値が右側のオペランドになります。</span><span class="sxs-lookup"><span data-stu-id="c7641-136">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="c7641-137">したがって、このルールでは、0 より大きな整数だけが有効です。</span><span class="sxs-lookup"><span data-stu-id="c7641-137">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="c7641-138">は、ハードコーディングされた数字です。`formula1`</span><span class="sxs-lookup"><span data-stu-id="c7641-138">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="c7641-139">コード時にどの値にすべきかわからない場合は、値の Excel 式を (文字列として) 使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="c7641-139">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="c7641-140">たとえば、「= A3」および「= SUM（A4、B5）」も、 `formula1`の値にできます。</span><span class="sxs-lookup"><span data-stu-id="c7641-140">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="c7641-141">他の二項演算子のリストについては、「[BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="c7641-142">また、「Between」と「NotBetween」の2つの三項演算子もあります。</span><span class="sxs-lookup"><span data-stu-id="c7641-142">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="c7641-143">これらを使用するには、オプションの `formula2` プロパティを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c7641-143">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="c7641-144">と `formula2` 値はバウンディング オペランドです。`formula1`</span><span class="sxs-lookup"><span data-stu-id="c7641-144">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="c7641-145">ユーザーがセルに入力しようとする値は、3 番目の (評価された) オペランドです。</span><span class="sxs-lookup"><span data-stu-id="c7641-145">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="c7641-146">以下は「Between」演算子の使用例です。</span><span class="sxs-lookup"><span data-stu-id="c7641-146">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="c7641-147">次の 2 つのルール プロパティは、 [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) オブジェクトをその値として取得します。</span><span class="sxs-lookup"><span data-stu-id="c7641-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="c7641-148">オブジェクトは `BasicDataValidation` と構成が似ています。つまり、プロパティ `formula1`、 `formula2`、`operator`が備わっており、同じ方法で使われます。`DateTimeDataValidation`</span><span class="sxs-lookup"><span data-stu-id="c7641-148">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="c7641-149">違うのは、数式プロパティで数値を使えませんが、 [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) 文字列 (または Excel の式) を入力できる点です。</span><span class="sxs-lookup"><span data-stu-id="c7641-149">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="c7641-150">以下は、2018 年 4 月の第 1 週の日付として有効な値を定義する例です。</span><span class="sxs-lookup"><span data-stu-id="c7641-150">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="c7641-151">リスト入力規則ルール タイプ</span><span class="sxs-lookup"><span data-stu-id="c7641-151">List validation rule type</span></span>

<span data-ttu-id="c7641-152">有限リストからの値のみが有効な値であるように指定するには、`list` オブジェクトの `DataValidationRule` プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c7641-152">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="c7641-153">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-153">The following is an example.</span></span> <span data-ttu-id="c7641-154">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-154">Note the following about this code:</span></span>

- <span data-ttu-id="c7641-155">「Names」という名前のワークシートがあり、「A1：A3」の範囲の値が名前であることが前提です。</span><span class="sxs-lookup"><span data-stu-id="c7641-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="c7641-156">プロパティは、有効な値のリストを指定します。`source`</span><span class="sxs-lookup"><span data-stu-id="c7641-156">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="c7641-157">名前を含んだ範囲が割り当てられています。</span><span class="sxs-lookup"><span data-stu-id="c7641-157">The range with the names has been assigned to it.</span></span> <span data-ttu-id="c7641-158">コンマ区切りのリストを割り当てることもできます。たとえば、「Sue、Ricky、Liz」です。</span><span class="sxs-lookup"><span data-stu-id="c7641-158">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="c7641-159">プロパティでは、ユーザーがセルを選択したときにセルにドロップダウン コントロールを表示するかどうかを指定します。`inCellDropDown`</span><span class="sxs-lookup"><span data-stu-id="c7641-159">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="c7641-160">に設定されている場合 、ドロップダウンには `source` からの値のリストが表示されます。`true`</span><span class="sxs-lookup"><span data-stu-id="c7641-160">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="c7641-161">カスタムの入力規則ルール タイプ</span><span class="sxs-lookup"><span data-stu-id="c7641-161">Custom validation rule type</span></span>

<span data-ttu-id="c7641-162">カスタムの入力規則式を指定するには、`custom` オブジェクトの `DataValidationRule` プロパティ を使用します。</span><span class="sxs-lookup"><span data-stu-id="c7641-162">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="c7641-163">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-163">The following is an example.</span></span> <span data-ttu-id="c7641-164">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-164">Note the following about this code:</span></span>

- <span data-ttu-id="c7641-165">ワークシートの A 列と B 列に **Athlete Name** と **Comments** という列をもつ 2 列のテーブルがあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="c7641-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="c7641-166">**Comments** 列の冗長性を軽減するには、アスリート名を含むデータを無効にします。</span><span class="sxs-lookup"><span data-stu-id="c7641-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="c7641-167">`SEARCH(A2,B2)` B2 の文字列での A2 の文字列の開始位置を返します。</span><span class="sxs-lookup"><span data-stu-id="c7641-167">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="c7641-168">B2 に A2 が含まれていない場合は、数値は返されません。</span><span class="sxs-lookup"><span data-stu-id="c7641-168">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="c7641-169">`ISNUMBER()` ブール値を返します。</span><span class="sxs-lookup"><span data-stu-id="c7641-169">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="c7641-170">したがって `formula` プロパティは、 **Comment** 列の有効なデータは、 **Athlete Name** 列の文字列を含まないデータであることを示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-170">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="c7641-171">入力規則エラー アラートを作成する</span><span class="sxs-lookup"><span data-stu-id="c7641-171">Create validation error alerts</span></span>

<span data-ttu-id="c7641-172">ユーザーがセルに無効なデータを入力しようとすると表示されるカスタムのエラー アラートを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c7641-172">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="c7641-173">次に簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-173">The following is a simple example:</span></span> <span data-ttu-id="c7641-174">このコードについては、次の点に注意してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-174">Note the following about this code:</span></span>

- <span data-ttu-id="c7641-175">プロパティは、ユーザーが情報アラート、警告、または「停止」アラートを取得するかどうかを決定します。`style`</span><span class="sxs-lookup"><span data-stu-id="c7641-175">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="c7641-176">実際のところ、ユーザーが無効なデータを追加できないようにするのは、`Stop` のみです。</span><span class="sxs-lookup"><span data-stu-id="c7641-176">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="c7641-177">と `Information` のポップアップには、設定にかかわらずユーザーが無効なデータを入力できるオプションがあります。`Warning`</span><span class="sxs-lookup"><span data-stu-id="c7641-177">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="c7641-178">プロパティの既定値は `true` です。`showAlert`</span><span class="sxs-lookup"><span data-stu-id="c7641-178">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="c7641-179">つまり、`showAlert` を `false` に設定するか、カスタムのメッセージ、タイトル、スタイルを設定するカスタム アラートを作成するかしない限り、Excel ホストは (`Stop` タイプの) 汎用アラートをポップアップ表示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-179">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="c7641-180">このコードはカスタム メッセージとタイトルを設定します。</span><span class="sxs-lookup"><span data-stu-id="c7641-180">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="c7641-181">詳細情報については、「 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="c7641-182">入力規則プロンプトを作成する</span><span class="sxs-lookup"><span data-stu-id="c7641-182">Create validation prompts</span></span>

<span data-ttu-id="c7641-183">ユーザーがデータ入力規則が適用されたセルの上でカーソルを動かすか、またはこのようなセルを選択するかしたときに表示される説明用ダイアログを作成できます。</span><span class="sxs-lookup"><span data-stu-id="c7641-183">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="c7641-184">例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c7641-184">The following is an example:</span></span>

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

<span data-ttu-id="c7641-185">詳細情報については、「 [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c7641-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="c7641-186">範囲からデータ入力規則を削除する</span><span class="sxs-lookup"><span data-stu-id="c7641-186">Remove data validation from a range</span></span>

<span data-ttu-id="c7641-187">範囲からデータ入力規則を削除するには、[Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c7641-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="c7641-188">削除する範囲は、データ入力規則を追加した範囲とまったく同じ範囲でなくてもかまいません。</span><span class="sxs-lookup"><span data-stu-id="c7641-188">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="c7641-189">範囲が同じでない場合は、2 つの範囲でオーバーラップしているセルがあれば、そのようなセルのみが削除されます。</span><span class="sxs-lookup"><span data-stu-id="c7641-189">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="c7641-190">範囲からデータ入力規則を削除すると、ユーザーが手動で範囲に追加したデータ入力規則も削除されます。</span><span class="sxs-lookup"><span data-stu-id="c7641-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="c7641-191">関連項目</span><span class="sxs-lookup"><span data-stu-id="c7641-191">See also</span></span>

- [<span data-ttu-id="c7641-192">Excel JavaScript API の中心概念</span><span class="sxs-lookup"><span data-stu-id="c7641-192">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c7641-193">DataValidation オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="c7641-193">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="c7641-194">Range オブジェクト (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="c7641-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
