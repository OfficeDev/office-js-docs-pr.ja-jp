---
ms.date: 07/15/2019
description: Excel 範囲、省略可能なパラメーター、呼び出しコンテキストなど、カスタム関数内でさまざまなパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: e5b75b098d64d5998b0393d5995896f0289337fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771424"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="3dedc-103">カスタム関数のパラメータオプション</span><span class="sxs-lookup"><span data-stu-id="3dedc-103">Custom functions parameter options</span></span>

<span data-ttu-id="3dedc-104">カスタム関数は、さまざまなパラメーターのオプションを使用して構成できます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="3dedc-105">オプションのパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-105">Optional parameters</span></span>

<span data-ttu-id="3dedc-106">通常のパラメーターは必須ですが、省略可能なパラメーターは必須ではありません。</span><span class="sxs-lookup"><span data-stu-id="3dedc-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="3dedc-107">ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="3dedc-108">次の例では、add 関数で3番目の番号を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="3dedc-109">この関数は Excel `=CONTOSO.ADD(first, second, [third])`のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="3dedc-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="3dedc-110">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="3dedc-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="3dedc-111">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> <span data-ttu-id="3dedc-112">省略可能なパラメーターに値が指定されていない場合、 `null`Excel によって値が割り当てられます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="3dedc-113">これは、TypeScript の既定の初期化されたパラメーターが期待どおりに動作しないことを意味します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="3dedc-114">そのため、この構文`function add(first:number, second:number, third=0):number`は0に初期化`third`されないため、使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="3dedc-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="3dedc-115">その代わりに、前の例のように TypeScript 構文を使用します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="3dedc-116">1つ以上のオプションパラメーターを含む関数を定義するときは、省略可能なパラメーターが null の場合の処理を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dedc-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="3dedc-117">次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="3dedc-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="3dedc-118">`zipCode`パラメーターが null の場合、既定値はに`98052`設定されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="3dedc-119">`dayOfWeek`パラメーターが null の場合は、水曜日に設定します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="3dedc-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="3dedc-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="3dedc-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="3dedc-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a><span data-ttu-id="3dedc-122">範囲パラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-122">Range parameters</span></span>

<span data-ttu-id="3dedc-123">カスタム関数は、入力パラメーターとして範囲のセルデータを受け入れることができます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="3dedc-124">関数は、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-124">A function can also return a range of data.</span></span> <span data-ttu-id="3dedc-125">Excel は、セルデータの範囲を2次元配列として渡します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="3dedc-126">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="3dedc-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="3dedc-127">次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="3dedc-128">この関数の JSON メタデータでは、パラメーターの`type`プロパティがに`matrix`設定されていることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="3dedc-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a><span data-ttu-id="3dedc-129">繰り返しパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-129">Repeating parameters</span></span>

<span data-ttu-id="3dedc-130">繰り返しパラメーターを使用すると、ユーザーは関数に一連のオプションの引数を入力できます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-130">A repeating parameter allows a user to enter a series of optional of arguments to a function.</span></span> <span data-ttu-id="3dedc-131">関数が呼び出されると、パラメーターの配列に値が提供されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-131">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="3dedc-132">パラメーター名が数値で終わる場合、各引数は、などの数値をインクリメントし`ADD(number1, [number2], [number3],…)`ます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-132">If the parameter name ends with a number, each argument will increment the number, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="3dedc-133">これは、Excel の組み込み関数で使用される規則に一致します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-133">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="3dedc-134">次の関数は、合計数、セルの住所、および範囲 (入力した場合) を合計します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-134">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

<span data-ttu-id="3dedc-135">この関数は`=CONTOSO.ADD([operands], [operands]...)` 、Excel ブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-135">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="3dedc-136">繰り返し単一値パラメータ</span><span class="sxs-lookup"><span data-stu-id="3dedc-136">Repeating single value parameter</span></span>

<span data-ttu-id="3dedc-137">繰り返し単一の値のパラメーターを使用すると、複数の単一の値を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-137">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="3dedc-138">たとえば、ユーザーは、「ADD (1, B2, 3)」と入力することができます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-138">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="3dedc-139">次の例は、単一の値のパラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3dedc-139">The following sample shows how to declare a single value parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a><span data-ttu-id="3dedc-140">単一範囲のパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-140">Single range parameter</span></span>

<span data-ttu-id="3dedc-141">単精度浮動小数点型 (single) のパラメーターは、技術的には繰り返しパラメーターではありませんが、宣言は繰り返しパラメーターによく似ているので、ここに記載されています。</span><span class="sxs-lookup"><span data-stu-id="3dedc-141">A single range parameter is not technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="3dedc-142">ユーザーには、Excel から1つの範囲が渡される追加 (A2: B3) として表示されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-142">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="3dedc-143">次の例は、1つの range パラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="3dedc-143">The following sample shows how to declare a single range parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a><span data-ttu-id="3dedc-144">繰り返し範囲のパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-144">Repeating range parameter</span></span>

<span data-ttu-id="3dedc-145">繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-145">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="3dedc-146">たとえば、ユーザーは ADD (5、B2、C3、8、E5: E8) を入力することができます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-146">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="3dedc-147">通常、繰り返し範囲は3次元の`number[][][]`行列として型で指定されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-147">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="3dedc-148">サンプルについては、繰り返しパラメーター (#repeating パラメーター) の主なサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dedc-148">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="3dedc-149">繰り返しパラメーターの宣言</span><span class="sxs-lookup"><span data-stu-id="3dedc-149">Declaring repeating parameters</span></span>
<span data-ttu-id="3dedc-150">Typescript で、パラメーターが多次元であることを示します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-150">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="3dedc-151">たとえば、 `ADD(values: number[])`は1次元配列`ADD(values:number[][])`を示し、2次元配列というように指定します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-151">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="3dedc-152">JavaScript では、 `@param values {number[]}` 1 次元`@param <name> {number[][]}`配列、2次元配列、およびその他の次元で使用します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-152">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="3dedc-153">手動で作成した JSON では、パラメーターが JSON `"repeating": true`ファイルで指定されていること、およびパラメーターがにマークさ`"dimensionality”: matrix`れていることを確認することを確認してください。</span><span class="sxs-lookup"><span data-stu-id="3dedc-153">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality”: matrix`.</span></span>

>[!NOTE]
><span data-ttu-id="3dedc-154">繰り返しパラメーターを含む関数には、最後のパラメーターとして、呼び出しパラメーターが自動的に含まれています。</span><span class="sxs-lookup"><span data-stu-id="3dedc-154">Functions containing repeating parameters automatically contain an invocation parameter as the last parameter.</span></span> <span data-ttu-id="3dedc-155">呼び出しパラメーターの詳細については、以下のセクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dedc-155">For more information on invocation parameters, see the following section.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="3dedc-156">呼び出しパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-156">Invocation parameter</span></span>

<span data-ttu-id="3dedc-157">すべてのカスタム関数には、 `invocation`最後の引数として引数が自動的に渡されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-157">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="3dedc-158">この引数は、呼び出し元のセルのアドレスなど、追加のコンテキストを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-158">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="3dedc-159">また、関数[をキャンセル](custom-functions-web-reqs.md#make-a-streaming-function)する関数ハンドラーなど、Excel に情報を送信するために使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-159">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="3dedc-160">パラメーターを宣言しない場合でも、カスタム関数にはこのパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="3dedc-160">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="3dedc-161">この引数は、Excel のユーザーには表示されません。</span><span class="sxs-lookup"><span data-stu-id="3dedc-161">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="3dedc-162">カスタム関数でを使用`invocation`する場合は、最後のパラメーターとして宣言します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-162">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="3dedc-163">次のコードサンプルでは、 `invocation`コンテキストが参照に対して明示的に指定されています。</span><span class="sxs-lookup"><span data-stu-id="3dedc-163">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
```

<span data-ttu-id="3dedc-164">このパラメーターを使用すると、呼び出し元のセルのコンテキストを取得できます。これは、[カスタム関数を呼び出すセルのアドレスを検索](#addressing-cells-context-parameter)するなどの一部のシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-164">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="3dedc-165">アドレス指定セルのコンテキストパラメーター</span><span class="sxs-lookup"><span data-stu-id="3dedc-165">Addressing cell's context parameter</span></span>

<span data-ttu-id="3dedc-166">場合によっては、カスタム関数を呼び出したセルのアドレスを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dedc-166">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="3dedc-167">これは、次のシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-167">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="3dedc-168">範囲の書式設定: セルのアドレスをキーとして使用し、データを保存します[。](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data)</span><span class="sxs-lookup"><span data-stu-id="3dedc-168">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="3dedc-169">Excel で [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) を使用して`OfficeRuntime.storage` からキーを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-169">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="3dedc-170">キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `OfficeRuntime.storage` に格納されているキャッシュされた値を表示します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-170">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="3dedc-171">調整: セル アドレスを使用して元のセルを検出し、処理が発生している場所での調整を行えます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-171">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="3dedc-172">関数内のアドレス指定セルのコンテキストを要求するには、次の例のように、関数を使用してセルのアドレスを検索する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3dedc-172">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="3dedc-173">セルのアドレスに関する情報は、関数のコメント`@requiresAddress`にタグ付けされている場合にのみ公開されます。</span><span class="sxs-lookup"><span data-stu-id="3dedc-173">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
```

<span data-ttu-id="3dedc-174">既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="3dedc-174">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="3dedc-175">たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。</span><span class="sxs-lookup"><span data-stu-id="3dedc-175">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3dedc-176">次のステップ</span><span class="sxs-lookup"><span data-stu-id="3dedc-176">Next steps</span></span>

<span data-ttu-id="3dedc-177">カスタム関数の[状態を保存](custom-functions-save-state.md)する方法、または[カスタム関数で揮発性の値](custom-functions-volatile.md)を使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="3dedc-177">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3dedc-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="3dedc-178">See also</span></span>

* [<span data-ttu-id="3dedc-179">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="3dedc-179">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="3dedc-180">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="3dedc-180">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3dedc-181">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="3dedc-181">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="3dedc-182">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="3dedc-182">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="3dedc-183">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="3dedc-183">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)