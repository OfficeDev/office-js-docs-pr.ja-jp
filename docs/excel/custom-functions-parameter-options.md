---
ms.date: 11/06/2020
description: Excel 範囲、省略可能なパラメーター、呼び出しコンテキストなど、カスタム関数内でさまざまなパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: 0a803a4d41354530584b25d2bf9df944af430909
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071621"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="da219-103">カスタム関数のパラメータオプション</span><span class="sxs-lookup"><span data-stu-id="da219-103">Custom functions parameter options</span></span>

<span data-ttu-id="da219-104">カスタム関数は、さまざまなパラメーターオプションを使用して構成できます。</span><span class="sxs-lookup"><span data-stu-id="da219-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="da219-105">オプションのパラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-105">Optional parameters</span></span>

<span data-ttu-id="da219-106">ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。</span><span class="sxs-lookup"><span data-stu-id="da219-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="da219-107">次の例では、add 関数で3番目の番号を追加することもできます。</span><span class="sxs-lookup"><span data-stu-id="da219-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="da219-108">この関数 `=CONTOSO.ADD(first, second, [third])` は Excel のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="da219-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="da219-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="da219-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="da219-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="da219-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="da219-111">省略可能なパラメーターに値が指定されていない場合、Excel によって値が割り当てられ `null` ます。</span><span class="sxs-lookup"><span data-stu-id="da219-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="da219-112">これは、TypeScript の既定の初期化されたパラメーターが期待どおりに動作しないことを意味します。</span><span class="sxs-lookup"><span data-stu-id="da219-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="da219-113">この構文は `function add(first:number, second:number, third=0):number` 0 に初期化されないため、使用しないで `third` ください。</span><span class="sxs-lookup"><span data-stu-id="da219-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="da219-114">その代わりに、前の例のように TypeScript 構文を使用します。</span><span class="sxs-lookup"><span data-stu-id="da219-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="da219-115">オプションのパラメーターが1つ以上含まれる関数を定義するときは、オプションのパラメーターが null の場合に何が起こるかを指定します。</span><span class="sxs-lookup"><span data-stu-id="da219-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="da219-116">次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="da219-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="da219-117">`zipCode`パラメーターが null の場合、既定値はに設定され `98052` ます。</span><span class="sxs-lookup"><span data-stu-id="da219-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="da219-118">`dayOfWeek`パラメーターが null の場合は、水曜日に設定します。</span><span class="sxs-lookup"><span data-stu-id="da219-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="da219-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="da219-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="da219-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="da219-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="da219-121">範囲パラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-121">Range parameters</span></span>

<span data-ttu-id="da219-122">カスタム関数は、入力パラメーターとして範囲のセルデータを受け入れることができます。</span><span class="sxs-lookup"><span data-stu-id="da219-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="da219-123">関数は、データの範囲を返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="da219-123">A function can also return a range of data.</span></span> <span data-ttu-id="da219-124">Excel は、セルデータの範囲を2次元配列として渡します。</span><span class="sxs-lookup"><span data-stu-id="da219-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="da219-125">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="da219-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="da219-126">次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="da219-126">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="da219-127">この関数の JSON メタデータでは、パラメーターの `type` プロパティがに設定されていることに注意して `matrix` ください。</span><span class="sxs-lookup"><span data-stu-id="da219-127">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

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

## <a name="repeating-parameters"></a><span data-ttu-id="da219-128">繰り返しパラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-128">Repeating parameters</span></span>

<span data-ttu-id="da219-129">繰り返しパラメーターを使用すると、ユーザーは関数に一連のオプションの引数を入力できます。</span><span class="sxs-lookup"><span data-stu-id="da219-129">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="da219-130">関数が呼び出されると、パラメーターの配列に値が提供されます。</span><span class="sxs-lookup"><span data-stu-id="da219-130">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="da219-131">パラメーター名が数値で終わる場合、各引数の値は、などの増分で増加し `ADD(number1, [number2], [number3],…)` ます。</span><span class="sxs-lookup"><span data-stu-id="da219-131">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="da219-132">これは、Excel の組み込み関数で使用される規則に一致します。</span><span class="sxs-lookup"><span data-stu-id="da219-132">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="da219-133">次の関数は、合計数、セルの住所、および範囲 (入力した場合) を合計します。</span><span class="sxs-lookup"><span data-stu-id="da219-133">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="da219-134">この関数は `=CONTOSO.ADD([operands], [operands]...)` 、Excel ブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="da219-134">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="da219-135">繰り返し単一値パラメータ</span><span class="sxs-lookup"><span data-stu-id="da219-135">Repeating single value parameter</span></span>

<span data-ttu-id="da219-136">繰り返し単一の値のパラメーターを使用すると、複数の単一の値を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="da219-136">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="da219-137">たとえば、ユーザーは、「ADD (1, B2, 3)」と入力することができます。</span><span class="sxs-lookup"><span data-stu-id="da219-137">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="da219-138">次の例は、単一の値のパラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="da219-138">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="da219-139">単一範囲のパラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-139">Single range parameter</span></span>

<span data-ttu-id="da219-140">単精度浮動小数点型 (single) のパラメーターは厳密には繰り返しパラメーターではありませんが、宣言は繰り返しパラメーターによく似ているので、ここに含まれています。</span><span class="sxs-lookup"><span data-stu-id="da219-140">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="da219-141">ユーザーには、Excel から1つの範囲が渡される追加 (A2: B3) として表示されます。</span><span class="sxs-lookup"><span data-stu-id="da219-141">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="da219-142">次の例は、1つの range パラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="da219-142">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="da219-143">繰り返し範囲のパラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-143">Repeating range parameter</span></span>

<span data-ttu-id="da219-144">繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="da219-144">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="da219-145">たとえば、ユーザーは ADD (5、B2、C3、8、E5: E8) を入力することができます。</span><span class="sxs-lookup"><span data-stu-id="da219-145">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="da219-146">通常、繰り返し範囲は `number[][][]` 3 次元の行列として型で指定されます。</span><span class="sxs-lookup"><span data-stu-id="da219-146">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="da219-147">サンプルについては、繰り返しパラメーター (#repeating パラメーター) の主なサンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="da219-147">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="da219-148">繰り返しパラメーターの宣言</span><span class="sxs-lookup"><span data-stu-id="da219-148">Declaring repeating parameters</span></span>
<span data-ttu-id="da219-149">Typescript で、パラメーターが多次元であることを示します。</span><span class="sxs-lookup"><span data-stu-id="da219-149">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="da219-150">たとえば、は  `ADD(values: number[])` 1 次元配列を示し、 `ADD(values:number[][])` 2 次元配列というように指定します。</span><span class="sxs-lookup"><span data-stu-id="da219-150">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="da219-151">JavaScript では、 `@param values {number[]}` 1 次元配列、 `@param <name> {number[][]}` 2 次元配列、およびその他の次元で使用します。</span><span class="sxs-lookup"><span data-stu-id="da219-151">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="da219-152">手動で作成した JSON では、パラメーターが JSON ファイルで指定されていること `"repeating": true` 、およびパラメーターがにマークされていることを確認することを確認してください `"dimensionality": matrix` 。</span><span class="sxs-lookup"><span data-stu-id="da219-152">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="da219-153">呼び出しパラメーター</span><span class="sxs-lookup"><span data-stu-id="da219-153">Invocation parameter</span></span>

<span data-ttu-id="da219-154">すべてのカスタム関数には `invocation` 、最後の引数として引数が自動的に渡されます。</span><span class="sxs-lookup"><span data-stu-id="da219-154">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="da219-155">この引数は、呼び出し元のセルのアドレスなど、追加のコンテキストを取得するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="da219-155">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="da219-156">また、関数 [をキャンセル](custom-functions-web-reqs.md#make-a-streaming-function)する関数ハンドラーなど、Excel に情報を送信するために使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="da219-156">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="da219-157">パラメーターを宣言しない場合でも、カスタム関数にはこのパラメーターがあります。</span><span class="sxs-lookup"><span data-stu-id="da219-157">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="da219-158">この引数は、Excel のユーザーには表示されません。</span><span class="sxs-lookup"><span data-stu-id="da219-158">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="da219-159">カスタム関数でを使用する場合は `invocation` 、最後のパラメーターとして宣言します。</span><span class="sxs-lookup"><span data-stu-id="da219-159">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="da219-160">次のコードサンプルでは、 `invocation` コンテキストが参照に対して明示的に指定されています。</span><span class="sxs-lookup"><span data-stu-id="da219-160">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="da219-161">次の手順</span><span class="sxs-lookup"><span data-stu-id="da219-161">Next steps</span></span>

<span data-ttu-id="da219-162">[カスタム関数で揮発性の値](custom-functions-volatile.md)を使用する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="da219-162">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="da219-163">関連項目</span><span class="sxs-lookup"><span data-stu-id="da219-163">See also</span></span>

* [<span data-ttu-id="da219-164">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="da219-164">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="da219-165">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="da219-165">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="da219-166">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="da219-166">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="da219-167">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="da219-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="da219-168">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="da219-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
