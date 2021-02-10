---
ms.date: 02/04/2021
description: Excel の範囲、オプションのパラメーター、呼び出しコンテキストなど、カスタム関数内で異なるパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: afe6947b1a1b9022a0284535b9ab1d68c9777c14
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173907"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="f3617-103">カスタム関数のパラメーター オプション</span><span class="sxs-lookup"><span data-stu-id="f3617-103">Custom functions parameter options</span></span>

<span data-ttu-id="f3617-104">カスタム関数は、さまざまなパラメーター オプションを使用して構成できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="f3617-105">オプションのパラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-105">Optional parameters</span></span>

<span data-ttu-id="f3617-106">ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="f3617-107">次のサンプルでは、add 関数は必要に応じて 3 番目の数値を追加できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="f3617-108">この関数は Excel と `=CONTOSO.ADD(first, second, [third])` 同様に表示されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f3617-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f3617-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="f3617-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f3617-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="f3617-111">省略可能なパラメーターに値を指定しない場合、Excel によって値が割り当てらされます `null` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="f3617-112">つまり、TypeScript の既定で初期化されたパラメーターは期待通り動作しません。</span><span class="sxs-lookup"><span data-stu-id="f3617-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="f3617-113">この構文は 0 に初期化 `function add(first:number, second:number, third=0):number` されないので使用 `third` してください。</span><span class="sxs-lookup"><span data-stu-id="f3617-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="f3617-114">代わりに、前の例で示した TypeScript 構文を使用します。</span><span class="sxs-lookup"><span data-stu-id="f3617-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="f3617-115">1 つ以上のオプション パラメーターを含む関数を定義する場合は、オプションパラメーターが null の場合の処理を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3617-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="f3617-116">次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="f3617-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="f3617-117">パラメーターが `zipCode` null の場合、既定値はに設定されます `98052` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="f3617-118">パラメーターが `dayOfWeek` null の場合は、水曜日に設定されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="f3617-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="f3617-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="f3617-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="f3617-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="f3617-121">範囲パラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-121">Range parameters</span></span>

<span data-ttu-id="f3617-122">カスタム関数は、入力パラメーターとしてセル データの範囲を受け入れる場合があります。</span><span class="sxs-lookup"><span data-stu-id="f3617-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="f3617-123">関数は、データの範囲を返す場合があります。</span><span class="sxs-lookup"><span data-stu-id="f3617-123">A function can also return a range of data.</span></span> <span data-ttu-id="f3617-124">Excel はセル データの範囲を 2 次元配列として渡します。</span><span class="sxs-lookup"><span data-stu-id="f3617-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="f3617-125">例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。</span><span class="sxs-lookup"><span data-stu-id="f3617-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f3617-126">次の関数はパラメーターを受け入れ、JSDOC 構文はパラメーターのプロパティをこの関数の JSON メタデータ `values` `number[][]` `dimensionality` `matrix` に設定します。</span><span class="sxs-lookup"><span data-stu-id="f3617-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="f3617-127">繰り返しパラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-127">Repeating parameters</span></span>

<span data-ttu-id="f3617-128">繰り返しパラメーターを使用すると、ユーザーは関数に一連のオプションの引数を入力できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="f3617-129">関数が呼び出される場合、値はパラメーターの配列で提供されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="f3617-130">パラメーター名の最後が数値の場合、各引数の数値は徐々に増加します。次に例を示します `ADD(number1, [number2], [number3],…)` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="f3617-131">これは、組み込みの Excel 関数で使用される規則に一致します。</span><span class="sxs-lookup"><span data-stu-id="f3617-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="f3617-132">次の関数は、入力されている場合、数値、セル アドレス、および範囲の合計を合計します。</span><span class="sxs-lookup"><span data-stu-id="f3617-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="f3617-133">この関数は Excel `=CONTOSO.ADD([operands], [operands]...)` ブックに表示されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="f3617-134">繰り返し単一値パラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-134">Repeating single value parameter</span></span>

<span data-ttu-id="f3617-135">繰り返し単一値パラメーターを使用すると、複数の単一の値を渡す可能性があります。</span><span class="sxs-lookup"><span data-stu-id="f3617-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="f3617-136">たとえば、ユーザーは ADD(1,B2,3) と入力できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="f3617-137">次のサンプルは、1 つの値パラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f3617-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="f3617-138">単一の範囲パラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-138">Single range parameter</span></span>

<span data-ttu-id="f3617-139">1 つの範囲パラメーターは技術的には繰り返しパラメーターではなく、宣言が繰り返しパラメーターと非常に似ているため、ここに含まれています。</span><span class="sxs-lookup"><span data-stu-id="f3617-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="f3617-140">ユーザーには ADD(A2:B3) と表示され、Excel から 1 つの範囲が渡されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="f3617-141">次のサンプルは、1 つの範囲パラメーターを宣言する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f3617-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="f3617-142">繰り返し範囲パラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-142">Repeating range parameter</span></span>

<span data-ttu-id="f3617-143">繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="f3617-144">たとえば、ユーザーは ADD(5,B2,C3,8,E5:E8) と入力できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="f3617-145">繰り返し範囲は、通常、3 次元マトリックス `number[][][]` である型で指定されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="f3617-146">サンプルについては、繰り返しパラメーター (#repeating-parameters) の一覧にあるメイン サンプルを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f3617-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="f3617-147">繰り返しパラメーターの宣言</span><span class="sxs-lookup"><span data-stu-id="f3617-147">Declaring repeating parameters</span></span>
<span data-ttu-id="f3617-148">Typescript で、パラメーターが多次元パラメーターかどうかを示します。</span><span class="sxs-lookup"><span data-stu-id="f3617-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="f3617-149">たとえば  `ADD(values: number[])` 、1 次元配列を示し `ADD(values:number[][])` 、2 次元配列を示す場合などです。</span><span class="sxs-lookup"><span data-stu-id="f3617-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="f3617-150">JavaScript では、1 次元配列、2 次元配列、およびより多くの次元 `@param values {number[]}` `@param <name> {number[][]}` に使用します。</span><span class="sxs-lookup"><span data-stu-id="f3617-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="f3617-151">手書き JSON の場合は、JSON ファイルでパラメーターが指定されているのを確認し、パラメーターにマーク `"repeating": true` が付けられているか確認します `"dimensionality": matrix` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="f3617-152">呼び出しパラメーター</span><span class="sxs-lookup"><span data-stu-id="f3617-152">Invocation parameter</span></span>

<span data-ttu-id="f3617-153">すべてのカスタム関数は、引数が明示的に宣言されていない場合でも、最後の入力パラメーターとして自動的 `invocation` に渡されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="f3617-154">この `invocation` パラメーターは、呼び出しオブジェクト [に対応](/javascript/api/custom-functions-runtime/customfunctions.invocation) します。</span><span class="sxs-lookup"><span data-stu-id="f3617-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="f3617-155">オブジェクトを使用して、カスタム関数を呼び出したセルのアドレスなど、追加 `Invocation` のコンテキストを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f3617-156">オブジェクトにアクセス `Invocation` するには、カスタム関数の `invocation` 最後のパラメーターとして宣言する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f3617-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="f3617-157">この `invocation` パラメーターは、Excel のユーザーのカスタム関数引数として表示されません。</span><span class="sxs-lookup"><span data-stu-id="f3617-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="f3617-158">次のサンプルは、パラメーターを使用して、カスタム関数を呼び出したセルのアドレス `invocation` を返す方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f3617-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f3617-159">このサンプルでは、 [オブジェクトの address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) プロパティを使用 `Invocation` します。</span><span class="sxs-lookup"><span data-stu-id="f3617-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="f3617-160">オブジェクトにアクセス `Invocation` するには、まず `CustomFunctions.Invocation` JSDoc でパラメーターとして宣言します。</span><span class="sxs-lookup"><span data-stu-id="f3617-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="f3617-161">次に、 `@requiresAddress` オブジェクトのプロパティにアクセスするために JSDoc `address` で宣言 `Invocation` します。</span><span class="sxs-lookup"><span data-stu-id="f3617-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="f3617-162">最後に、関数内でプロパティを取得して返 `address` します。</span><span class="sxs-lookup"><span data-stu-id="f3617-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<span data-ttu-id="f3617-163">Excel では、オブジェクトのプロパティを呼び出すカスタム関数は、関数を呼び出したセルの形式に従って絶対アドレス `address` `Invocation` `SheetName!RelativeCellAddress` を返します。</span><span class="sxs-lookup"><span data-stu-id="f3617-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f3617-164">たとえば、入力パラメーターがセル F6 の **[価格** ] というシートにある場合、返されるパラメーターのアドレス値は次のようになります `Prices!F6` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="f3617-165">この `invocation` パラメーターは、Excel に情報を送信するためにも使用できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="f3617-166">詳 [しくは、「ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function) 」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f3617-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="f3617-167">パラメーターのアドレスを検出する</span><span class="sxs-lookup"><span data-stu-id="f3617-167">Detect the address of a parameter</span></span>

<span data-ttu-id="f3617-168">呼び出しパラメーター [と組](#invocation-parameter)み合わせて、 [呼](/javascript/api/custom-functions-runtime/customfunctions.invocation) び出しオブジェクトを使用して、カスタム関数入力パラメーターのアドレスを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="f3617-169">呼び出されると、 [オブジェクトの parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) プロパティを使用すると、関数は、すべての入力パラメーター `Invocation` のアドレスを返すことができます。</span><span class="sxs-lookup"><span data-stu-id="f3617-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="f3617-170">これは、入力データ型が異なる可能性があるシナリオで役立ちます。</span><span class="sxs-lookup"><span data-stu-id="f3617-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="f3617-171">入力パラメーターのアドレスを使用して、入力値の数値形式を確認できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="f3617-172">必要に応じて、数値の書式を入力前に調整できます。</span><span class="sxs-lookup"><span data-stu-id="f3617-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="f3617-173">入力パラメーターのアドレスを使用して、入力値に後続の計算に関連する可能性のある関連プロパティが含されているかどうかを検出することもできます。</span><span class="sxs-lookup"><span data-stu-id="f3617-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!NOTE]
> <span data-ttu-id="f3617-174">Yo Office ジェネレーターの代わりに、手動で作成された [JSON](custom-functions-json.md) メタデータを使用してパラメーター アドレスを返す場合、オブジェクトにはプロパティが設定され、オブジェクトにはプロパティが設定されている必要 `options` `requiresParameterAddresses` `true` `result` `dimensionality` があります `matrix` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-174">If you're working with [manually-created JSON metadata](custom-functions-json.md) to return parameter addresses instead of the Yo Office generator, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="f3617-175">次のカスタム関数は、3 つの入力パラメーターを受け取り、各パラメーターのオブジェクトのプロパティを取得し、アドレス `parameterAddresses` `Invocation` を返します。</span><span class="sxs-lookup"><span data-stu-id="f3617-175">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<span data-ttu-id="f3617-176">プロパティを呼び出すカスタム関数が実行されると、関数を呼び出したセルの形式に従ってパラメーター アドレス `parameterAddresses` `SheetName!RelativeCellAddress` が返されます。</span><span class="sxs-lookup"><span data-stu-id="f3617-176">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="f3617-177">たとえば、入力パラメーターがセル D8 の **Costs** というシートにある場合、返されるパラメーター のアドレス値は次のようになります `Costs!D8` 。</span><span class="sxs-lookup"><span data-stu-id="f3617-177">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="f3617-178">カスタム関数に複数のパラメーターが含まれていますが、複数のパラメーター アドレスが返された場合、返されるアドレスは複数のセルにこぼれ、関数を呼び出したセルから垂直方向に降順になります。</span><span class="sxs-lookup"><span data-stu-id="f3617-178">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="f3617-179">次の手順</span><span class="sxs-lookup"><span data-stu-id="f3617-179">Next steps</span></span>

<span data-ttu-id="f3617-180">カスタム関数で揮発性値 [を使用する方法について説明します](custom-functions-volatile.md)。</span><span class="sxs-lookup"><span data-stu-id="f3617-180">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f3617-181">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="f3617-181">See also</span></span>

* [<span data-ttu-id="f3617-182">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="f3617-182">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="f3617-183">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="f3617-183">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="f3617-184">カスタム関数の JSON メタデータを手動で作成する</span><span class="sxs-lookup"><span data-stu-id="f3617-184">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f3617-185">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="f3617-185">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f3617-186">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="f3617-186">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
