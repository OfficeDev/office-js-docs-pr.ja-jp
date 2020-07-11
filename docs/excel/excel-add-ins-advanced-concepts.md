---
title: Excel JavaScript API を使用した高度なプログラミングの概念
description: Excel アドインが Office JavaScript API オブジェクト モデルを使用して Excel 内のオブジェクトを操作する方法について説明します。
ms.date: 07/01/2020
localization_priority: Priority
ms.openlocfilehash: 81602f48231f20b50a454134bc789dfdee2bbc12
ms.sourcegitcommit: 4f2f1c0a8ee777a43bb28efa226684261f4c4b9f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/08/2020
ms.locfileid: "45081397"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="23fc6-103">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="23fc6-103">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="23fc6-104">この記事では、「[Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」の情報を基にして、より高度な概念をいくつか説明します。これらは Excel 2016 以降の複雑なアドインを構築するために不可欠です。</span><span class="sxs-lookup"><span data-stu-id="23fc6-104">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="23fc6-105">Excel 用の Office.js API</span><span class="sxs-lookup"><span data-stu-id="23fc6-105">Office.js APIs for Excel</span></span>

<span data-ttu-id="23fc6-106">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む Office JavaScript API を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="23fc6-106">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="23fc6-107">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="23fc6-107">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="23fc6-108">**共通 API**: Office 2013 で導入された[共通 API](/javascript/api/office) を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-108">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="23fc6-109">Excel 2016 以降を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。</span><span class="sxs-lookup"><span data-stu-id="23fc6-109">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="23fc6-110">例:</span><span class="sxs-lookup"><span data-stu-id="23fc6-110">For example:</span></span>

- <span data-ttu-id="23fc6-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span><span class="sxs-lookup"><span data-stu-id="23fc6-111">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="23fc6-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span><span class="sxs-lookup"><span data-stu-id="23fc6-112">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="23fc6-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span><span class="sxs-lookup"><span data-stu-id="23fc6-113">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>

- <span data-ttu-id="23fc6-114">[Document](/javascript/api/office/office.document): `Document` オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-114">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="23fc6-115">次の図は、Excel JavaScript API または共通 API を使用するタイミングを示しています。</span><span class="sxs-lookup"><span data-stu-id="23fc6-115">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Excel JS API と共通 API の違いを示す画像](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a><span data-ttu-id="23fc6-117">要件セット</span><span class="sxs-lookup"><span data-stu-id="23fc6-117">Requirement sets</span></span>

<span data-ttu-id="23fc6-118">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="23fc6-118">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="23fc6-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span><span class="sxs-lookup"><span data-stu-id="23fc6-119">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="23fc6-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="23fc6-120">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="23fc6-121">実行時に要件セットのサポートを確認する</span><span class="sxs-lookup"><span data-stu-id="23fc6-121">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="23fc6-122">次のコード サンプルは、アドインが実行されているホスト アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="23fc6-122">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="23fc6-123">マニフェストで要件セットのサポートを定義する</span><span class="sxs-lookup"><span data-stu-id="23fc6-123">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="23fc6-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span><span class="sxs-lookup"><span data-stu-id="23fc6-124">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="23fc6-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span><span class="sxs-lookup"><span data-stu-id="23fc6-125">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="23fc6-126">次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office ホスト アプリケーションのすべて読み込まれる必要があることを指定する、アドインのマニフェストの `Requirements`Requirements 要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="23fc6-126">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="23fc6-127">Excel on the web、Windows、iPad などの Office ホストのプラットフォームすべてでアドインを使用できるようにするには、マニフェストで要件セットのサポートを定義するのではなく、実行時に要件のサポートを確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="23fc6-127">To make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="23fc6-128">Office.js 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="23fc6-128">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="23fc6-129">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="23fc6-129">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="23fc6-130">オブジェクトのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="23fc6-130">Loading the properties of an object</span></span>

<span data-ttu-id="23fc6-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span><span class="sxs-lookup"><span data-stu-id="23fc6-131">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="23fc6-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span><span class="sxs-lookup"><span data-stu-id="23fc6-132">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span>

### <a name="method-details"></a><span data-ttu-id="23fc6-133">メソッドの詳細</span><span class="sxs-lookup"><span data-stu-id="23fc6-133">Method details</span></span>

#### `load(propertyNames?: string | string[])`

<span data-ttu-id="23fc6-134">オブジェクトの指定されたプロパティを読み込むコマンドを待ち行列に入れます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-134">Queues up a command to load the specified properties of the object.</span></span> <span data-ttu-id="23fc6-135">プロパティを読み取る前に、`context.sync()` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="23fc6-135">You must call `context.sync()` before reading the properties.</span></span>

#### <a name="syntax"></a><span data-ttu-id="23fc6-136">構文</span><span class="sxs-lookup"><span data-stu-id="23fc6-136">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="23fc6-137">パラメーター</span><span class="sxs-lookup"><span data-stu-id="23fc6-137">Parameters</span></span>

|<span data-ttu-id="23fc6-138">**パラメーター**</span><span class="sxs-lookup"><span data-stu-id="23fc6-138">**Parameter**</span></span>|<span data-ttu-id="23fc6-139">**型**</span><span class="sxs-lookup"><span data-stu-id="23fc6-139">**Type**</span></span>|<span data-ttu-id="23fc6-140">**説明**</span><span class="sxs-lookup"><span data-stu-id="23fc6-140">**Description**</span></span>|
|:------------|:-------|:----------|
|`propertyNames`|<span data-ttu-id="23fc6-141">object</span><span class="sxs-lookup"><span data-stu-id="23fc6-141">object</span></span>|<span data-ttu-id="23fc6-142">オプション。</span><span class="sxs-lookup"><span data-stu-id="23fc6-142">Optional.</span></span> <span data-ttu-id="23fc6-143">プロパティ名を、コンマで区切られた文字列または 1 つの配列として指定します。</span><span class="sxs-lookup"><span data-stu-id="23fc6-143">Accepts property names as comma-delimited string or an array.</span></span>|

#### <a name="returns"></a><span data-ttu-id="23fc6-144">戻り値</span><span class="sxs-lookup"><span data-stu-id="23fc6-144">Returns</span></span>

<span data-ttu-id="23fc6-145">void</span><span class="sxs-lookup"><span data-stu-id="23fc6-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="23fc6-146">例</span><span class="sxs-lookup"><span data-stu-id="23fc6-146">Example</span></span>

<span data-ttu-id="23fc6-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span><span class="sxs-lookup"><span data-stu-id="23fc6-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="23fc6-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span><span class="sxs-lookup"><span data-stu-id="23fc6-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="23fc6-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span><span class="sxs-lookup"><span data-stu-id="23fc6-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a><span data-ttu-id="23fc6-150">オプションのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="23fc6-150">Load option properties</span></span>

<span data-ttu-id="23fc6-151">`load()` メソッドを呼び出すときに、コンマで区切られた文字列または配列を渡す代わりに、次のプロパティを含むオブジェクトを渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span>

|<span data-ttu-id="23fc6-152">**プロパティ**</span><span class="sxs-lookup"><span data-stu-id="23fc6-152">**Property**</span></span>|<span data-ttu-id="23fc6-153">**型**</span><span class="sxs-lookup"><span data-stu-id="23fc6-153">**Type**</span></span>|<span data-ttu-id="23fc6-154">**説明**</span><span class="sxs-lookup"><span data-stu-id="23fc6-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="23fc6-155">object</span><span class="sxs-lookup"><span data-stu-id="23fc6-155">object</span></span>|<span data-ttu-id="23fc6-156">Contains a comma-delimited list or an array of scalar property names.</span><span class="sxs-lookup"><span data-stu-id="23fc6-156">Contains a comma-delimited list or an array of scalar property names.</span></span> <span data-ttu-id="23fc6-157">Optional.</span><span class="sxs-lookup"><span data-stu-id="23fc6-157">Optional.</span></span>|
|`expand`|<span data-ttu-id="23fc6-158">object</span><span class="sxs-lookup"><span data-stu-id="23fc6-158">object</span></span>|<span data-ttu-id="23fc6-159">Contains a comma-delimited list or an array of navigational property names.</span><span class="sxs-lookup"><span data-stu-id="23fc6-159">Contains a comma-delimited list or an array of navigational property names.</span></span> <span data-ttu-id="23fc6-160">Optional.</span><span class="sxs-lookup"><span data-stu-id="23fc6-160">Optional.</span></span>|
|`top`|<span data-ttu-id="23fc6-161">int</span><span class="sxs-lookup"><span data-stu-id="23fc6-161">int</span></span>| <span data-ttu-id="23fc6-162">Specifies the maximum number of collection items that can be included in the result.</span><span class="sxs-lookup"><span data-stu-id="23fc6-162">Specifies the maximum number of collection items that can be included in the result.</span></span> <span data-ttu-id="23fc6-163">Optional.</span><span class="sxs-lookup"><span data-stu-id="23fc6-163">Optional.</span></span> <span data-ttu-id="23fc6-164">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="23fc6-164">You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="23fc6-165">int</span><span class="sxs-lookup"><span data-stu-id="23fc6-165">int</span></span>|<span data-ttu-id="23fc6-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span><span class="sxs-lookup"><span data-stu-id="23fc6-166">Specify the number of items in the collection that are to be skipped and not included in the result.</span></span> <span data-ttu-id="23fc6-167">If `top` is specified, the result set will start after skipping the specified number of items.</span><span class="sxs-lookup"><span data-stu-id="23fc6-167">If `top` is specified, the result set will start after skipping the specified number of items.</span></span> <span data-ttu-id="23fc6-168">Optional.</span><span class="sxs-lookup"><span data-stu-id="23fc6-168">Optional.</span></span> <span data-ttu-id="23fc6-169">You can only use this option when you use the object notation option.</span><span class="sxs-lookup"><span data-stu-id="23fc6-169">You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="23fc6-170">次のコードサンプルは、`name` プロパティと `address`コレクション内の各ワークシートの使用範囲を選択して、ワークシートコレクションを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-170">The following code sample loads a worksheet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="23fc6-171">また、コレクションの上位 5 つのワークシートのみを読み込むように指定しています。</span><span class="sxs-lookup"><span data-stu-id="23fc6-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="23fc6-172">`top: 10` と `skip: 5` を属性値として指定することで、次の 5 つのワークシートのセットを処理できます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span>

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

### <a name="calling-load-without-parameters"></a><span data-ttu-id="23fc6-173">パラメーターを使用せずに `load` を呼び出す</span><span class="sxs-lookup"><span data-stu-id="23fc6-173">Calling `load` without parameters</span></span>

<span data-ttu-id="23fc6-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span><span class="sxs-lookup"><span data-stu-id="23fc6-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="23fc6-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span><span class="sxs-lookup"><span data-stu-id="23fc6-175">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="23fc6-176">パラメーターのない `load` ステートメントで返されるデータの量は、サービスのサイズ制限を超える場合があります。</span><span class="sxs-lookup"><span data-stu-id="23fc6-176">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="23fc6-177">古いアドインのリスクを軽減するために、明示的に要求しない限り `load` によって返されないプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="23fc6-177">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="23fc6-178">次のプロパティは、そのような負荷操作から除外されます。</span><span class="sxs-lookup"><span data-stu-id="23fc6-178">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="23fc6-179">スカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="23fc6-179">Scalar and navigation properties</span></span>

<span data-ttu-id="23fc6-180">プロパティには、**スカラー**と**ナビゲーション**という 2 つのカテゴリがあります。</span><span class="sxs-lookup"><span data-stu-id="23fc6-180">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="23fc6-181">スカラー プロパティは、文字列、整数、JSON 構造体などの割り当て可能な型です。</span><span class="sxs-lookup"><span data-stu-id="23fc6-181">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="23fc6-182">ナビゲーション プロパティは、プロパティを直接割り当てるのではなく、読み取り専用のオブジェクトと、そのフィールドが割り当てられているオブジェクトのコレクションです。</span><span class="sxs-lookup"><span data-stu-id="23fc6-182">Navigation properties are readonly objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="23fc6-183">たとえば、[ワークシート](/javascript/api/excel/excel.worksheet) オブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はナビゲーション プロパティです。</span><span class="sxs-lookup"><span data-stu-id="23fc6-183">For example, `name` and `position` members on the [Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span> <span data-ttu-id="23fc6-184">[DataValidation](/javascript/api/excel/excel.datavalidation) オブジェクトの `prompt` は、サブプロパティ (`dv.prompt.title = "MyPrompt" // will not set the title`) を設定するのではなく、JSON オブジェクト (`dv.prompt = { title: "MyPrompt"}`) を使用して設定する必要があるスカラー プロパティの例です。</span><span class="sxs-lookup"><span data-stu-id="23fc6-184">`prompt` on the [DataValidation](/javascript/api/excel/excel.datavalidation) object is an example of a scalar property that must be set using a JSON object (`dv.prompt = { title: "MyPrompt"}`), instead of setting the sub-properties (`dv.prompt.title = "MyPrompt" // will not set the title`).</span></span>

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="23fc6-185">`object.load()` を使用したスカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="23fc6-185">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="23fc6-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span><span class="sxs-lookup"><span data-stu-id="23fc6-186">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="23fc6-187">Additionally, navigation properties cannot be loaded directly.</span><span class="sxs-lookup"><span data-stu-id="23fc6-187">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="23fc6-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span><span class="sxs-lookup"><span data-stu-id="23fc6-188">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="23fc6-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span><span class="sxs-lookup"><span data-stu-id="23fc6-189">For example, to load the font name for a range, you must specify the `format` and `font` navigation properties as the path to the `name` property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="23fc6-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span><span class="sxs-lookup"><span data-stu-id="23fc6-190">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="23fc6-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="23fc6-191">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="23fc6-192">You do not need to load the property before you set it.</span><span class="sxs-lookup"><span data-stu-id="23fc6-192">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="23fc6-193">オブジェクトのプロパティを設定する</span><span class="sxs-lookup"><span data-stu-id="23fc6-193">Setting properties of an object</span></span>

<span data-ttu-id="23fc6-194">Setting properties on an object with nested navigation properties can be cumbersome.</span><span class="sxs-lookup"><span data-stu-id="23fc6-194">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="23fc6-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="23fc6-195">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="23fc6-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span><span class="sxs-lookup"><span data-stu-id="23fc6-196">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="23fc6-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="23fc6-197">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="23fc6-198">The common (shared) APIs do not support this method.</span><span class="sxs-lookup"><span data-stu-id="23fc6-198">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="23fc6-199">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="23fc6-199">set (properties: object, options: object)</span></span>

<span data-ttu-id="23fc6-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span><span class="sxs-lookup"><span data-stu-id="23fc6-200">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object.</span></span> <span data-ttu-id="23fc6-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span><span class="sxs-lookup"><span data-stu-id="23fc6-201">If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="23fc6-202">構文</span><span class="sxs-lookup"><span data-stu-id="23fc6-202">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="23fc6-203">パラメーター</span><span class="sxs-lookup"><span data-stu-id="23fc6-203">Parameters</span></span>

|<span data-ttu-id="23fc6-204">**パラメーター**</span><span class="sxs-lookup"><span data-stu-id="23fc6-204">**Parameter**</span></span>|<span data-ttu-id="23fc6-205">**型**</span><span class="sxs-lookup"><span data-stu-id="23fc6-205">**Type**</span></span>|<span data-ttu-id="23fc6-206">**説明**</span><span class="sxs-lookup"><span data-stu-id="23fc6-206">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="23fc6-207">object</span><span class="sxs-lookup"><span data-stu-id="23fc6-207">object</span></span>|<span data-ttu-id="23fc6-208">メソッドが呼び出されるオブジェクトの同じ Office.js 型のオブジェクト、またはメソッドが呼び出されるオブジェクトの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトのいずれかです。</span><span class="sxs-lookup"><span data-stu-id="23fc6-208">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="23fc6-209">object</span><span class="sxs-lookup"><span data-stu-id="23fc6-209">object</span></span>|<span data-ttu-id="23fc6-210">Optional.</span><span class="sxs-lookup"><span data-stu-id="23fc6-210">Optional.</span></span> <span data-ttu-id="23fc6-211">Can only be passed when the first parameter is a JavaScript object.</span><span class="sxs-lookup"><span data-stu-id="23fc6-211">Can only be passed when the first parameter is a JavaScript object.</span></span> <span data-ttu-id="23fc6-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span><span class="sxs-lookup"><span data-stu-id="23fc6-212">The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="23fc6-213">戻り値</span><span class="sxs-lookup"><span data-stu-id="23fc6-213">Returns</span></span>

<span data-ttu-id="23fc6-214">void</span><span class="sxs-lookup"><span data-stu-id="23fc6-214">void</span></span>

#### <a name="example"></a><span data-ttu-id="23fc6-215">例</span><span class="sxs-lookup"><span data-stu-id="23fc6-215">Example</span></span>

<span data-ttu-id="23fc6-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span><span class="sxs-lookup"><span data-stu-id="23fc6-216">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object.</span></span> <span data-ttu-id="23fc6-217">This example assumes that there is data in range **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="23fc6-217">This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a><span data-ttu-id="23fc6-218">&#42;OrNullObject メソッド</span><span class="sxs-lookup"><span data-stu-id="23fc6-218">&#42;OrNullObject methods</span></span>

<span data-ttu-id="23fc6-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span><span class="sxs-lookup"><span data-stu-id="23fc6-219">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="23fc6-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span><span class="sxs-lookup"><span data-stu-id="23fc6-220">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="23fc6-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span><span class="sxs-lookup"><span data-stu-id="23fc6-221">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="23fc6-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span><span class="sxs-lookup"><span data-stu-id="23fc6-222">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="23fc6-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span><span class="sxs-lookup"><span data-stu-id="23fc6-223">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="23fc6-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span><span class="sxs-lookup"><span data-stu-id="23fc6-224">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="23fc6-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span><span class="sxs-lookup"><span data-stu-id="23fc6-225">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="23fc6-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span><span class="sxs-lookup"><span data-stu-id="23fc6-226">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="23fc6-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span><span class="sxs-lookup"><span data-stu-id="23fc6-227">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a><span data-ttu-id="23fc6-228">関連項目</span><span class="sxs-lookup"><span data-stu-id="23fc6-228">See also</span></span>

* [<span data-ttu-id="23fc6-229">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="23fc6-229">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="23fc6-230">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="23fc6-230">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="23fc6-231">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="23fc6-231">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="23fc6-232">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="23fc6-232">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
