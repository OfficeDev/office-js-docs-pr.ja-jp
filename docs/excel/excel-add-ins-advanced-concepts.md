---
title: Excel JavaScript API を使用した高度なプログラミングの概念
description: ''
ms.date: 10/03/2018
localization_priority: Priority
ms.openlocfilehash: 1e623e87cbda8b8aeb6e51104bec818d62bf489e
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388483"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="7ec69-102">Excel JavaScript API を使用した高度なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="7ec69-102">Advanced programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="7ec69-103">この記事では、「[Excel JavaScript API を使用した基本的なプログラミングの概念](excel-add-ins-core-concepts.md)」の情報を基にして、より高度な概念をいくつか説明します。これらは Excel 2016 以降の複雑なアドインを構築するために不可欠です。</span><span class="sxs-lookup"><span data-stu-id="7ec69-103">This article builds upon the information in [Fundamental programming concepts with the Excel JavaScript API](excel-add-ins-core-concepts.md) to describe some of the more advanced concepts that are essential to building complex add-ins for Excel 2016 or later.</span></span>

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="7ec69-104">Excel 用の Office.js API</span><span class="sxs-lookup"><span data-stu-id="7ec69-104">Office.js APIs for Excel</span></span>

<span data-ttu-id="7ec69-105">Excel アドインは、次の 2 つの JavaScript オブジェクト モデルを含む JavaScript API for Office を使用して、Excel のオブジェクトを操作します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-105">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="7ec69-106">**Excel JavaScript API**:Office 2016 で導入された [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) には、ワークシート、範囲、表、グラフなどへのアクセスに使用できる、厳密に型指定されたオブジェクトが用意されています。</span><span class="sxs-lookup"><span data-stu-id="7ec69-106">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="7ec69-107">**共通 API**: Office 2013 で導入された[共通 API](../reference/javascript-api-for-office.md) を使用すると、Word、Excel、PowerPoint など複数の種類のホスト アプリケーションに共通する UI、ダイアログ、クライアント設定などの機能にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-107">**Common APIs**: Introduced with Office 2013, the [Common API](../reference/javascript-api-for-office.md) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of host applications such as Word, Excel, and PowerPoint.</span></span>

<span data-ttu-id="7ec69-108">Excel 2016 以降を対象にしたアドインでは、機能の大部分を Excel JavaScript API を使用して開発する可能性がありますが、共通 API のオブジェクトも使用します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-108">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="7ec69-109">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-109">For example:</span></span>

- <span data-ttu-id="7ec69-110">[Context](https://docs.microsoft.com/javascript/api/office/office.context): **Context** オブジェクトは、アドインのランタイム環境を表し、API の主要なオブジェクトへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-110">[Context](https://docs.microsoft.com/javascript/api/office/office.context): The **Context** object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="7ec69-111">これは `contentLanguage` や `officeTheme` などのブック構成の詳細で構成され、`host` や `platform` などのアドインのランタイム環境に関する情報も提供します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-111">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="7ec69-112">さらに、`requirements.isSetSupported()` メソッドも提供されます。これを使用すると、指定した要件セットが、アドインが実行されている Excel アプリケーションでサポートされているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-112">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span> 

- <span data-ttu-id="7ec69-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document):**Document** オブジェクトは `getFileAsync()` メソッドを提供します。これを使用すると、アドインが実行されている Excel ファイルをダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-113">[Document](https://docs.microsoft.com/javascript/api/office/office.document): The **Document** object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span> 

## <a name="requirement-sets"></a><span data-ttu-id="7ec69-114">要件セット</span><span class="sxs-lookup"><span data-stu-id="7ec69-114">Requirement sets</span></span>

<span data-ttu-id="7ec69-115">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="7ec69-115">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="7ec69-116">Office アドインはランタイム チェックを実行できます。または、マニフェストで指定されている要件セットを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-116">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office host supports the APIs that the add-in needs.</span></span> <span data-ttu-id="7ec69-117">サポートされている各プラットフォームで使用できる特定の要件セットを確認するには、「[Excel JavaScript API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7ec69-117">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="7ec69-118">実行時に要件セットのサポートを確認する</span><span class="sxs-lookup"><span data-stu-id="7ec69-118">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="7ec69-119">次のコード サンプルは、アドインが実行されているホスト アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7ec69-119">The following code sample shows how to determine whether the host application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="7ec69-120">マニフェストで要件セットのサポートを定義する</span><span class="sxs-lookup"><span data-stu-id="7ec69-120">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="7ec69-121">アドインのマニフェストで [Requirements 要素](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements) を使用して、アドインをアクティブにするために必要な最小要件セットや API メソッド (またはその両方) を指定できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-121">You can use the [Requirements element](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="7ec69-122">Office ホストまたはプラットフォームが、マニフェストの **Requirements** 要素で指定した要件セットまたは API メソッドをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、**[個人用アドイン]** に表示されるアドインの一覧にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="7ec69-122">If the Office host or platform doesn't support the requirement sets or API methods that are specified in the **Requirements** element of the manifest, the add-in won't run in that host or platform, and won't display in the list of add-ins that are shown in **My Add-ins**.</span></span> 

<span data-ttu-id="7ec69-123">次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office ホスト アプリケーションのすべて読み込まれる必要があることを指定する、アドインのマニフェストの **Requirements** 要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="7ec69-123">The following code sample shows the **Requirements** element in an add-in manifest which specifies that the add-in should load in all Office host applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="7ec69-124">Excel for Windows、Excel Online、Excel for iPad などの Office ホストのプラットフォームすべてでアドインを使用できるようにするには、マニフェストで要件セットのサポートを定義するのではなく、実行時に要件のサポートを確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="7ec69-124">To make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="7ec69-125">Office.js 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="7ec69-125">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="7ec69-126">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="7ec69-126">For information about Common API requirement sets, see [Office Common API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>

## <a name="loading-the-properties-of-an-object"></a><span data-ttu-id="7ec69-127">オブジェクトのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="7ec69-127">Loading the properties of an object</span></span>

<span data-ttu-id="7ec69-128">Excel JavaScript オブジェクトで `load()` メソッドを呼び出すと、API は `sync()` メソッドの実行時にオブジェクトを JavaScript メモリに読み込むように指示されます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-128">Calling the `load()` method on an Excel JavaScript object instructs the API to load the object into JavaScript memory when the `sync()` method runs.</span></span> <span data-ttu-id="7ec69-129">`load()` メソッドには、読み込むプロパティのコンマで区切られた名前を含む文字列や、読み込むプロパティを指定するオブジェクト、改ページのオプションなどを指定できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-129">The `load()` method accepts a string that contains comma-delimited names of properties to load or an object that specifies properties to load, pagination options, etc.</span></span> 

> [!NOTE]
> <span data-ttu-id="7ec69-130">パラメーターを指定せずにオブジェクト (またはコレクション) の `load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティ (またはコレクション内のすべてのオブジェクトのすべてのスカラー プロパティ) が読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-130">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object (or all scalar properties of all objects in the collection) will be loaded.</span></span> <span data-ttu-id="7ec69-131">Excel ホスト アプリケーションとアドイン間のデータ転送量を減らすには、読み込むプロパティを明示的に指定しないで `load()` メソッドを呼び出さないようにします。</span><span class="sxs-lookup"><span data-stu-id="7ec69-131">To reduce the amount of data transfer between the Excel host application and the add-in, you should avoid calling the `load()` method without explicitly specifying which properties to load.</span></span>

### <a name="method-details"></a><span data-ttu-id="7ec69-132">メソッドの詳細</span><span class="sxs-lookup"><span data-stu-id="7ec69-132">Method details</span></span>

#### <a name="loadparam-object"></a><span data-ttu-id="7ec69-133">load(param: object)</span><span class="sxs-lookup"><span data-stu-id="7ec69-133">load(param: object)</span></span>

<span data-ttu-id="7ec69-134">JavaScript レイヤーで作成されたプロキシ オブジェクトに、パラメーターで指定されているプロパティとオブジェクトの値を設定します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-134">Fills the proxy object created in JavaScript layer with property and object values specified by the parameters.</span></span>

#### <a name="syntax"></a><span data-ttu-id="7ec69-135">構文</span><span class="sxs-lookup"><span data-stu-id="7ec69-135">Syntax</span></span>

```js
object.load(param);
```

#### <a name="parameters"></a><span data-ttu-id="7ec69-136">パラメーター</span><span class="sxs-lookup"><span data-stu-id="7ec69-136">Parameters</span></span>

|<span data-ttu-id="7ec69-137">**パラメーター**</span><span class="sxs-lookup"><span data-stu-id="7ec69-137">**Parameter**</span></span>|<span data-ttu-id="7ec69-138">**型**</span><span class="sxs-lookup"><span data-stu-id="7ec69-138">**Type**</span></span>|<span data-ttu-id="7ec69-139">**説明**</span><span class="sxs-lookup"><span data-stu-id="7ec69-139">**Description**</span></span>|
|:------------|:-------|:----------|
|`param`|<span data-ttu-id="7ec69-140">object</span><span class="sxs-lookup"><span data-stu-id="7ec69-140">object</span></span>|<span data-ttu-id="7ec69-141">省略可能。</span><span class="sxs-lookup"><span data-stu-id="7ec69-141">Optional.</span></span> <span data-ttu-id="7ec69-142">パラメーターとリレーションシップ名を、コンマで区切られた文字列または 1 つの配列として指定します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-142">Accepts parameter and relationship names as comma-delimited string or an array.</span></span> <span data-ttu-id="7ec69-143">オブジェクトを渡して、選択プロパティとナビゲーション プロパティを設定することもできます (次の例を参照)。</span><span class="sxs-lookup"><span data-stu-id="7ec69-143">An object can also be passed to set the selection and navigation properties (as shown in the example below).</span></span>|

#### <a name="returns"></a><span data-ttu-id="7ec69-144">戻り値</span><span class="sxs-lookup"><span data-stu-id="7ec69-144">Returns</span></span>

<span data-ttu-id="7ec69-145">void</span><span class="sxs-lookup"><span data-stu-id="7ec69-145">void</span></span>

#### <a name="example"></a><span data-ttu-id="7ec69-146">例</span><span class="sxs-lookup"><span data-stu-id="7ec69-146">Example</span></span>

<span data-ttu-id="7ec69-147">次のコード サンプルでは、別の範囲のプロパティをコピーして 1 つの Excel 範囲のプロパティを設定します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-147">The following code sample sets the properties of one Excel range by copying the properties of another range.</span></span> <span data-ttu-id="7ec69-148">プロパティ値にアクセスして対象範囲に書き込む前に、ソース オブジェクトを最初に読み込む必要があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="7ec69-148">Note that the source object must be loaded first, before its property values can be accessed and written to the target range.</span></span> <span data-ttu-id="7ec69-149">この例では、2 つの範囲 (**B2:E2** および **B7:E7**) のデータがあり、2 つの範囲の書式設定が最初は異なっていると仮定します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-149">This example assumes that there is data the two ranges (**B2:E2** and **B7:E7**) and that the two ranges are initially formatted differently.</span></span>

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

### <a name="load-option-properties"></a><span data-ttu-id="7ec69-150">オプションのプロパティを読み込む</span><span class="sxs-lookup"><span data-stu-id="7ec69-150">Load option properties</span></span>

<span data-ttu-id="7ec69-151">`load()` メソッドを呼び出すときに、コンマで区切られた文字列または配列を渡す代わりに、次のプロパティを含むオブジェクトを渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-151">As an alternative to passing a comma-delimited string or array when you call the `load()` method, you can pass an object that contains the following properties.</span></span> 

|<span data-ttu-id="7ec69-152">**プロパティ**</span><span class="sxs-lookup"><span data-stu-id="7ec69-152">**Property**</span></span>|<span data-ttu-id="7ec69-153">**型**</span><span class="sxs-lookup"><span data-stu-id="7ec69-153">**Type**</span></span>|<span data-ttu-id="7ec69-154">**説明**</span><span class="sxs-lookup"><span data-stu-id="7ec69-154">**Description**</span></span>|
|:-----------|:-------|:----------|
|`select`|<span data-ttu-id="7ec69-155">object</span><span class="sxs-lookup"><span data-stu-id="7ec69-155">object</span></span>|<span data-ttu-id="7ec69-p109">パラメーター/リレーションシップの名前のコンマ区切りリストまたは配列が含まれます。省略可能。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p109">Contains a comma-delimited list or an array of parameter/relationship names. Optional.</span></span>|
|`expand`|<span data-ttu-id="7ec69-158">object</span><span class="sxs-lookup"><span data-stu-id="7ec69-158">object</span></span>|<span data-ttu-id="7ec69-p110">リレーションシップ名のコンマ区切りリストまたは配列が含まれています。省略可能。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p110">Contains a comma-delimited list or an array of relationship names. Optional.</span></span>|
|`top`|<span data-ttu-id="7ec69-161">int</span><span class="sxs-lookup"><span data-stu-id="7ec69-161">int</span></span>| <span data-ttu-id="7ec69-p111">結果に含めることができるコレクション項目の最大数を指定します。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p111">Specifies the maximum number of collection items that can be included in the result. Optional. You can only use this option when you use the object notation option.</span></span>|
|`skip`|<span data-ttu-id="7ec69-165">int</span><span class="sxs-lookup"><span data-stu-id="7ec69-165">int</span></span>|<span data-ttu-id="7ec69-p112">スキップされて結果に組み込まれないコレクション内の項目の数を指定します。`top` が指定されている場合は、指定された数の項目がスキップされた後で結果セットが開始されます。省略可能。このオプションは、オブジェクト表記オプションを使用する場合にのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p112">Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the result set will start after skipping the specified number of items. Optional. You can only use this option when you use the object notation option.</span></span>|

<span data-ttu-id="7ec69-170">次のコード サンプルは、コレクション内の各ワークシートの使用範囲の `name` プロパティと `address` を選択することにより、ワークシートのコレクションを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-170">The following code sample loads a workskeet collection by selecting the `name` property and the `address` of the used range for each worksheet in the collection.</span></span> <span data-ttu-id="7ec69-171">また、コレクションの上位 5 つのワークシートのみを読み込むように指定しています。</span><span class="sxs-lookup"><span data-stu-id="7ec69-171">It also specifies that only the top five worksheets in the collection should be loaded.</span></span> <span data-ttu-id="7ec69-172">`top: 10` と `skip: 5` を属性値として指定することで、次の 5 つのワークシートのセットを処理できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-172">You could process the next set of five worksheets by specifying `top: 10` and `skip: 5` as attribute values.</span></span> 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a><span data-ttu-id="7ec69-173">スカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="7ec69-173">Scalar and navigation properties</span></span> 

<span data-ttu-id="7ec69-174">Excel JavaScript API のリファレンス ドキュメントでは、オブジェクトのメンバーが**プロパティ**と**リレーションシップ**の 2 つのカテゴリにグループ化されています。</span><span class="sxs-lookup"><span data-stu-id="7ec69-174">In the Excel JavaScript API reference documentation, you may notice that object members are grouped into two categories: **properties** and **relationships**.</span></span> <span data-ttu-id="7ec69-175">オブジェクトのプロパティは、文字列、整数、ブール値などのスカラー メンバーです。一方、オブジェクトのリレーションシップ (ナビゲーション プロパティとも呼ばれる) は、オブジェクトまたはオブジェクトのコレクションのいずれかであるメンバーです。</span><span class="sxs-lookup"><span data-stu-id="7ec69-175">A property of an object is a scalar member such as a string, an integer, or a boolean value, while a relationship of an object (also known as a navigation property) is a member that is either an object or collection of objects.</span></span> <span data-ttu-id="7ec69-176">たとえば、[Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) オブジェクトの `name` メンバーと `position` メンバーはスカラー プロパティですが、`protection` と `tables` はリレーションシップ (ナビゲーション プロパティ) です。</span><span class="sxs-lookup"><span data-stu-id="7ec69-176">For example, `name` and `position` members on the [Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are relationships (navigation properties).</span></span> 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a><span data-ttu-id="7ec69-177">`object.load()` を使用したスカラー プロパティとナビゲーション プロパティ</span><span class="sxs-lookup"><span data-stu-id="7ec69-177">Scalar properties and navigation properties with `object.load()`</span></span>

<span data-ttu-id="7ec69-178">パラメーターを指定しないで `object.load()` メソッドを呼び出すと、オブジェクトのすべてのスカラー プロパティが読み込まれます。オブジェクトのナビゲーション プロパティは読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="7ec69-178">Calling the `object.load()` method with no parameters specified will load all scalar properties of the object; navigation properties of the object will not be loaded.</span></span> <span data-ttu-id="7ec69-179">さらに、ナビゲーション プロパティを直接読み込むことはできません。</span><span class="sxs-lookup"><span data-stu-id="7ec69-179">Additionally, navigation properties cannot be loaded directly.</span></span> <span data-ttu-id="7ec69-180">その代わりに、`load()` メソッドを使用して、目的のナビゲーション プロパティ内のスカラー プロパティを個別に参照する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7ec69-180">Instead, you should use the `load()` method to reference individual scalar properties within the desired navigation property.</span></span> <span data-ttu-id="7ec69-181">たとえば、範囲のフォント名を読み込むには、**name** プロパティへのパスとして **format** および **font** ナビゲーション プロパティを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7ec69-181">For example, to load the font name for a range, you must specify the **format** and **font** navigation properties as the path to the **name** property:</span></span>

```js
someRange.load("format/font/name")
```

> [!NOTE]
> <span data-ttu-id="7ec69-182">Excel JavaScript API を使用すると、パスを詳しく調べることでナビゲーション プロパティのスカラー プロパティを設定できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-182">With the Excel JavaScript API, you can set scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="7ec69-183">たとえば、`someRange.format.font.size = 10;` を使用して範囲のフォント サイズを設定できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-183">For example, you could set the font size for a range by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="7ec69-184">設定前にプロパティを読み込む必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7ec69-184">You do not need to load the property before you set it.</span></span> 

## <a name="setting-properties-of-an-object"></a><span data-ttu-id="7ec69-185">オブジェクトのプロパティを設定する</span><span class="sxs-lookup"><span data-stu-id="7ec69-185">Setting properties of an object</span></span>

<span data-ttu-id="7ec69-186">入れ子になったナビゲーション プロパティを持つオブジェクトのプロパティを設定するのは面倒です。</span><span class="sxs-lookup"><span data-stu-id="7ec69-186">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="7ec69-187">前述のナビゲーション パスを使用してプロパティを個別に設定する代わりに、Excel JavaScript API のすべてのオブジェクトで使用できる、`object.set()` メソッドを使用できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-187">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on all objects in the Excel JavaScript API.</span></span> <span data-ttu-id="7ec69-188">このメソッドを使用すると、同じ Office.js 型の別のオブジェクト、またはメソッドが呼び出されるオブジェクトのプロパティと同様に構造化されたプロパティを持つ JavaScript オブジェクトを渡すことによって、オブジェクトの複数のプロパティを一度に設定できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-188">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

> [!NOTE]
> <span data-ttu-id="7ec69-189">`set()` メソッドは、Excel JavaScript API などホスト固有の Office JavaScript API のオブジェクトでのみ実装されます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-189">The `set()` method is implemented only for objects within the host-specific Office JavaScript APIs, such as the Excel JavaScript API.</span></span> <span data-ttu-id="7ec69-190">共通 (共有) API は、このメソッドをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="7ec69-190">The common (shared) APIs do not support this method.</span></span> 

### <a name="set-properties-object-options-object"></a><span data-ttu-id="7ec69-191">set (properties: object, options: object)</span><span class="sxs-lookup"><span data-stu-id="7ec69-191">set (properties: object, options: object)</span></span>

<span data-ttu-id="7ec69-p119">メソッドが呼び出されるオブジェクトのプロパティは、渡されたオブジェクトの対応するプロパテに指定された値に設定されます。`properties` パラメーターが JavaScript オブジェクトの場合、メソッドが呼び出される読み取り専用プロパティに対応する渡されたオブジェクトの任意のプロパティは、`options` パラメーターの値に応じて、無視されるか、例外のスローが発生します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p119">Properties of the object on which the method is called are set to the values that are specified by the corresponding properties of the passed-in object. If the `properties` parameter is a JavaScript object, any property of the passed-in object that corresponds to a read-only property in the object on which the method is called will either be ignored or cause an exception to be thrown, depending on the value of the `options` parameter.</span></span>

#### <a name="syntax"></a><span data-ttu-id="7ec69-194">構文</span><span class="sxs-lookup"><span data-stu-id="7ec69-194">Syntax</span></span>

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a><span data-ttu-id="7ec69-195">パラメーター</span><span class="sxs-lookup"><span data-stu-id="7ec69-195">Parameters</span></span>

|<span data-ttu-id="7ec69-196">**パラメーター**</span><span class="sxs-lookup"><span data-stu-id="7ec69-196">**Parameter**</span></span>|<span data-ttu-id="7ec69-197">**型**</span><span class="sxs-lookup"><span data-stu-id="7ec69-197">**Type**</span></span>|<span data-ttu-id="7ec69-198">**説明**</span><span class="sxs-lookup"><span data-stu-id="7ec69-198">**Description**</span></span>|
|:------------|:--------|:----------|
|`properties`|<span data-ttu-id="7ec69-199">object</span><span class="sxs-lookup"><span data-stu-id="7ec69-199">object</span></span>|<span data-ttu-id="7ec69-200">メソッドが呼び出されるオブジェクトの同じ Office.js 型のオブジェクト、またはメソッドが呼び出されるオブジェクトの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトのいずれかです。</span><span class="sxs-lookup"><span data-stu-id="7ec69-200">Either an object of the same Office.js type of the object on which the method is called, or a JavaScript object with property names and types that mirror the structure of the object on which the method is called.</span></span>|
|`options`|<span data-ttu-id="7ec69-201">object</span><span class="sxs-lookup"><span data-stu-id="7ec69-201">object</span></span>|<span data-ttu-id="7ec69-p120">省略可能。最初のパラメーターが JavaScript オブジェクトの場合にのみ渡すことができます。オブジェクトには、次のプロパティを含めることができます。`throwOnReadOnly?: boolean` (既定値は `true`。渡された JavaScript オブジェクトに読み取り専用プロパティが含まれている場合は、エラーをスローします。)</span><span class="sxs-lookup"><span data-stu-id="7ec69-p120">Optional. Can only be passed when the first parameter is a JavaScript object. The object can contain the following property: `throwOnReadOnly?: boolean` (Default is `true`: throw an error if the passed in JavaScript object includes read-only properties.)</span></span>|

#### <a name="returns"></a><span data-ttu-id="7ec69-205">戻り値</span><span class="sxs-lookup"><span data-stu-id="7ec69-205">Returns</span></span>

<span data-ttu-id="7ec69-206">void</span><span class="sxs-lookup"><span data-stu-id="7ec69-206">void</span></span>    

#### <a name="example"></a><span data-ttu-id="7ec69-207">例</span><span class="sxs-lookup"><span data-stu-id="7ec69-207">Example</span></span>

<span data-ttu-id="7ec69-p121">次のコード サンプルは、`set()` メソッドを呼び出し、**Range** オブジェクトのプロパティの構造を反映するプロパティ名と型を持つ JavaScript オブジェクトを渡すことによって、範囲のいくつかの書式プロパティを設定します。この例では、範囲 **B2:E2** にデータがあると仮定します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-p121">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the **Range** object. This example assumes that there is data in range **B2:E2**.</span></span>

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
## <a name="42ornullobject-methods"></a><span data-ttu-id="7ec69-210">&#42;OrNullObject メソッド</span><span class="sxs-lookup"><span data-stu-id="7ec69-210">&#42;OrNullObject methods</span></span>

<span data-ttu-id="7ec69-211">多くの Excel JavaScript API メソッドは、API の条件が満たされない場合に例外を返します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-211">Many Excel JavaScript API methods will return an exception when the condition of the API is not met.</span></span> <span data-ttu-id="7ec69-212">たとえば、ブックに存在しないワークシート名を指定してワークシートを取得しようとすると、`getItem()` メソッドは `ItemNotFound` 例外を返します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-212">For example, if you attempt to get a worksheet by specifying a worksheet name that doesn't exist in the workbook, the `getItem()` method will return an `ItemNotFound` exception.</span></span> 

<span data-ttu-id="7ec69-213">このようなシナリオの複雑な例外処理ロジックを実装する代わりに、Excel JavaScript API のいくつかのメソッドで使用できる `*OrNullObject` メソッドのバリエーションを使用できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-213">Instead of implementing complex exception handling logic for scenarios like this, you can use the `*OrNullObject` method variant that's available for several methods in the Excel JavaScript API.</span></span> <span data-ttu-id="7ec69-214">指定された項目が存在しない場合、`*OrNullObject` メソッドは例外をスローするのではなく、null オブジェクト (JavaScript `null` ではない) を返します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-214">An `*OrNullObject` method will return a null object (not the JavaScript `null`) rather than throwing an exception if the specified item doesn't exist.</span></span> <span data-ttu-id="7ec69-215">たとえば、**Worksheets** などのコレクションで `getItemOrNullObject()` メソッドを呼び出して、コレクションからのアイテムの取得を試行できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-215">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to attempt to retrieve an item from the collection.</span></span> <span data-ttu-id="7ec69-216">`getItemOrNullObject()` メソッドは、指定された項目が存在する場合はその項目を返し、それ以外の場合は null オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-216">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns a null object.</span></span> <span data-ttu-id="7ec69-217">返される null オブジェクトには、ブール型プロパティ `isNullObject` が含まれています。これを評価して、オブジェクトが存在するかどうかを判断できます。</span><span class="sxs-lookup"><span data-stu-id="7ec69-217">The null object that is returned contains the boolean property `isNullObject` that you can evaluate to determine whether the object exists.</span></span>

<span data-ttu-id="7ec69-218">次のコード サンプルは `getItemOrNullObject()` メソッドを使用して、"Data" という名前のワークシートの取得を試行します。</span><span class="sxs-lookup"><span data-stu-id="7ec69-218">The following code sample attempts to retrieve a worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="7ec69-219">メソッドが null オブジェクトを返す場合は、新しいシートを作成し、そのシート上で操作を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7ec69-219">If the method returns a null object, a new sheet needs to be created before actions can taken on the sheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="7ec69-220">関連項目</span><span class="sxs-lookup"><span data-stu-id="7ec69-220">See also</span></span>
 
* [<span data-ttu-id="7ec69-221">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="7ec69-221">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
* <span data-ttu-id="7ec69-222">
  [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)</span><span class="sxs-lookup"><span data-stu-id="7ec69-222">[Excel add-ins code samples](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)</span></span>
* [<span data-ttu-id="7ec69-223">Excel の JavaScript API を使用した、パフォーマンスの最適化</span><span class="sxs-lookup"><span data-stu-id="7ec69-223">Excel JavaScript API performance optimization</span></span>](performance.md)
* [<span data-ttu-id="7ec69-224">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7ec69-224">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
