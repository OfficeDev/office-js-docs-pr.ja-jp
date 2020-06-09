---
ms.date: 04/29/2020
description: Excel カスタム関数をローカライズします。
title: カスタム関数をローカライズする
localization_priority: Normal
ms.openlocfilehash: 427bff029c5e85caa216f628df450525ee187c17
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609297"
---
# <a name="localize-custom-functions"></a><span data-ttu-id="8552d-103">カスタム関数をローカライズする</span><span class="sxs-lookup"><span data-stu-id="8552d-103">Localize custom functions</span></span>

<span data-ttu-id="8552d-104">アドインとカスタム関数名の両方をローカライズできます。</span><span class="sxs-lookup"><span data-stu-id="8552d-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="8552d-105">そのためには、ローカライズされた関数名を、XML マニフェストファイルの関数の JSON ファイルとロケール情報に提供します。</span><span class="sxs-lookup"><span data-stu-id="8552d-105">To do so, provide localized function names in the functions' JSON file and locale information in the XML manifest file.</span></span>

>[!IMPORTANT]
> <span data-ttu-id="8552d-106">自動生成されたメタデータはローカライズには機能しないため、JSON ファイルを手動で更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8552d-106">Auto-generated metadata doesn't work for localization so you need to update the JSON file manually.</span></span> <span data-ttu-id="8552d-107">これを行う方法については、「 [Excel のカスタム関数のメタデータ](custom-functions-json.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8552d-107">To learn how to do this, see [Metadata for custom functions in Excel](custom-functions-json.md)</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a><span data-ttu-id="8552d-108">関数名をローカライズする</span><span class="sxs-lookup"><span data-stu-id="8552d-108">Localize function names</span></span>

<span data-ttu-id="8552d-109">カスタム関数をローカライズするには、言語ごとに新しい JSON メタデータファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="8552d-109">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="8552d-110">各言語 JSON ファイルで、 `name` `description` ターゲット言語でプロパティを作成します。</span><span class="sxs-lookup"><span data-stu-id="8552d-110">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="8552d-111">英語の既定のファイルの名前は、**関数 json**です。</span><span class="sxs-lookup"><span data-stu-id="8552d-111">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="8552d-112">各 JSON ファイルのファイル名にロケールを使用します。たとえば、**関数**を識別するために使用します。</span><span class="sxs-lookup"><span data-stu-id="8552d-112">Use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="8552d-113">は `name` 、 `description` Excel に表示され、ローカライズされています。</span><span class="sxs-lookup"><span data-stu-id="8552d-113">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="8552d-114">ただし、 `id` 各関数のはローカライズされていません。</span><span class="sxs-lookup"><span data-stu-id="8552d-114">However, the `id` of each function isn't localized.</span></span> <span data-ttu-id="8552d-115">`id`このプロパティでは、Excel によって関数が一意であると識別されますが、設定された後に変更することはできません。</span><span class="sxs-lookup"><span data-stu-id="8552d-115">The `id` property is how Excel identifies your function as unique and shouldn't be changed once it is set.</span></span>

<span data-ttu-id="8552d-116">次の JSON は、"掛け算" というプロパティを持つ関数を定義する方法を示して `id` います。</span><span class="sxs-lookup"><span data-stu-id="8552d-116">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="8552d-117">`name` `description` 関数のおよびプロパティは、ドイツ語にローカライズされています。</span><span class="sxs-lookup"><span data-stu-id="8552d-117">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="8552d-118">各パラメーター `name` と `description` は、ドイツ語にローカライズされています。</span><span class="sxs-lookup"><span data-stu-id="8552d-118">Each parameter `name` and `description` is also localized for German.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

<span data-ttu-id="8552d-119">前の JSON を次の JSON と比較して英語を比較します。</span><span class="sxs-lookup"><span data-stu-id="8552d-119">Compare the previous JSON with the following JSON for English.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a><span data-ttu-id="8552d-120">アドインをローカライズする</span><span class="sxs-lookup"><span data-stu-id="8552d-120">Localize your add-in</span></span>

<span data-ttu-id="8552d-121">各言語の JSON ファイルを作成した後、各 JSON メタデータファイルの URL を指定する各ロケールの上書き値で XML マニフェストファイルを更新します。</span><span class="sxs-lookup"><span data-stu-id="8552d-121">After creating a JSON file for each language, update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="8552d-122">次のマニフェスト XML は、 `en-us` (ドイツ) 用の JSON ファイルの上書き URL を含む既定のロケールを示して `de-de` います。</span><span class="sxs-lookup"><span data-stu-id="8552d-122">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="8552d-123">**関数の de**ファイルには、ローカライズされたドイツ語の関数名と id が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8552d-123">The **functions-de.json** file contains the localized German function names and ids.</span></span>

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

<span data-ttu-id="8552d-124">アドインのローカライズプロセスの詳細については、「 [Office アドインのローカライズ](../develop/localization.md#control-localization-from-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8552d-124">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="8552d-125">次の手順</span><span class="sxs-lookup"><span data-stu-id="8552d-125">Next steps</span></span>
<span data-ttu-id="8552d-126">[カスタム関数の名前付け規則](custom-functions-naming.md)について、または[エラー処理のベストプラクティス](custom-functions-errors.md)を検出する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="8552d-126">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8552d-127">関連項目</span><span class="sxs-lookup"><span data-stu-id="8552d-127">See also</span></span>

* [<span data-ttu-id="8552d-128">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="8552d-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8552d-129">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="8552d-129">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="8552d-130">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="8552d-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)
