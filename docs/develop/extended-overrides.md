---
title: マニフェストの拡張オーバーライドを処理する
description: マニフェストの拡張オーバーライドを使用して機能拡張機能を構成する方法について学習します。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505573"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a><span data-ttu-id="659b9-103">マニフェストの拡張オーバーライドを使用する</span><span class="sxs-lookup"><span data-stu-id="659b9-103">Work with Extended Overrides of the manifest</span></span>

<span data-ttu-id="659b9-104">Office アドインの一部の機能拡張機能は、アドインの XML マニフェストではなく、サーバーでホストされている JSON ファイルで構成されます。</span><span class="sxs-lookup"><span data-stu-id="659b9-104">Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="659b9-105">この記事では、アドイン マニフェストOfficeアドインでの役割について理解している必要があります。最近Office [場合は、アドインの XML](add-in-manifests.md)マニフェストを参照してください。</span><span class="sxs-lookup"><span data-stu-id="659b9-105">This article assumes that you're familiar with Office add-in manifests and their role in add-ins. Please read [Office Add-ins XML manifest](add-in-manifests.md), if you haven't recently.</span></span>

<span data-ttu-id="659b9-106">次の表は、機能のドキュメントへのリンクと共に、拡張オーバーライドを必要とする機能拡張機能を指定します。</span><span class="sxs-lookup"><span data-stu-id="659b9-106">The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.</span></span>

| <span data-ttu-id="659b9-107">機能</span><span class="sxs-lookup"><span data-stu-id="659b9-107">Feature</span></span> | <span data-ttu-id="659b9-108">開発手順</span><span class="sxs-lookup"><span data-stu-id="659b9-108">Development Instructions</span></span> |
| :----- | :----- |
| <span data-ttu-id="659b9-109">キーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="659b9-109">Keyboard shortcuts</span></span> | [<span data-ttu-id="659b9-110">カスタム キーボード ショートカットをアドインOffice追加する</span><span class="sxs-lookup"><span data-stu-id="659b9-110">Add Custom keyboard shortcuts to your Office Add-ins</span></span>](../design/keyboard-shortcuts.md) |

<span data-ttu-id="659b9-111">JSON 形式を定義するスキーマは [、拡張マニフェスト スキーマです](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。</span><span class="sxs-lookup"><span data-stu-id="659b9-111">The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!TIP]
> <span data-ttu-id="659b9-112">この記事はやや抽象的です。</span><span class="sxs-lookup"><span data-stu-id="659b9-112">This article is somewhat abstract.</span></span> <span data-ttu-id="659b9-113">表の記事の 1 つを読んで、概念をわかりやすくする方法を検討してください。</span><span class="sxs-lookup"><span data-stu-id="659b9-113">Consider reading one of the articles in the table to add clarity to the concepts.</span></span>

## <a name="tell-office-where-to-find-the-json-file"></a><span data-ttu-id="659b9-114">JSON ファイルOffice場所を確認する</span><span class="sxs-lookup"><span data-stu-id="659b9-114">Tell Office where to find the JSON file</span></span>

<span data-ttu-id="659b9-115">マニフェストを使用して、JSON Office場所を確認します。</span><span class="sxs-lookup"><span data-stu-id="659b9-115">Use the manifest to tell Office where to find the JSON file.</span></span> <span data-ttu-id="659b9-116">マニフェスト *内* の要素の直下 (内部ではない) `<VersionOverrides>` に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。</span><span class="sxs-lookup"><span data-stu-id="659b9-116">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="659b9-117">属性を `Url` JSON ファイルの完全な URL に設定します。</span><span class="sxs-lookup"><span data-stu-id="659b9-117">Set the `Url` attribute to the full URL of a JSON file.</span></span> <span data-ttu-id="659b9-118">最も単純な要素の例を次に示 `<ExtendedOverrides>` します。</span><span class="sxs-lookup"><span data-stu-id="659b9-118">The following is an example of the simplest possible `<ExtendedOverrides>` element.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="659b9-119">次に、非常に単純な拡張オーバーライド JSON ファイルの例を示します。</span><span class="sxs-lookup"><span data-stu-id="659b9-119">The following is an example of a very simple extended overrides JSON file.</span></span> <span data-ttu-id="659b9-120">これは、アドインの作業ウィンドウを開く関数 (他の場所で定義) にキーボード ショートカット Ctrl + Shift +A を割り当てる。</span><span class="sxs-lookup"><span data-stu-id="659b9-120">It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.</span></span>

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a><span data-ttu-id="659b9-121">拡張上書きファイルをローカライズする</span><span class="sxs-lookup"><span data-stu-id="659b9-121">Localize the extended overrides file</span></span>

<span data-ttu-id="659b9-122">アドインが複数のロケールをサポートしている場合は、要素の属性を使用して、ローカライズされたリソースのOfficeを `ResourceUrl` `<ExtendedOverrides>` ポイントできます。</span><span class="sxs-lookup"><span data-stu-id="659b9-122">If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the `<ExtendedOverrides>` element to point Office to a file of localized resources.</span></span> <span data-ttu-id="659b9-123">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="659b9-123">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="659b9-124">リソース ファイルを作成して使用する方法、拡張オーバーライド ファイル内のリソースを参照する方法、およびここで説明していない追加のオプションの詳細については [、「Localize extended overrides」](localization.md#localize-extended-overrides)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="659b9-124">For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).</span></span>
