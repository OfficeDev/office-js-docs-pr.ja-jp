---
title: マニフェスト ファイルの ExtendedOverrides 要素
description: マニフェストの JSON 形式の拡張子の URL を指定します。
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505473"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="c3c29-103">ExtendedOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="c3c29-103">ExtendedOverrides element</span></span>

<span data-ttu-id="c3c29-104">マニフェストを拡張する JSON 形式のファイルの完全な URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="c3c29-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span> <span data-ttu-id="c3c29-105">この要素とその子孫要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 [を参照してください](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="c3c29-105">For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="c3c29-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c3c29-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c3c29-107">構文</span><span class="sxs-lookup"><span data-stu-id="c3c29-107">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="c3c29-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="c3c29-108">Contained in</span></span>

[<span data-ttu-id="c3c29-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c3c29-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c3c29-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="c3c29-110">Can contain</span></span>

|<span data-ttu-id="c3c29-111">要素</span><span class="sxs-lookup"><span data-stu-id="c3c29-111">Element</span></span>|<span data-ttu-id="c3c29-112">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c3c29-112">Content</span></span>|<span data-ttu-id="c3c29-113">メール</span><span class="sxs-lookup"><span data-stu-id="c3c29-113">Mail</span></span>|<span data-ttu-id="c3c29-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="c3c29-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="c3c29-115">トークン</span><span class="sxs-lookup"><span data-stu-id="c3c29-115">Tokens</span></span>](tokens.md)|||<span data-ttu-id="c3c29-116">x</span><span class="sxs-lookup"><span data-stu-id="c3c29-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="c3c29-117">属性</span><span class="sxs-lookup"><span data-stu-id="c3c29-117">Attributes</span></span>

|<span data-ttu-id="c3c29-118">属性</span><span class="sxs-lookup"><span data-stu-id="c3c29-118">Attribute</span></span>|<span data-ttu-id="c3c29-119">説明</span><span class="sxs-lookup"><span data-stu-id="c3c29-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="c3c29-120">URL (必須)</span><span class="sxs-lookup"><span data-stu-id="c3c29-120">Url (required)</span></span>| <span data-ttu-id="c3c29-121">拡張の完全な URL は JSON ファイルを上書きします。</span><span class="sxs-lookup"><span data-stu-id="c3c29-121">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="c3c29-122">将来、この値は、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。</span><span class="sxs-lookup"><span data-stu-id="c3c29-122">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="c3c29-123">「 [例」を参照してください](#examples)。</span><span class="sxs-lookup"><span data-stu-id="c3c29-123">See [Examples](#examples).</span></span>|
|<span data-ttu-id="c3c29-124">ResourcesUrl (オプション)</span><span class="sxs-lookup"><span data-stu-id="c3c29-124">ResourcesUrl (optional)</span></span> | <span data-ttu-id="c3c29-125">属性で指定されたファイルの、ローカライズされた文字列などの補足リソースを提供するファイルの完全な `Url` URL。</span><span class="sxs-lookup"><span data-stu-id="c3c29-125">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="c3c29-126">これは、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。</span><span class="sxs-lookup"><span data-stu-id="c3c29-126">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="examples"></a><span data-ttu-id="c3c29-127">例</span><span class="sxs-lookup"><span data-stu-id="c3c29-127">Examples</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="c3c29-128">将来、この値は、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。</span><span class="sxs-lookup"><span data-stu-id="c3c29-128">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="c3c29-129">次に例を示します。</span><span class="sxs-lookup"><span data-stu-id="c3c29-129">The following is an example.</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```
