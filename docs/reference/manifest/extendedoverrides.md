---
title: マニフェストファイルの ExtendedOverrides 要素
description: マニフェストの JSON 形式の拡張機能の Url を指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996709"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="72653-103">ExtendedOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="72653-103">ExtendedOverrides element</span></span>

<span data-ttu-id="72653-104">マニフェストを拡張する JSON 形式のファイルの完全な Url を指定します。</span><span class="sxs-lookup"><span data-stu-id="72653-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span>

<span data-ttu-id="72653-105">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72653-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="72653-106">構文</span><span class="sxs-lookup"><span data-stu-id="72653-106">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="72653-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="72653-107">Contained in</span></span>

[<span data-ttu-id="72653-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="72653-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="72653-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="72653-109">Can contain</span></span>

|<span data-ttu-id="72653-110">要素</span><span class="sxs-lookup"><span data-stu-id="72653-110">Element</span></span>|<span data-ttu-id="72653-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72653-111">Content</span></span>|<span data-ttu-id="72653-112">メール</span><span class="sxs-lookup"><span data-stu-id="72653-112">Mail</span></span>|<span data-ttu-id="72653-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="72653-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="72653-114">トークン</span><span class="sxs-lookup"><span data-stu-id="72653-114">Tokens</span></span>](tokens.md)|||<span data-ttu-id="72653-115">x</span><span class="sxs-lookup"><span data-stu-id="72653-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="72653-116">属性</span><span class="sxs-lookup"><span data-stu-id="72653-116">Attributes</span></span>

|<span data-ttu-id="72653-117">属性</span><span class="sxs-lookup"><span data-stu-id="72653-117">Attribute</span></span>|<span data-ttu-id="72653-118">説明</span><span class="sxs-lookup"><span data-stu-id="72653-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="72653-119">Url (必須)</span><span class="sxs-lookup"><span data-stu-id="72653-119">Url (required)</span></span>| <span data-ttu-id="72653-120">拡張オーバーライド JSON ファイルの完全な URL。</span><span class="sxs-lookup"><span data-stu-id="72653-120">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="72653-121">これは、 [token](tokens.md) 要素によって定義されたトークンを使用する URL テンプレートである場合があります。</span><span class="sxs-lookup"><span data-stu-id="72653-121">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|
|<span data-ttu-id="72653-122">ResourcesUrl (省略可能)</span><span class="sxs-lookup"><span data-stu-id="72653-122">ResourcesUrl (optional)</span></span> | <span data-ttu-id="72653-123">属性で指定されているファイルについて、ローカライズされた文字列などの補足情報を提供するファイルの完全な URL `Url` 。</span><span class="sxs-lookup"><span data-stu-id="72653-123">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="72653-124">これは、 [token](tokens.md) 要素によって定義されたトークンを使用する URL テンプレートである場合があります。</span><span class="sxs-lookup"><span data-stu-id="72653-124">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="example"></a><span data-ttu-id="72653-125">例</span><span class="sxs-lookup"><span data-stu-id="72653-125">Example</span></span>

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
