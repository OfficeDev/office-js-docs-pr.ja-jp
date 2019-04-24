---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 11/16/2018
localization_priority: Normal
ms.openlocfilehash: 8eeaedffcc143b0e8d61e9c151f3786b67a0e3fc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449730"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="ca343-103">マニフェスト要素の正しい順序を確認する方法</span><span class="sxs-lookup"><span data-stu-id="ca343-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="ca343-104">Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="ca343-105">必要な順序は、[[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] フォルダー内の XSD ファイルで指定されています。</span><span class="sxs-lookup"><span data-stu-id="ca343-105">The required ordering is specified in the XSD files in the [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) folder.</span></span> <span data-ttu-id="ca343-106">XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。</span><span class="sxs-lookup"><span data-stu-id="ca343-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="ca343-107">例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-107">For example, in the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="ca343-108">`<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="ca343-109">順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="ca343-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="ca343-110">[Office アドイン検証ツール](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator)では、要素の順序が間違っている場合と、要素が間違った親の下にある場合とで、同じエラー メッセージが使用されます。</span><span class="sxs-lookup"><span data-stu-id="ca343-110">The [Office Add-in Validator](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="ca343-111">エラーには、子要素が親要素の有効な子ではないと表示されます。</span><span class="sxs-lookup"><span data-stu-id="ca343-111">The error says the child element is not a valid child of the parent element.</span></span> <span data-ttu-id="ca343-112">そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子*である*と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。</span><span class="sxs-lookup"><span data-stu-id="ca343-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="ca343-113">ある親要素の子要素の正しい順序を確認するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="ca343-113">To find the correct order for the child elements of a given parent element, take the following steps.</span></span> <span data-ttu-id="ca343-114">(XDSD ファイルは非常に複雑なため、ここでご紹介するのは簡素化されたプロセスです。</span><span class="sxs-lookup"><span data-stu-id="ca343-114">(This is a simplified process, as XSD files are quite complex.</span></span> <span data-ttu-id="ca343-115">XSD ファイルの詳細な説明はこのドキュメントの目的の範囲外です。)</span><span class="sxs-lookup"><span data-stu-id="ca343-115">Fully parsing XSD files is out of the scope of this document.)</span></span>

1. <span data-ttu-id="ca343-116">作成するアドインの種類の [[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] の下にある サブフォルダーを開きます。</span><span class="sxs-lookup"><span data-stu-id="ca343-116">Open the subfolder under [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) for the type of add-in that you are creating.</span></span> 
2. <span data-ttu-id="ca343-117">親要素が複合型として定義されている XSD ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ca343-117">Open the XSD file where the parent element is defined as a complex type.</span></span> <span data-ttu-id="ca343-118">どのファイルが複合型と定義されているかがわからない場合は、確認できるまで複数のファイルで手順 3 を行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-118">If you don't know which file has the definition, you may have to do step 3 on multiple files until you find it.</span></span>
3. <span data-ttu-id="ca343-119">`<xs:complexType name="PARENT_ELEMENT">` を検索します。"PARENT_ELEMENT" は親要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="ca343-119">Search for `<xs:complexType name="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element.</span></span>
4. <span data-ttu-id="ca343-120">PARENT_ELEMENT の定義の中に、`<xs:sequence>` という名前の要素が (通常は) あります。</span><span class="sxs-lookup"><span data-stu-id="ca343-120">Inside the definition for the PARENT_ELEMENT, there is (usually) an element called `<xs:sequence>`.</span></span> <span data-ttu-id="ca343-121">[TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd) での `<SuperTip>` の定義を次に示します。</span><span class="sxs-lookup"><span data-stu-id="ca343-121">The following is the definition for `<SuperTip>` from [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd).</span></span>

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

<span data-ttu-id="ca343-122">`<xs:sequence>` は、含まれる可能性のある子要素を*正しい表示順序で*一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="ca343-122">The `<xs:sequence>` lists the possible child elements, *in the order in which they must appear*.</span></span> <span data-ttu-id="ca343-123">ここに含まれる要素が必須という意味では*ありません*。</span><span class="sxs-lookup"><span data-stu-id="ca343-123">This does *not* mean all of them are mandatory.</span></span> <span data-ttu-id="ca343-124">子要素の `minOccurs` の値が **0** の場合、この子要素は省略可能です。</span><span class="sxs-lookup"><span data-stu-id="ca343-124">If the `minOccurs` value for a child element is **0**, then the child element is optional.</span></span> <span data-ttu-id="ca343-125">*ただし、この子要素が含まれる場合は、`<xs:sequence>` 要素で指定された順序で配置する必要があります*。</span><span class="sxs-lookup"><span data-stu-id="ca343-125">*But if it is present, it must be in the order specified by the `<xs:sequence>` element*.</span></span>

<span data-ttu-id="ca343-126">`<xs:sequence>` 要素がない場合、または*ある*場合でもこの子要素が一覧表示されていない場合 (子要素のレファレンス ドキュメントで、この子要素が親要素の有効な子で*ある*と示される場合でも)、XSD ファイルのどこかで親要素の複合型の定義が追加の子要素によって拡張されています。</span><span class="sxs-lookup"><span data-stu-id="ca343-126">If there is no `<xs:sequence>` element, or there *is* but the child element is not listed (even though the reference documentation for the child element indicates that it *is* valid for the parent), then the parent element's complex type definition has been extended with additional child elements somewhere else in the XSD file.</span></span> <span data-ttu-id="ca343-127">たとえば、`OfficeApp` 複合型の定義は、使用可能な子として `Requirements` を一覧に表示しません。</span><span class="sxs-lookup"><span data-stu-id="ca343-127">For example, the definition for the `OfficeApp` complex type does not list `Requirements` as a possible child.</span></span> <span data-ttu-id="ca343-128">ただし、ファイルの後の方で (`TaskPaneApp` 複合型 の定義の中で) `OfficeApp` の定義は拡張され、追加の有効な子として `Requirements` が追加されています。</span><span class="sxs-lookup"><span data-stu-id="ca343-128">But later in the file (within the definition for the `TaskPaneApp` complex type), the definition of `OfficeApp` is extended and `Requirements` is added as an additional valid child.</span></span>

<span data-ttu-id="ca343-129">拡張された定義を探すには次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="ca343-129">To find the extended definitions follow these steps:</span></span>

1. <span data-ttu-id="ca343-130">ファイルの最初から初めて、`<xs:extension base="PARENT_ELEMENT">` を検索します。"PARENT_ELEMENT" は親要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="ca343-130">Starting at the top of the file, search for `<xs:extension base="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element.</span></span> <span data-ttu-id="ca343-131">拡張された定義は複数ある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-131">There may be more than one extension.</span></span>
2. <span data-ttu-id="ca343-132">作業中のコンテキストに関連する拡張された定義を探します。</span><span class="sxs-lookup"><span data-stu-id="ca343-132">Find the extension that is relevant to the context in which you are working.</span></span> <span data-ttu-id="ca343-133">たとえば、`OfficeApp` 複合型は `ContentApp`、`MailApp` および `TaskPaneApp` 複合型の中でそれぞれ拡張されます。</span><span class="sxs-lookup"><span data-stu-id="ca343-133">For example, the `OfficeApp` complex type is extended within the `ContentApp` and `MailApp` complex types as well as within the `TaskPaneApp` complex type.</span></span>

<span data-ttu-id="ca343-134">ファイル内の各 `<xs:extension base="PARENT_ELEMENT">` には独自の `<xs:sequence>` があり、親要素に対して有効な追加の子要素を一覧表示します。</span><span class="sxs-lookup"><span data-stu-id="ca343-134">Each `<xs:extension base="PARENT_ELEMENT">` in the file has its own `<xs:sequence>` that lists additional valid child elements for the parent.</span></span> <span data-ttu-id="ca343-135">拡張された一覧上の子要素は、常に親要素の複合型定義内の元のリストの子要素の*後ろに*配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ca343-135">Child elements on an extended list must always be *after* the child elements in the original list in the parent's complex type definition.</span></span>

## <a name="see-also"></a><span data-ttu-id="ca343-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="ca343-136">See also</span></span>

- [<span data-ttu-id="ca343-137">Office アドイン マニフェストのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="ca343-137">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
