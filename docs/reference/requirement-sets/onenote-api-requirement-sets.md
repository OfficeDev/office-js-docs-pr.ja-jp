---
title: OneNote JavaScript API の要件セット
description: OneNote JavaScript API の要件セットについて説明します。
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350191"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="3a8bc-103">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="3a8bc-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="3a8bc-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3a8bc-107">次の表は、OneNote の要件セット、それらの要件セットをサポートする Office クライアント アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-107">The following table lists the OneNote requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="3a8bc-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="3a8bc-108">Requirement set</span></span>  |  <span data-ttu-id="3a8bc-109">Office on the web</span><span class="sxs-lookup"><span data-stu-id="3a8bc-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="3a8bc-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="3a8bc-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | <span data-ttu-id="3a8bc-111">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="3a8bc-111">September 2016</span></span> |  

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="3a8bc-112">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="3a8bc-112">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="3a8bc-p102">OneNote JavaScript API 1.1 は、API の最初のバージョンです。API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](../../onenote/onenote-add-ins-programming-overview.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-p102">OneNote JavaScript API 1.1 is the first version of the API. For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="3a8bc-115">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="3a8bc-115">Runtime requirement support check</span></span>

<span data-ttu-id="3a8bc-116">実行時に、アドインは次を行うことによって、特定の Office アプリケーションが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-116">At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following:</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="3a8bc-117">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="3a8bc-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="3a8bc-p103">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。Office アプリケーションまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのアプリケーションまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-p103">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="3a8bc-120">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office クライアント アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-120">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3a8bc-121">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="3a8bc-121">Office Common API requirement sets</span></span>

<span data-ttu-id="3a8bc-122">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="3a8bc-122">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3a8bc-123">関連項目</span><span class="sxs-lookup"><span data-stu-id="3a8bc-123">See also</span></span>

- [<span data-ttu-id="3a8bc-124">OneNote JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="3a8bc-124">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="3a8bc-125">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="3a8bc-125">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3a8bc-126">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="3a8bc-126">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3a8bc-127">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="3a8bc-127">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
