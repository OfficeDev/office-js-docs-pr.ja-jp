---
title: OneNote JavaScript API の要件セット
description: OneNote JavaScript API の要件セットの詳細
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 1adc3554cfce5cafa94afefdb1f2a2130817288e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611345"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="9865f-103">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="9865f-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="9865f-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9865f-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="9865f-107">次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="9865f-107">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="9865f-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="9865f-108">Requirement set</span></span>  |  <span data-ttu-id="9865f-109">Office on the web</span><span class="sxs-lookup"><span data-stu-id="9865f-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="9865f-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="9865f-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="9865f-111">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="9865f-111">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="9865f-112">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="9865f-112">Office Common API requirement sets</span></span>

<span data-ttu-id="9865f-113">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9865f-113">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="9865f-114">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="9865f-114">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="9865f-115">OneNote JavaScript API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="9865f-115">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="9865f-116">API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](../../onenote/onenote-add-ins-programming-overview.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9865f-116">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="9865f-117">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="9865f-117">Runtime requirement support check</span></span>

<span data-ttu-id="9865f-118">実行時に、アドインは次を行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="9865f-118">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="9865f-119">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="9865f-119">Manifest-based requirement support check</span></span>

<span data-ttu-id="9865f-120">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="9865f-120">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="9865f-121">Office ホストまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="9865f-121">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="9865f-122">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="9865f-122">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="9865f-123">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="9865f-123">Office Common API requirement sets</span></span>

<span data-ttu-id="9865f-124">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9865f-124">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9865f-125">関連項目</span><span class="sxs-lookup"><span data-stu-id="9865f-125">See also</span></span>

- [<span data-ttu-id="9865f-126">OneNote JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="9865f-126">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="9865f-127">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="9865f-127">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="9865f-128">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="9865f-128">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="9865f-129">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="9865f-129">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
