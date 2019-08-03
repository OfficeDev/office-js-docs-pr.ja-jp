---
title: OneNote JavaScript API の要件セット
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 3a1e5133b36af612156fb272651f1775e916a0fe
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064873"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="4e763-102">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4e763-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="4e763-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4e763-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="4e763-106">次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="4e763-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="4e763-107">要件セット</span><span class="sxs-lookup"><span data-stu-id="4e763-107">Requirement set</span></span>  |  <span data-ttu-id="4e763-108">Web 上の Office</span><span class="sxs-lookup"><span data-stu-id="4e763-108">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="4e763-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="4e763-109">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="4e763-110">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="4e763-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="4e763-111">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4e763-111">Office Common API requirement sets</span></span>

<span data-ttu-id="4e763-112">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4e763-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="4e763-113">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="4e763-113">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="4e763-114">OneNote JavaScript API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="4e763-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="4e763-115">API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4e763-115">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="4e763-116">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="4e763-116">Runtime requirement support check</span></span>

<span data-ttu-id="4e763-117">実行時に、アドインは、次の手順に従って、特定のホストが API 要件セットをサポートしているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="4e763-117">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="4e763-118">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="4e763-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="4e763-119">アドインマニフェスト`Requirements`の要素を使用して、アドインが使用する必要がある重要な要件セットまたは API メンバーを指定します。</span><span class="sxs-lookup"><span data-stu-id="4e763-119">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="4e763-120">Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、アドインには表示されません。</span><span class="sxs-lookup"><span data-stu-id="4e763-120">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="4e763-121">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="4e763-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="4e763-122">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4e763-122">Office Common API requirement sets</span></span>

<span data-ttu-id="4e763-123">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4e763-123">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4e763-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="4e763-124">See also</span></span>

- [<span data-ttu-id="4e763-125">OneNote JavaScript API リファレンスドキュメント</span><span class="sxs-lookup"><span data-stu-id="4e763-125">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="4e763-126">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="4e763-126">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="4e763-127">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="4e763-127">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="4e763-128">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="4e763-128">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
