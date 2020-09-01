---
title: PowerPoint JavaScript API の要件セット
description: PowerPoint JavaScript API の要件セットの詳細情報。
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: b2b5d4b7b5a0677812f227b6a32683c35bbf1662
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293507"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="b2eef-103">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="b2eef-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="b2eef-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b2eef-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b2eef-107">次の表は、PowerPoint の要件セット、それらの要件セットをサポートする Office クライアント アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="b2eef-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="b2eef-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="b2eef-108">Requirement set</span></span>  |  <span data-ttu-id="b2eef-109">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="b2eef-109">Office on Windows</span></span><br><span data-ttu-id="b2eef-110">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b2eef-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="b2eef-111">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="b2eef-111">Office on iPad</span></span><br><span data-ttu-id="b2eef-112">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b2eef-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="b2eef-113">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="b2eef-113">Office on Mac</span></span><br><span data-ttu-id="b2eef-114">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b2eef-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="b2eef-115">Office on the web</span><span class="sxs-lookup"><span data-stu-id="b2eef-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="b2eef-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="b2eef-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="b2eef-117">バージョン 1810 (ビルド 11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="b2eef-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="b2eef-118">2.17 以降</span><span class="sxs-lookup"><span data-stu-id="b2eef-118">2.17 or later</span></span> | <span data-ttu-id="b2eef-119">16.19 以降</span><span class="sxs-lookup"><span data-stu-id="b2eef-119">16.19 or later</span></span> | <span data-ttu-id="b2eef-120">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="b2eef-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="b2eef-121">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="b2eef-121">Office versions and build numbers</span></span>

<span data-ttu-id="b2eef-122">Office のバージョンとビルド番号の詳細については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2eef-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="b2eef-123">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="b2eef-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="b2eef-124">PowerPoint JavaScript API 1.1 には、新しいプレゼンテーションを作成するための 1 つの API が含まれます。</span><span class="sxs-lookup"><span data-stu-id="b2eef-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="b2eef-125">API の詳細については、「[PowerPoint の JavaScript API](../../powerpoint/powerpoint-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b2eef-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="b2eef-126">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="b2eef-126">Runtime requirement support check</span></span>

<span data-ttu-id="b2eef-127">実行時に、アドインは次を行うことによって、特定のアプリケーションが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="b2eef-127">At runtime, add-ins can check if a particular application supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="b2eef-128">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="b2eef-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="b2eef-129">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="b2eef-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="b2eef-130">Office アプリケーションまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのアプリケーションまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="b2eef-130">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="b2eef-131">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office クライアント アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="b2eef-131">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b2eef-132">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="b2eef-132">Office Common API requirement sets</span></span>

<span data-ttu-id="b2eef-133">PowerPoint のほとんどのアドイン機能は、共通の API セットから取得されます。</span><span class="sxs-lookup"><span data-stu-id="b2eef-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="b2eef-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b2eef-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b2eef-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="b2eef-135">See also</span></span>

- [<span data-ttu-id="b2eef-136">PowerPoint JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="b2eef-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="b2eef-137">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="b2eef-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b2eef-138">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="b2eef-138">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="b2eef-139">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="b2eef-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
