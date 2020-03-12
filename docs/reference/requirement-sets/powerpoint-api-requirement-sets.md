---
title: PowerPoint JavaScript API の要件セット
description: ''
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: ef76077c3a2a975fae8a0dc101e8e1b42ef66094
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600698"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="04f2a-102">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="04f2a-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="04f2a-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="04f2a-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="04f2a-106">次の表は、PowerPoint の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="04f2a-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="04f2a-107">要件セット</span><span class="sxs-lookup"><span data-stu-id="04f2a-107">Requirement set</span></span>  |  <span data-ttu-id="04f2a-108">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="04f2a-108">Office on Windows</span></span><br><span data-ttu-id="04f2a-109">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="04f2a-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="04f2a-110">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="04f2a-110">Office on iPad</span></span><br><span data-ttu-id="04f2a-111">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="04f2a-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="04f2a-112">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="04f2a-112">Office on Mac</span></span><br><span data-ttu-id="04f2a-113">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="04f2a-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="04f2a-114">Office on the web</span><span class="sxs-lookup"><span data-stu-id="04f2a-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="04f2a-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="04f2a-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="04f2a-116">バージョン 1810 (ビルド 11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="04f2a-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="04f2a-117">2.17 以降</span><span class="sxs-lookup"><span data-stu-id="04f2a-117">2.17 or later</span></span> | <span data-ttu-id="04f2a-118">16.19 以降</span><span class="sxs-lookup"><span data-stu-id="04f2a-118">16.19 or later</span></span> | <span data-ttu-id="04f2a-119">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="04f2a-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="04f2a-120">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="04f2a-120">Office versions and build numbers</span></span>

<span data-ttu-id="04f2a-121">Office のバージョンとビルド番号の詳細については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="04f2a-121">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="04f2a-122">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="04f2a-122">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="04f2a-123">PowerPoint JavaScript API 1.1 には、新しいプレゼンテーションを作成するための 1 つの API が含まれます。</span><span class="sxs-lookup"><span data-stu-id="04f2a-123">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="04f2a-124">API の詳細については、「[PowerPoint の JavaScript API](../../powerpoint/powerpoint-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="04f2a-124">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="04f2a-125">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="04f2a-125">Runtime requirement support check</span></span>

<span data-ttu-id="04f2a-126">実行時に、アドインは次を行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="04f2a-126">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="04f2a-127">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="04f2a-127">Manifest-based requirement support check</span></span>

<span data-ttu-id="04f2a-128">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで `Requirements` 要素を使用します。</span><span class="sxs-lookup"><span data-stu-id="04f2a-128">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="04f2a-129">Office ホストまたはプラットフォームが、`Requirements` 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="04f2a-129">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="04f2a-130">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="04f2a-130">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="04f2a-131">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="04f2a-131">Office Common API requirement sets</span></span>

<span data-ttu-id="04f2a-132">PowerPoint のほとんどのアドイン機能は、共通の API セットから取得されます。</span><span class="sxs-lookup"><span data-stu-id="04f2a-132">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="04f2a-133">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="04f2a-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="04f2a-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="04f2a-134">See also</span></span>

- [<span data-ttu-id="04f2a-135">PowerPoint JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="04f2a-135">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="04f2a-136">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="04f2a-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="04f2a-137">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="04f2a-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="04f2a-138">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="04f2a-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
