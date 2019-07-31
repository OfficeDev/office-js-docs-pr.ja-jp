---
title: PowerPoint JavaScript API の要件セット
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941149"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="8be64-102">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="8be64-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="8be64-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8be64-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="8be64-106">次の表に、PowerPoint の要件セット、それらの要件セットをサポートする Office ホストアプリケーション、ビルドバージョンまたは使用可能な日付を示します。</span><span class="sxs-lookup"><span data-stu-id="8be64-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="8be64-107">要件セット</span><span class="sxs-lookup"><span data-stu-id="8be64-107">Requirement set</span></span>  |  <span data-ttu-id="8be64-108">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="8be64-108">Office on Windows</span></span><br><span data-ttu-id="8be64-109">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="8be64-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="8be64-110">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="8be64-110">Office on iPad</span></span><br><span data-ttu-id="8be64-111">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="8be64-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="8be64-112">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="8be64-112">Office on Mac</span></span><br><span data-ttu-id="8be64-113">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="8be64-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="8be64-114">Web 上の Office</span><span class="sxs-lookup"><span data-stu-id="8be64-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="8be64-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="8be64-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="8be64-116">バージョン 1810 (ビルド 11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="8be64-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="8be64-117">2.17 以降</span><span class="sxs-lookup"><span data-stu-id="8be64-117">2.17 or later</span></span> | <span data-ttu-id="8be64-118">16.19 以降</span><span class="sxs-lookup"><span data-stu-id="8be64-118">16.19 or later</span></span> | <span data-ttu-id="8be64-119">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="8be64-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="8be64-120">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="8be64-120">Office versions and build numbers</span></span>

<span data-ttu-id="8be64-121">Office のバージョンとビルド番号の詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8be64-121">For more information about Office versions and build numbers, see:</span></span>

- [<span data-ttu-id="8be64-122">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="8be64-122">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="8be64-123">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="8be64-123">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="8be64-124">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="8be64-124">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="8be64-125">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="8be64-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="8be64-126">PowerPoint JavaScript API 1.1 には、新しいプレゼンテーションを作成するための単一の API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8be64-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="8be64-127">API の詳細については、「 [JAVASCRIPT api For PowerPoint](../../powerpoint/powerpoint-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8be64-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="8be64-128">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="8be64-128">Runtime requirement support check</span></span>

<span data-ttu-id="8be64-129">実行時に、アドインは、次の手順に従って、特定のホストが API 要件セットをサポートしているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="8be64-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="8be64-130">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="8be64-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="8be64-131">アドインマニフェスト`Requirements`の要素を使用して、アドインが使用する必要がある重要な要件セットまたは API メンバーを指定します。</span><span class="sxs-lookup"><span data-stu-id="8be64-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="8be64-132">Office ホストまたはプラットフォームが、 `Requirements`要素で指定されている要件セットや API メンバーをサポートしていない場合、アドインはそのホストまたはプラットフォームでは実行されず、アドインには表示されません。</span><span class="sxs-lookup"><span data-stu-id="8be64-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="8be64-133">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="8be64-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="8be64-134">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="8be64-134">Office Common API requirement sets</span></span>

<span data-ttu-id="8be64-135">PowerPoint アドインのほとんどの機能は、共通 API セットから取得されます。</span><span class="sxs-lookup"><span data-stu-id="8be64-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="8be64-136">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8be64-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8be64-137">関連項目</span><span class="sxs-lookup"><span data-stu-id="8be64-137">See also</span></span>

- [<span data-ttu-id="8be64-138">PowerPoint JavaScript API リファレンスドキュメント</span><span class="sxs-lookup"><span data-stu-id="8be64-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="8be64-139">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="8be64-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="8be64-140">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="8be64-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="8be64-141">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="8be64-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
