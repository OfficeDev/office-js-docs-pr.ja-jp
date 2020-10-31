---
title: PowerPoint JavaScript API の要件セット
description: PowerPoint JavaScript API の要件セットの詳細情報。
ms.date: 10/26/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: cf9ab510e4b35a140c77ee958279cb85a2189fa2
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774729"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="58d13-103">PowerPoint JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="58d13-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="58d13-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="58d13-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="58d13-107">次の表は、PowerPoint の要件セット、それらの要件セットをサポートする Office クライアント アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="58d13-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="58d13-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="58d13-108">Requirement set</span></span>  |  <span data-ttu-id="58d13-109">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="58d13-109">Office on Windows</span></span><br><span data-ttu-id="58d13-110">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="58d13-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="58d13-111">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="58d13-111">Office on iPad</span></span><br><span data-ttu-id="58d13-112">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="58d13-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="58d13-113">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="58d13-113">Office on Mac</span></span><br><span data-ttu-id="58d13-114">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="58d13-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="58d13-115">Office on the web</span><span class="sxs-lookup"><span data-stu-id="58d13-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| [<span data-ttu-id="58d13-116">プレビュー</span><span class="sxs-lookup"><span data-stu-id="58d13-116">Preview</span></span>](powerpoint-preview-apis.md)  | <span data-ttu-id="58d13-117">プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://insider.office.com)に参加する必要があります)。</span><span class="sxs-lookup"><span data-stu-id="58d13-117">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com)).</span></span> |
| <span data-ttu-id="58d13-118">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="58d13-118">PowerPointApi 1.1</span></span> | <span data-ttu-id="58d13-119">バージョン 1810 (ビルド 11001.20074) 以降</span><span class="sxs-lookup"><span data-stu-id="58d13-119">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="58d13-120">2.17 以降</span><span class="sxs-lookup"><span data-stu-id="58d13-120">2.17 or later</span></span> | <span data-ttu-id="58d13-121">16.19 以降</span><span class="sxs-lookup"><span data-stu-id="58d13-121">16.19 or later</span></span> | <span data-ttu-id="58d13-122">2018 年 10 月</span><span class="sxs-lookup"><span data-stu-id="58d13-122">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="58d13-123">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="58d13-123">Office versions and build numbers</span></span>

<span data-ttu-id="58d13-124">Office のバージョンとビルド番号の詳細については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58d13-124">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="58d13-125">PowerPoint JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="58d13-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="58d13-126">PowerPoint JavaScript API 1.1 には、[新しいプレゼンテーションを作成するための 1 つの API](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-) が含まれます。</span><span class="sxs-lookup"><span data-stu-id="58d13-126">PowerPoint JavaScript API 1.1 contains a [single API to create a new presentation](/javascript/api/powerpoint#powerpoint-createpresentation-base64file-).</span></span> <span data-ttu-id="58d13-127">API の詳細については、「[プレゼンテーションを作成する](../../powerpoint/powerpoint-add-ins.md#create-a-presentation)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58d13-127">For details about the API, see [Create a presentation](../../powerpoint/powerpoint-add-ins.md#create-a-presentation).</span></span>

## <a name="how-to-use-powerpoint-requirement-sets-at-runtime-and-in-the-manifest"></a><span data-ttu-id="58d13-128">実行時およびマニフェストで PowerPoint 要件セットを使用する方法</span><span class="sxs-lookup"><span data-stu-id="58d13-128">How to use PowerPoint requirement sets at runtime and in the manifest</span></span>

> [!NOTE]
> <span data-ttu-id="58d13-129">このセクションでは、[Office バージョンと要件セット](../../develop/office-versions-and-requirement-sets.md) の概要、および [Office アプリケーションと API 要件の指定](../../develop/specify-office-hosts-and-api-requirements.md) について理解していることを前提としています。</span><span class="sxs-lookup"><span data-stu-id="58d13-129">This section assumes you're familiar with the overview of requirement sets at [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md) and [Specify Office applications and API requirements](../../develop/specify-office-hosts-and-api-requirements.md).</span></span>

<span data-ttu-id="58d13-130">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="58d13-130">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="58d13-131">Office アドインはランタイム チェックを実行できます。または、マニフェストで指定されている要件セットを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="58d13-131">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="58d13-132">実行時に要件セットのサポートを確認する</span><span class="sxs-lookup"><span data-stu-id="58d13-132">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="58d13-133">次のコード サンプルは、アドインが実行されている Office アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="58d13-133">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
} else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="58d13-134">マニフェストで要件セットのサポートを定義する</span><span class="sxs-lookup"><span data-stu-id="58d13-134">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="58d13-135">アドインのマニフェストで [Requirements 要素](../manifest/requirements.md) を使用して、アドインをアクティブにするために必要な最小要件セットや API メソッド (またはその両方) を指定できます。</span><span class="sxs-lookup"><span data-stu-id="58d13-135">You can use the [Requirements element](../manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="58d13-136">Office アプリケーションまたはプラットフォームが、マニフェストの `Requirements` 要素で指定されている要件セットまたは API メソッドをサポートしていない場合、アドインはそのアプリケーションまたはプラットフォームで実行されず、 **[マイ アドイン]** に表示されるアドインのリストに表示されません。アドインが完全な機能のために特定の要件セットを必要とするが、要件セットをサポートしていないプラットフォームのユーザーにも価値を提供できる場合は、マニフェストの要件セットのサポートを定義する代わりに、上記のように実行時に要件サポートを確認することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="58d13-136">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins** . If your add-in requires a specific requirement set for full functionality, but it can provide value even to users on platforms that do not support the requirement set, we recommend that you check for requirement support at runtime as described above, instead of defining requirement set support in the manifest.</span></span>

<span data-ttu-id="58d13-137">次のコード サンプルは、アドインが PowerPointApi 要件セットのバージョン 1.1 以上をサポートする Office クライアント アプリケーションのすべてで読み込まれる必要があることを指定する、アドインのマニフェストの `Requirements` 要素を示しています。</span><span class="sxs-lookup"><span data-stu-id="58d13-137">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support PowerPointApi requirement set version 1.1 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="58d13-138">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="58d13-138">Office Common API requirement sets</span></span>

<span data-ttu-id="58d13-139">PowerPoint のほとんどのアドイン機能は、共通の API セットから取得されます。</span><span class="sxs-lookup"><span data-stu-id="58d13-139">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="58d13-140">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="58d13-140">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="58d13-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="58d13-141">See also</span></span>

- [<span data-ttu-id="58d13-142">PowerPoint JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="58d13-142">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="58d13-143">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="58d13-143">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="58d13-144">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="58d13-144">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="58d13-145">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="58d13-145">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
