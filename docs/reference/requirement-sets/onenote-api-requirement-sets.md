---
title: OneNote JavaScript API の要件セット
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: c0cec9375dc0d3db4a09e44a9b73935af0691bde
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433825"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="fb7a4-102">OneNote JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="fb7a4-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="fb7a4-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="fb7a4-106">次の表は、OneNote の要件セット、それらの要件セットをサポートする Office ホスト アプリケーション、ビルド バージョンまたは一般提供開始日の一覧です。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="fb7a4-107">要件セット</span><span class="sxs-lookup"><span data-stu-id="fb7a4-107">Requirement set</span></span>  |  <span data-ttu-id="fb7a4-108">Office Online</span><span class="sxs-lookup"><span data-stu-id="fb7a4-108">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="fb7a4-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="fb7a4-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="fb7a4-110">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="fb7a4-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="fb7a4-111">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="fb7a4-111">Office common API requirement sets</span></span>

<span data-ttu-id="fb7a4-112">共通 API の要件セットについて詳しくは、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-112">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="fb7a4-113">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="fb7a4-113">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="fb7a4-114">OneNote JavaScript API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="fb7a4-115">API について詳しくは、「[OneNote の JavaScript API のプログラミングの概要](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-115">For details about the API, see the [OneNote JavaScript API](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview) reference topics.</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="fb7a4-116">ランタイム要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="fb7a4-116">Runtime requirement support check</span></span>

<span data-ttu-id="fb7a4-117">実行時に、アドインは次のチェックを行うことによって、特定のホストが API 要件をサポートしているかどうかをチェックできます。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-117">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="fb7a4-118">マニフェストに基づく要件のサポートのチェック</span><span class="sxs-lookup"><span data-stu-id="fb7a4-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="fb7a4-p103">アドインで必須の、重要な要件セットまたは API メンバーを指定するには、アドインのマニフェストで Requirements 要素を使用します。Office ホストまたはプラットフォームが、Requirements 要素で指定した要件セットまたは API メンバーをサポートしない場合、アドインはそのホストまたはプラットフォームでは実行されず、[個人用アドイン] にも表示されません。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="fb7a4-121">OneNoteApi 要件セット、バージョン 1.1 をサポートするすべての Office ホスト アプリケーションで読み込まれるアドインのコード例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="fb7a4-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="fb7a4-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="fb7a4-122">See also</span></span>

- [<span data-ttu-id="fb7a4-123">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="fb7a4-123">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="fb7a4-124">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="fb7a4-124">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="fb7a4-125">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="fb7a4-125">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
