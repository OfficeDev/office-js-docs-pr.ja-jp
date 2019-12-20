---
title: Office.-mailbox-要件セット1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 4d63dfe1b32de2ac7fe55f324f938b85a865ec02
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814914"
---
# <a name="userprofile"></a><span data-ttu-id="39003-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="39003-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="39003-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="39003-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="39003-104">Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="39003-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="39003-105">要件</span><span class="sxs-lookup"><span data-stu-id="39003-105">Requirements</span></span>

|<span data-ttu-id="39003-106">要件</span><span class="sxs-lookup"><span data-stu-id="39003-106">Requirement</span></span>| <span data-ttu-id="39003-107">値</span><span class="sxs-lookup"><span data-stu-id="39003-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="39003-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="39003-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="39003-109">1.1</span><span class="sxs-lookup"><span data-stu-id="39003-109">1.1</span></span>|
|[<span data-ttu-id="39003-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39003-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="39003-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39003-111">ReadItem</span></span>|
|[<span data-ttu-id="39003-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="39003-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="39003-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="39003-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="39003-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="39003-114">Properties</span></span>

| <span data-ttu-id="39003-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="39003-115">Property</span></span> | <span data-ttu-id="39003-116">最小値</span><span class="sxs-lookup"><span data-stu-id="39003-116">Minimum</span></span><br><span data-ttu-id="39003-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="39003-117">permission level</span></span> | <span data-ttu-id="39003-118">モード</span><span class="sxs-lookup"><span data-stu-id="39003-118">Modes</span></span> | <span data-ttu-id="39003-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="39003-119">Return type</span></span> | <span data-ttu-id="39003-120">最小値</span><span class="sxs-lookup"><span data-stu-id="39003-120">Minimum</span></span><br><span data-ttu-id="39003-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="39003-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="39003-122">displayName</span><span class="sxs-lookup"><span data-stu-id="39003-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#displayname) | <span data-ttu-id="39003-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39003-123">ReadItem</span></span> | <span data-ttu-id="39003-124">作成</span><span class="sxs-lookup"><span data-stu-id="39003-124">Compose</span></span><br><span data-ttu-id="39003-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="39003-125">Read</span></span> | <span data-ttu-id="39003-126">String</span><span class="sxs-lookup"><span data-stu-id="39003-126">String</span></span> | [<span data-ttu-id="39003-127">1.1</span><span class="sxs-lookup"><span data-stu-id="39003-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="39003-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="39003-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#emailaddress) | <span data-ttu-id="39003-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39003-129">ReadItem</span></span> | <span data-ttu-id="39003-130">作成</span><span class="sxs-lookup"><span data-stu-id="39003-130">Compose</span></span><br><span data-ttu-id="39003-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="39003-131">Read</span></span> | <span data-ttu-id="39003-132">String</span><span class="sxs-lookup"><span data-stu-id="39003-132">String</span></span> | [<span data-ttu-id="39003-133">1.1</span><span class="sxs-lookup"><span data-stu-id="39003-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="39003-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="39003-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.3#timezone) | <span data-ttu-id="39003-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="39003-135">ReadItem</span></span> | <span data-ttu-id="39003-136">作成</span><span class="sxs-lookup"><span data-stu-id="39003-136">Compose</span></span><br><span data-ttu-id="39003-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="39003-137">Read</span></span> | <span data-ttu-id="39003-138">String</span><span class="sxs-lookup"><span data-stu-id="39003-138">String</span></span> | [<span data-ttu-id="39003-139">1.1</span><span class="sxs-lookup"><span data-stu-id="39003-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
