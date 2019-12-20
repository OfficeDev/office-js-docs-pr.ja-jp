---
title: Office.-mailbox-要件セット1.5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 6b5229c1bc300d11714f3aa2cf8fa8ff2465667c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814781"
---
# <a name="userprofile"></a><span data-ttu-id="31663-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="31663-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="31663-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="31663-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="31663-104">Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="31663-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="31663-105">要件</span><span class="sxs-lookup"><span data-stu-id="31663-105">Requirements</span></span>

|<span data-ttu-id="31663-106">要件</span><span class="sxs-lookup"><span data-stu-id="31663-106">Requirement</span></span>| <span data-ttu-id="31663-107">値</span><span class="sxs-lookup"><span data-stu-id="31663-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="31663-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="31663-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="31663-109">1.1</span><span class="sxs-lookup"><span data-stu-id="31663-109">1.1</span></span>|
|[<span data-ttu-id="31663-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="31663-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="31663-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31663-111">ReadItem</span></span>|
|[<span data-ttu-id="31663-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="31663-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="31663-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="31663-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="31663-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="31663-114">Properties</span></span>

| <span data-ttu-id="31663-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="31663-115">Property</span></span> | <span data-ttu-id="31663-116">最小値</span><span class="sxs-lookup"><span data-stu-id="31663-116">Minimum</span></span><br><span data-ttu-id="31663-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="31663-117">permission level</span></span> | <span data-ttu-id="31663-118">モード</span><span class="sxs-lookup"><span data-stu-id="31663-118">Modes</span></span> | <span data-ttu-id="31663-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="31663-119">Return type</span></span> | <span data-ttu-id="31663-120">最小値</span><span class="sxs-lookup"><span data-stu-id="31663-120">Minimum</span></span><br><span data-ttu-id="31663-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="31663-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="31663-122">displayName</span><span class="sxs-lookup"><span data-stu-id="31663-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="31663-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31663-123">ReadItem</span></span> | <span data-ttu-id="31663-124">作成</span><span class="sxs-lookup"><span data-stu-id="31663-124">Compose</span></span><br><span data-ttu-id="31663-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="31663-125">Read</span></span> | <span data-ttu-id="31663-126">String</span><span class="sxs-lookup"><span data-stu-id="31663-126">String</span></span> | [<span data-ttu-id="31663-127">1.1</span><span class="sxs-lookup"><span data-stu-id="31663-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="31663-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="31663-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="31663-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31663-129">ReadItem</span></span> | <span data-ttu-id="31663-130">作成</span><span class="sxs-lookup"><span data-stu-id="31663-130">Compose</span></span><br><span data-ttu-id="31663-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="31663-131">Read</span></span> | <span data-ttu-id="31663-132">String</span><span class="sxs-lookup"><span data-stu-id="31663-132">String</span></span> | [<span data-ttu-id="31663-133">1.1</span><span class="sxs-lookup"><span data-stu-id="31663-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="31663-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="31663-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="31663-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="31663-135">ReadItem</span></span> | <span data-ttu-id="31663-136">作成</span><span class="sxs-lookup"><span data-stu-id="31663-136">Compose</span></span><br><span data-ttu-id="31663-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="31663-137">Read</span></span> | <span data-ttu-id="31663-138">String</span><span class="sxs-lookup"><span data-stu-id="31663-138">String</span></span> | [<span data-ttu-id="31663-139">1.1</span><span class="sxs-lookup"><span data-stu-id="31663-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
