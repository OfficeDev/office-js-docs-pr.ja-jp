---
title: Office.-mailbox-要件セット1.4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0532a9971a05412d37334f4c5a4b6b12654f61f3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950993"
---
# <a name="userprofile"></a><span data-ttu-id="a91c2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a91c2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a91c2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a91c2-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).userProfile</span></span>

<span data-ttu-id="a91c2-104">Outlook アドインのユーザーに関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="a91c2-104">Provides information about the user in an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a91c2-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="a91c2-105">Requirements</span></span>

|<span data-ttu-id="a91c2-106">要件</span><span class="sxs-lookup"><span data-stu-id="a91c2-106">Requirement</span></span>| <span data-ttu-id="a91c2-107">値</span><span class="sxs-lookup"><span data-stu-id="a91c2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a91c2-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a91c2-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a91c2-109">1.1</span><span class="sxs-lookup"><span data-stu-id="a91c2-109">1.1</span></span>|
|[<span data-ttu-id="a91c2-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a91c2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a91c2-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a91c2-111">ReadItem</span></span>|
|[<span data-ttu-id="a91c2-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a91c2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a91c2-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a91c2-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="a91c2-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="a91c2-114">Properties</span></span>

| <span data-ttu-id="a91c2-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="a91c2-115">Property</span></span> | <span data-ttu-id="a91c2-116">最小値</span><span class="sxs-lookup"><span data-stu-id="a91c2-116">Minimum</span></span><br><span data-ttu-id="a91c2-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a91c2-117">permission level</span></span> | <span data-ttu-id="a91c2-118">モード</span><span class="sxs-lookup"><span data-stu-id="a91c2-118">Modes</span></span> | <span data-ttu-id="a91c2-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="a91c2-119">Return type</span></span> | <span data-ttu-id="a91c2-120">最小値</span><span class="sxs-lookup"><span data-stu-id="a91c2-120">Minimum</span></span><br><span data-ttu-id="a91c2-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="a91c2-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="a91c2-122">displayName</span><span class="sxs-lookup"><span data-stu-id="a91c2-122">displayName</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#displayname) | <span data-ttu-id="a91c2-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a91c2-123">ReadItem</span></span> | <span data-ttu-id="a91c2-124">作成</span><span class="sxs-lookup"><span data-stu-id="a91c2-124">Compose</span></span><br><span data-ttu-id="a91c2-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="a91c2-125">Read</span></span> | <span data-ttu-id="a91c2-126">文字列</span><span class="sxs-lookup"><span data-stu-id="a91c2-126">String</span></span> | [<span data-ttu-id="a91c2-127">1.1</span><span class="sxs-lookup"><span data-stu-id="a91c2-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a91c2-128">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a91c2-128">emailAddress</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#emailaddress) | <span data-ttu-id="a91c2-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a91c2-129">ReadItem</span></span> | <span data-ttu-id="a91c2-130">作成</span><span class="sxs-lookup"><span data-stu-id="a91c2-130">Compose</span></span><br><span data-ttu-id="a91c2-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="a91c2-131">Read</span></span> | <span data-ttu-id="a91c2-132">文字列</span><span class="sxs-lookup"><span data-stu-id="a91c2-132">String</span></span> | [<span data-ttu-id="a91c2-133">1.1</span><span class="sxs-lookup"><span data-stu-id="a91c2-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a91c2-134">timeZone</span><span class="sxs-lookup"><span data-stu-id="a91c2-134">timeZone</span></span>](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5#timezone) | <span data-ttu-id="a91c2-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a91c2-135">ReadItem</span></span> | <span data-ttu-id="a91c2-136">作成</span><span class="sxs-lookup"><span data-stu-id="a91c2-136">Compose</span></span><br><span data-ttu-id="a91c2-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="a91c2-137">Read</span></span> | <span data-ttu-id="a91c2-138">文字列</span><span class="sxs-lookup"><span data-stu-id="a91c2-138">String</span></span> | [<span data-ttu-id="a91c2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a91c2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
