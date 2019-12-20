---
title: Office.--の要件セット1.1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: a1787cb00b5d373c2051d40ccc219b05c8bea4af
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815026"
---
# <a name="diagnostics"></a><span data-ttu-id="33365-102">診断</span><span class="sxs-lookup"><span data-stu-id="33365-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="33365-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="33365-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="33365-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="33365-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="33365-105">要件</span><span class="sxs-lookup"><span data-stu-id="33365-105">Requirements</span></span>

|<span data-ttu-id="33365-106">要件</span><span class="sxs-lookup"><span data-stu-id="33365-106">Requirement</span></span>| <span data-ttu-id="33365-107">値</span><span class="sxs-lookup"><span data-stu-id="33365-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="33365-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="33365-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="33365-109">1.1</span><span class="sxs-lookup"><span data-stu-id="33365-109">1.1</span></span>|
|[<span data-ttu-id="33365-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33365-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="33365-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33365-111">ReadItem</span></span>|
|[<span data-ttu-id="33365-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="33365-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="33365-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="33365-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="33365-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="33365-114">Properties</span></span>

| <span data-ttu-id="33365-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="33365-115">Property</span></span> | <span data-ttu-id="33365-116">最小値</span><span class="sxs-lookup"><span data-stu-id="33365-116">Minimum</span></span><br><span data-ttu-id="33365-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="33365-117">permission level</span></span> | <span data-ttu-id="33365-118">モード</span><span class="sxs-lookup"><span data-stu-id="33365-118">Modes</span></span> | <span data-ttu-id="33365-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="33365-119">Return type</span></span> | <span data-ttu-id="33365-120">最小値</span><span class="sxs-lookup"><span data-stu-id="33365-120">Minimum</span></span><br><span data-ttu-id="33365-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="33365-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="33365-122">名</span><span class="sxs-lookup"><span data-stu-id="33365-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#hostname) | <span data-ttu-id="33365-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33365-123">ReadItem</span></span> | <span data-ttu-id="33365-124">作成</span><span class="sxs-lookup"><span data-stu-id="33365-124">Compose</span></span><br><span data-ttu-id="33365-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="33365-125">Read</span></span> | <span data-ttu-id="33365-126">String</span><span class="sxs-lookup"><span data-stu-id="33365-126">String</span></span> | [<span data-ttu-id="33365-127">1.1</span><span class="sxs-lookup"><span data-stu-id="33365-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="33365-128">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="33365-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#hostversion) | <span data-ttu-id="33365-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33365-129">ReadItem</span></span> | <span data-ttu-id="33365-130">作成</span><span class="sxs-lookup"><span data-stu-id="33365-130">Compose</span></span><br><span data-ttu-id="33365-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="33365-131">Read</span></span> | <span data-ttu-id="33365-132">String</span><span class="sxs-lookup"><span data-stu-id="33365-132">String</span></span> | [<span data-ttu-id="33365-133">1.1</span><span class="sxs-lookup"><span data-stu-id="33365-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="33365-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="33365-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.1#owaview) | <span data-ttu-id="33365-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="33365-135">ReadItem</span></span> | <span data-ttu-id="33365-136">作成</span><span class="sxs-lookup"><span data-stu-id="33365-136">Compose</span></span><br><span data-ttu-id="33365-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="33365-137">Read</span></span> | <span data-ttu-id="33365-138">String</span><span class="sxs-lookup"><span data-stu-id="33365-138">String</span></span> | [<span data-ttu-id="33365-139">1.1</span><span class="sxs-lookup"><span data-stu-id="33365-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
