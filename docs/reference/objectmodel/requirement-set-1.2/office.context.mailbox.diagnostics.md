---
title: Office.context.mailbox.diagnostics - 要件セット 1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dad9d35c397351938944d89bf98e450427cb74a3
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814984"
---
# <a name="diagnostics"></a><span data-ttu-id="d1f8e-102">診断</span><span class="sxs-lookup"><span data-stu-id="d1f8e-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="d1f8e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="d1f8e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="d1f8e-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="d1f8e-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1f8e-105">要件</span><span class="sxs-lookup"><span data-stu-id="d1f8e-105">Requirements</span></span>

|<span data-ttu-id="d1f8e-106">要件</span><span class="sxs-lookup"><span data-stu-id="d1f8e-106">Requirement</span></span>| <span data-ttu-id="d1f8e-107">値</span><span class="sxs-lookup"><span data-stu-id="d1f8e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1f8e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1f8e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d1f8e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="d1f8e-109">1.1</span></span>|
|[<span data-ttu-id="d1f8e-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1f8e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1f8e-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1f8e-111">ReadItem</span></span>|
|[<span data-ttu-id="d1f8e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1f8e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1f8e-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1f8e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="d1f8e-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="d1f8e-114">Properties</span></span>

| <span data-ttu-id="d1f8e-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="d1f8e-115">Property</span></span> | <span data-ttu-id="d1f8e-116">最小値</span><span class="sxs-lookup"><span data-stu-id="d1f8e-116">Minimum</span></span><br><span data-ttu-id="d1f8e-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1f8e-117">permission level</span></span> | <span data-ttu-id="d1f8e-118">モード</span><span class="sxs-lookup"><span data-stu-id="d1f8e-118">Modes</span></span> | <span data-ttu-id="d1f8e-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="d1f8e-119">Return type</span></span> | <span data-ttu-id="d1f8e-120">最小値</span><span class="sxs-lookup"><span data-stu-id="d1f8e-120">Minimum</span></span><br><span data-ttu-id="d1f8e-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="d1f8e-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="d1f8e-122">名</span><span class="sxs-lookup"><span data-stu-id="d1f8e-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostname) | <span data-ttu-id="d1f8e-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1f8e-123">ReadItem</span></span> | <span data-ttu-id="d1f8e-124">作成</span><span class="sxs-lookup"><span data-stu-id="d1f8e-124">Compose</span></span><br><span data-ttu-id="d1f8e-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1f8e-125">Read</span></span> | <span data-ttu-id="d1f8e-126">String</span><span class="sxs-lookup"><span data-stu-id="d1f8e-126">String</span></span> | [<span data-ttu-id="d1f8e-127">1.1</span><span class="sxs-lookup"><span data-stu-id="d1f8e-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d1f8e-128">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="d1f8e-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#hostversion) | <span data-ttu-id="d1f8e-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1f8e-129">ReadItem</span></span> | <span data-ttu-id="d1f8e-130">作成</span><span class="sxs-lookup"><span data-stu-id="d1f8e-130">Compose</span></span><br><span data-ttu-id="d1f8e-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1f8e-131">Read</span></span> | <span data-ttu-id="d1f8e-132">String</span><span class="sxs-lookup"><span data-stu-id="d1f8e-132">String</span></span> | [<span data-ttu-id="d1f8e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="d1f8e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d1f8e-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="d1f8e-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.2#owaview) | <span data-ttu-id="d1f8e-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1f8e-135">ReadItem</span></span> | <span data-ttu-id="d1f8e-136">作成</span><span class="sxs-lookup"><span data-stu-id="d1f8e-136">Compose</span></span><br><span data-ttu-id="d1f8e-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1f8e-137">Read</span></span> | <span data-ttu-id="d1f8e-138">String</span><span class="sxs-lookup"><span data-stu-id="d1f8e-138">String</span></span> | [<span data-ttu-id="d1f8e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d1f8e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
