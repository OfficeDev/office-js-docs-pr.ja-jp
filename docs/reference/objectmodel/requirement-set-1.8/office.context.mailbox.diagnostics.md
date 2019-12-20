---
title: Office.--の要件セット1.8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2c5ea33ffd8bc560288935f7ee65ebb93aadf1aa
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814251"
---
# <a name="diagnostics"></a><span data-ttu-id="cb986-102">診断</span><span class="sxs-lookup"><span data-stu-id="cb986-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="cb986-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="cb986-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="cb986-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="cb986-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb986-105">要件</span><span class="sxs-lookup"><span data-stu-id="cb986-105">Requirements</span></span>

|<span data-ttu-id="cb986-106">要件</span><span class="sxs-lookup"><span data-stu-id="cb986-106">Requirement</span></span>| <span data-ttu-id="cb986-107">値</span><span class="sxs-lookup"><span data-stu-id="cb986-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb986-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb986-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb986-109">1.1</span><span class="sxs-lookup"><span data-stu-id="cb986-109">1.1</span></span>|
|[<span data-ttu-id="cb986-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cb986-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb986-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb986-111">ReadItem</span></span>|
|[<span data-ttu-id="cb986-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb986-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb986-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb986-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="cb986-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="cb986-114">Properties</span></span>

| <span data-ttu-id="cb986-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="cb986-115">Property</span></span> | <span data-ttu-id="cb986-116">最小値</span><span class="sxs-lookup"><span data-stu-id="cb986-116">Minimum</span></span><br><span data-ttu-id="cb986-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cb986-117">permission level</span></span> | <span data-ttu-id="cb986-118">モード</span><span class="sxs-lookup"><span data-stu-id="cb986-118">Modes</span></span> | <span data-ttu-id="cb986-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="cb986-119">Return type</span></span> | <span data-ttu-id="cb986-120">最小値</span><span class="sxs-lookup"><span data-stu-id="cb986-120">Minimum</span></span><br><span data-ttu-id="cb986-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="cb986-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="cb986-122">名</span><span class="sxs-lookup"><span data-stu-id="cb986-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostname) | <span data-ttu-id="cb986-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb986-123">ReadItem</span></span> | <span data-ttu-id="cb986-124">作成</span><span class="sxs-lookup"><span data-stu-id="cb986-124">Compose</span></span><br><span data-ttu-id="cb986-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb986-125">Read</span></span> | <span data-ttu-id="cb986-126">String</span><span class="sxs-lookup"><span data-stu-id="cb986-126">String</span></span> | [<span data-ttu-id="cb986-127">1.1</span><span class="sxs-lookup"><span data-stu-id="cb986-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb986-128">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="cb986-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#hostversion) | <span data-ttu-id="cb986-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb986-129">ReadItem</span></span> | <span data-ttu-id="cb986-130">作成</span><span class="sxs-lookup"><span data-stu-id="cb986-130">Compose</span></span><br><span data-ttu-id="cb986-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb986-131">Read</span></span> | <span data-ttu-id="cb986-132">String</span><span class="sxs-lookup"><span data-stu-id="cb986-132">String</span></span> | [<span data-ttu-id="cb986-133">1.1</span><span class="sxs-lookup"><span data-stu-id="cb986-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb986-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="cb986-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8#owaview) | <span data-ttu-id="cb986-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb986-135">ReadItem</span></span> | <span data-ttu-id="cb986-136">作成</span><span class="sxs-lookup"><span data-stu-id="cb986-136">Compose</span></span><br><span data-ttu-id="cb986-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb986-137">Read</span></span> | <span data-ttu-id="cb986-138">String</span><span class="sxs-lookup"><span data-stu-id="cb986-138">String</span></span> | [<span data-ttu-id="cb986-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cb986-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
