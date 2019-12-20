---
title: Office.--の要件セット1.3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 99658c1829e6021a79f72dcbeaff10b65d0a6dc7
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814307"
---
# <a name="diagnostics"></a><span data-ttu-id="7f761-102">診断</span><span class="sxs-lookup"><span data-stu-id="7f761-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="7f761-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="7f761-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="7f761-104">Outlook アドインに診断情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="7f761-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f761-105">要件</span><span class="sxs-lookup"><span data-stu-id="7f761-105">Requirements</span></span>

|<span data-ttu-id="7f761-106">要件</span><span class="sxs-lookup"><span data-stu-id="7f761-106">Requirement</span></span>| <span data-ttu-id="7f761-107">値</span><span class="sxs-lookup"><span data-stu-id="7f761-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f761-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7f761-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7f761-109">1.1</span><span class="sxs-lookup"><span data-stu-id="7f761-109">1.1</span></span>|
|[<span data-ttu-id="7f761-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7f761-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7f761-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f761-111">ReadItem</span></span>|
|[<span data-ttu-id="7f761-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7f761-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f761-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7f761-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="7f761-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7f761-114">Properties</span></span>

| <span data-ttu-id="7f761-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7f761-115">Property</span></span> | <span data-ttu-id="7f761-116">最小値</span><span class="sxs-lookup"><span data-stu-id="7f761-116">Minimum</span></span><br><span data-ttu-id="7f761-117">アクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7f761-117">permission level</span></span> | <span data-ttu-id="7f761-118">モード</span><span class="sxs-lookup"><span data-stu-id="7f761-118">Modes</span></span> | <span data-ttu-id="7f761-119">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="7f761-119">Return type</span></span> | <span data-ttu-id="7f761-120">最小値</span><span class="sxs-lookup"><span data-stu-id="7f761-120">Minimum</span></span><br><span data-ttu-id="7f761-121">要件セット</span><span class="sxs-lookup"><span data-stu-id="7f761-121">requirement set</span></span> |
|---|---|---|---|:---:|
| [<span data-ttu-id="7f761-122">名</span><span class="sxs-lookup"><span data-stu-id="7f761-122">hostName</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3#hostname) | <span data-ttu-id="7f761-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f761-123">ReadItem</span></span> | <span data-ttu-id="7f761-124">作成</span><span class="sxs-lookup"><span data-stu-id="7f761-124">Compose</span></span><br><span data-ttu-id="7f761-125">読み取り</span><span class="sxs-lookup"><span data-stu-id="7f761-125">Read</span></span> | <span data-ttu-id="7f761-126">String</span><span class="sxs-lookup"><span data-stu-id="7f761-126">String</span></span> | [<span data-ttu-id="7f761-127">1.1</span><span class="sxs-lookup"><span data-stu-id="7f761-127">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7f761-128">上 diagnostics.hostversion</span><span class="sxs-lookup"><span data-stu-id="7f761-128">hostVersion</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3#hostversion) | <span data-ttu-id="7f761-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f761-129">ReadItem</span></span> | <span data-ttu-id="7f761-130">作成</span><span class="sxs-lookup"><span data-stu-id="7f761-130">Compose</span></span><br><span data-ttu-id="7f761-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="7f761-131">Read</span></span> | <span data-ttu-id="7f761-132">String</span><span class="sxs-lookup"><span data-stu-id="7f761-132">String</span></span> | [<span data-ttu-id="7f761-133">1.1</span><span class="sxs-lookup"><span data-stu-id="7f761-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7f761-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="7f761-134">OWAView</span></span>](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.3#owaview) | <span data-ttu-id="7f761-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7f761-135">ReadItem</span></span> | <span data-ttu-id="7f761-136">作成</span><span class="sxs-lookup"><span data-stu-id="7f761-136">Compose</span></span><br><span data-ttu-id="7f761-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="7f761-137">Read</span></span> | <span data-ttu-id="7f761-138">String</span><span class="sxs-lookup"><span data-stu-id="7f761-138">String</span></span> | [<span data-ttu-id="7f761-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7f761-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
