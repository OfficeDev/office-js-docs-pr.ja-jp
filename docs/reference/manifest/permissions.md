---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、Office アドインの API アクセスレベルを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 603494b61ef126b35cb5cdff8c5f5b911bd25840
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611492"
---
# <a name="permissions-element"></a><span data-ttu-id="cbd40-103">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="cbd40-103">Permissions element</span></span>

<span data-ttu-id="cbd40-104">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbd40-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="cbd40-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="cbd40-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cbd40-106">構文</span><span class="sxs-lookup"><span data-stu-id="cbd40-106">Syntax</span></span>

<span data-ttu-id="cbd40-107">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="cbd40-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="cbd40-108">メール アドインの場合</span><span class="sxs-lookup"><span data-stu-id="cbd40-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="cbd40-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="cbd40-109">Contained in</span></span>

[<span data-ttu-id="cbd40-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cbd40-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="cbd40-111">注釈</span><span class="sxs-lookup"><span data-stu-id="cbd40-111">Remarks</span></span>

<span data-ttu-id="cbd40-112">詳細については、「[アドインで API を使用するためのアクセス許可を要求](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)する」と「 [Outlook アドインのアクセス許可につい](../../outlook/understanding-outlook-add-in-permissions.md)て」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cbd40-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
