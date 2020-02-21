---
title: マニフェスト ファイルの Permissions 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165546"
---
# <a name="permissions-element"></a><span data-ttu-id="c31d1-102">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="c31d1-102">Permissions element</span></span>

<span data-ttu-id="c31d1-103">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c31d1-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="c31d1-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="c31d1-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c31d1-105">構文</span><span class="sxs-lookup"><span data-stu-id="c31d1-105">Syntax</span></span>

<span data-ttu-id="c31d1-106">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="c31d1-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="c31d1-107">メール アドインの場合</span><span class="sxs-lookup"><span data-stu-id="c31d1-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="c31d1-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="c31d1-108">Contained in</span></span>

[<span data-ttu-id="c31d1-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c31d1-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="c31d1-110">注釈</span><span class="sxs-lookup"><span data-stu-id="c31d1-110">Remarks</span></span>

<span data-ttu-id="c31d1-111">詳細については、「[アドインで API を使用するためのアクセス許可を要求](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)する」と「 [Outlook アドインのアクセス許可につい](../../outlook/understanding-outlook-add-in-permissions.md)て」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c31d1-111">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
