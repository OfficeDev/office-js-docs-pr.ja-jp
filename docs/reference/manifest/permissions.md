---
title: マニフェスト ファイルの Permissions 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851307"
---
# <a name="permissions-element"></a><span data-ttu-id="b58a0-102">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="b58a0-102">Permissions element</span></span>

<span data-ttu-id="b58a0-103">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b58a0-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="b58a0-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="b58a0-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b58a0-105">構文</span><span class="sxs-lookup"><span data-stu-id="b58a0-105">Syntax</span></span>

<span data-ttu-id="b58a0-106">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="b58a0-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="b58a0-107">メール アドインの場合</span><span class="sxs-lookup"><span data-stu-id="b58a0-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="b58a0-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b58a0-108">Contained in</span></span>

[<span data-ttu-id="b58a0-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b58a0-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b58a0-110">注釈</span><span class="sxs-lookup"><span data-stu-id="b58a0-110">Remarks</span></span>

<span data-ttu-id="b58a0-111">詳細については、「[アドインで API を使用するためのアクセス許可を要求](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)する」と「 [Outlook アドインのアクセス許可につい](/outlook/add-ins/understanding-outlook-add-in-permissions)て」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b58a0-111">For more detail, see [Requesting permissions for API use in add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
