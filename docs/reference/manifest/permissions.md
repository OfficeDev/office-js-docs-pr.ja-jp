---
title: マニフェスト ファイルの Permissions 要素
description: Permissions 要素は、Office アドインの API アクセスレベルを指定します。
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006459"
---
# <a name="permissions-element"></a><span data-ttu-id="6c25b-103">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="6c25b-103">Permissions element</span></span>

<span data-ttu-id="6c25b-104">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6c25b-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="6c25b-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="6c25b-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6c25b-106">構文</span><span class="sxs-lookup"><span data-stu-id="6c25b-106">Syntax</span></span>

<span data-ttu-id="6c25b-107">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="6c25b-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="6c25b-108">メール アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="6c25b-108">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="6c25b-109">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="6c25b-109">Contained in</span></span>

[<span data-ttu-id="6c25b-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="6c25b-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="6c25b-111">注釈</span><span class="sxs-lookup"><span data-stu-id="6c25b-111">Remarks</span></span>

<span data-ttu-id="6c25b-112">詳細については、「[コンテンツアドインと作業ウィンドウアドインでの API 使用のアクセス許可を要求](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)する」と「 [Outlook アドインのアクセス許可につい](../../outlook/understanding-outlook-add-in-permissions.md)て」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6c25b-112">For more details, see [Requesting permissions for API use in content and task pane add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
