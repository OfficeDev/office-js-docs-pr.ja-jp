---
title: マニフェスト ファイルの Permissions 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 9193651ec0c795cdb55eb3fc6576dbacd59e0fb2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432355"
---
# <a name="permissions-element"></a><span data-ttu-id="337c2-102">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="337c2-102">Permissions element</span></span>

<span data-ttu-id="337c2-103">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="337c2-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="337c2-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="337c2-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="337c2-105">構文</span><span class="sxs-lookup"><span data-stu-id="337c2-105">Syntax</span></span>

<span data-ttu-id="337c2-106">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="337c2-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="337c2-107">メール アドインの場合</span><span class="sxs-lookup"><span data-stu-id="337c2-107">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="337c2-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="337c2-108">Contained in</span></span>

[<span data-ttu-id="337c2-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="337c2-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="337c2-110">解説</span><span class="sxs-lookup"><span data-stu-id="337c2-110">Remarks</span></span>

<span data-ttu-id="337c2-111">詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)」と「[Outlook アドインのアクセス許可について](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="337c2-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
