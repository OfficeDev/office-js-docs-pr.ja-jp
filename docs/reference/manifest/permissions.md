---
title: マニフェスト ファイルの Permissions 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450661"
---
# <a name="permissions-element"></a><span data-ttu-id="9868c-102">Permissions 要素</span><span class="sxs-lookup"><span data-stu-id="9868c-102">Permissions element</span></span>

<span data-ttu-id="9868c-103">Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="9868c-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="9868c-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="9868c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9868c-105">構文</span><span class="sxs-lookup"><span data-stu-id="9868c-105">Syntax</span></span>

<span data-ttu-id="9868c-106">コンテンツ アドインおよび作業ウィンドウ アドインの場合:</span><span class="sxs-lookup"><span data-stu-id="9868c-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="9868c-107">メール アドインの場合</span><span class="sxs-lookup"><span data-stu-id="9868c-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="9868c-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="9868c-108">Contained in</span></span>

[<span data-ttu-id="9868c-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="9868c-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="9868c-110">注釈</span><span class="sxs-lookup"><span data-stu-id="9868c-110">Remarks</span></span>

<span data-ttu-id="9868c-111">詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)」と「[Outlook アドインのアクセス許可について](/outlook/add-ins/understanding-outlook-add-in-permissions)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="9868c-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
