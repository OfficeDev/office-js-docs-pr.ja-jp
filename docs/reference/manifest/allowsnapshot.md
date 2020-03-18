---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720961"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="7baf8-103">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="7baf8-103">AllowSnapshot element</span></span>

<span data-ttu-id="7baf8-104">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="7baf8-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="7baf8-105">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7baf8-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="7baf8-106">構文</span><span class="sxs-lookup"><span data-stu-id="7baf8-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="7baf8-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="7baf8-107">Contained in</span></span>

[<span data-ttu-id="7baf8-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="7baf8-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="7baf8-109">解説</span><span class="sxs-lookup"><span data-stu-id="7baf8-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="7baf8-110">**AllowSnapshot** の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="7baf8-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="7baf8-111">この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。</span><span class="sxs-lookup"><span data-stu-id="7baf8-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="7baf8-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="7baf8-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

