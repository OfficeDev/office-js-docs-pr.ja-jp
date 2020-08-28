---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ea910e1ad747e304dbc6ab4fbdcf44a9610dab19
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294277"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="68955-103">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="68955-103">AllowSnapshot element</span></span>

<span data-ttu-id="68955-104">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="68955-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="68955-105">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="68955-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="68955-106">構文</span><span class="sxs-lookup"><span data-stu-id="68955-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="68955-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="68955-107">Contained in</span></span>

[<span data-ttu-id="68955-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="68955-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="68955-109">解説</span><span class="sxs-lookup"><span data-stu-id="68955-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="68955-110">**AllowSnapshot** の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="68955-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="68955-111">これにより、Office アドインをサポートしていないバージョンの Office アプリケーションでドキュメントを開くユーザーに対してアドインのイメージが表示されるようになります。または、アプリケーションがアドインをホストしているサーバーに接続できない場合は、アドインの静的イメージを提供します。</span><span class="sxs-lookup"><span data-stu-id="68955-111">This makes an image of the add-in visible for users that open the document in a version of the Office application that doesn't support Office Add-ins, or provides a static image of the add-in if the application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="68955-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="68955-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>
