---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c46dcd882592c0b015dae4b9774533b96fe75cfe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608790"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="763ba-103">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="763ba-103">AllowSnapshot element</span></span>

<span data-ttu-id="763ba-104">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="763ba-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="763ba-105">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="763ba-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="763ba-106">構文</span><span class="sxs-lookup"><span data-stu-id="763ba-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="763ba-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="763ba-107">Contained in</span></span>

[<span data-ttu-id="763ba-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="763ba-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="763ba-109">解説</span><span class="sxs-lookup"><span data-stu-id="763ba-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="763ba-110">**AllowSnapshot** の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="763ba-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="763ba-111">この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。</span><span class="sxs-lookup"><span data-stu-id="763ba-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="763ba-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="763ba-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

