---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f1aced0ce37b01c277ea5a8621f6c7764d2f761b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432348"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="3fcd6-102">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="3fcd6-102">AllowSnapshot element</span></span>

<span data-ttu-id="3fcd6-103">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="3fcd6-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="3fcd6-104">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="3fcd6-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="3fcd6-105">構文</span><span class="sxs-lookup"><span data-stu-id="3fcd6-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="3fcd6-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="3fcd6-106">Contained in</span></span>

[<span data-ttu-id="3fcd6-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3fcd6-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="3fcd6-108">解説</span><span class="sxs-lookup"><span data-stu-id="3fcd6-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="3fcd6-109">**AllowSnapshot** の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="3fcd6-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="3fcd6-110">この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。</span><span class="sxs-lookup"><span data-stu-id="3fcd6-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="3fcd6-111">しかしこれは、アドインがホストされるドキュメントから、アドインに表示される機密性の高い情報に直接アクセスできるということでもあります。</span><span class="sxs-lookup"><span data-stu-id="3fcd6-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

