---
title: マニフェスト ファイルの AllowSnapshot 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02d44167dd1fd46ec6316f3e04393c99f19c9ff0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450675"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="695f9-102">AllowSnapshot 要素</span><span class="sxs-lookup"><span data-stu-id="695f9-102">AllowSnapshot element</span></span>

<span data-ttu-id="695f9-103">ホスト ドキュメントと共にコンテンツ アドインのスナップショット イメージを保存するかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="695f9-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="695f9-104">**アドインの種類:** コンテンツ</span><span class="sxs-lookup"><span data-stu-id="695f9-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="695f9-105">構文</span><span class="sxs-lookup"><span data-stu-id="695f9-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="695f9-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="695f9-106">Contained in</span></span>

[<span data-ttu-id="695f9-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="695f9-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="695f9-108">解説</span><span class="sxs-lookup"><span data-stu-id="695f9-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="695f9-109">**AllowSnapshot** の既定値は `true` です。</span><span class="sxs-lookup"><span data-stu-id="695f9-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="695f9-110">この場合、Office アドインをサポートしていないバージョンのホスト アプリケーションでドキュメントを開くユーザーがアドインのイメージを表示できるようになったり、ホスト アプリケーションがアドインをホストするサーバーに接続できない場合にアドインの静的イメージが提供されたりします。</span><span class="sxs-lookup"><span data-stu-id="695f9-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="695f9-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="695f9-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

