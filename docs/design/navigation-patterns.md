---
title: Office アドインのナビゲーション パターン
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: f0482f7742c6fab97fe563d61d21135c072f8f8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449143"
---
# <a name="navigation-patterns"></a><span data-ttu-id="4ae36-102">ナビゲーション パターン</span><span class="sxs-lookup"><span data-stu-id="4ae36-102">Navigation patterns</span></span>

<span data-ttu-id="4ae36-103">アドインの主な機能には、特定のコマンドの種類と限られた画面領域を介してアクセスします。</span><span class="sxs-lookup"><span data-stu-id="4ae36-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="4ae36-104">ナビゲーションは直感的で、コンテキストを提供し、アドイン全体においてユーザーが簡単に移動できることが重要です。</span><span class="sxs-lookup"><span data-stu-id="4ae36-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="4ae36-105">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="4ae36-105">Best practices</span></span>

| <span data-ttu-id="4ae36-106">するべきこと</span><span class="sxs-lookup"><span data-stu-id="4ae36-106">Do</span></span>    | <span data-ttu-id="4ae36-107">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="4ae36-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="4ae36-108">ユーザーに分かりやすいナビゲーション オプションが表示されるようにする。</span><span class="sxs-lookup"><span data-stu-id="4ae36-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="4ae36-109">標準的ではない UI を使用してナビゲーション プロセスを複雑にしない。</span><span class="sxs-lookup"><span data-stu-id="4ae36-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="4ae36-110">可能な場合には以下のコンポーネントを利用して、ユーザーがアドイン内でナビゲートできるようにする。</span><span class="sxs-lookup"><span data-stu-id="4ae36-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="4ae36-111">ユーザーが、アドインにおける現在の場所またはコンテキストを理解しにくいという状況を避ける。</span><span class="sxs-lookup"><span data-stu-id="4ae36-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="4ae36-112">コマンド バー</span><span class="sxs-lookup"><span data-stu-id="4ae36-112">Command Bar</span></span>

<span data-ttu-id="4ae36-113">コマンド バーは、ウィンドウ、パネル、またはその上にある親領域のコンテンツを操作するコマンドを格納するサーフェスです。</span><span class="sxs-lookup"><span data-stu-id="4ae36-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="4ae36-114">オプション機能には、ハンバーガー メニューのアクセス ポイント、検索、およびサイド コマンドが含まれます。</span><span class="sxs-lookup"><span data-stu-id="4ae36-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![コマンド - デスクトップ作業ウィンドウの仕様](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="4ae36-116">タブ バー</span><span class="sxs-lookup"><span data-stu-id="4ae36-116">Tab Bar</span></span>

<span data-ttu-id="4ae36-117">テキストとアイコンが縦に並んだボタンが使用されたナビゲーションを表示します。</span><span class="sxs-lookup"><span data-stu-id="4ae36-117">Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="4ae36-118">タブ バーを使用すると、短くてわかりやすいタイトルのタブが使用されたナビゲーションを表示できます。</span><span class="sxs-lookup"><span data-stu-id="4ae36-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![タブ バー - デスクトップ作業ウィンドウの仕様](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="4ae36-120">[戻る] ボタン</span><span class="sxs-lookup"><span data-stu-id="4ae36-120">Back Button</span></span>

<span data-ttu-id="4ae36-121">[戻る] ボタンを使用すると、ドリルダウンのナビゲーション操作から戻ることができます。</span><span class="sxs-lookup"><span data-stu-id="4ae36-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="4ae36-122">このパターンは、ユーザーが順序のある一連の手順に従えるようにするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="4ae36-122">This pattern helps ensure users follow an ordered series of steps.</span></span>  

![[戻る] ボタン - デスクトップ作業ウィンドウの仕様](../images/add-in-back-button.png)
