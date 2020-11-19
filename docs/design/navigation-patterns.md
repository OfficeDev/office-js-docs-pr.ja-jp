---
title: Office アドインのナビゲーション パターン
description: コマンドバー、タブバー、および [戻る] ボタンを使用して、Office アドインのナビゲーションを設計するためのベストプラクティスについて説明します。
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132033"
---
# <a name="navigation-patterns"></a><span data-ttu-id="d547b-103">ナビゲーション パターン</span><span class="sxs-lookup"><span data-stu-id="d547b-103">Navigation patterns</span></span>

<span data-ttu-id="d547b-104">アドインの主な機能には、特定のコマンドの種類と限られた画面領域を介してアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d547b-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="d547b-105">ナビゲーションは直感的で、コンテキストを提供し、アドイン全体においてユーザーが簡単に移動できることが重要です。</span><span class="sxs-lookup"><span data-stu-id="d547b-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="d547b-106">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="d547b-106">Best practices</span></span>

| <span data-ttu-id="d547b-107">するべきこと</span><span class="sxs-lookup"><span data-stu-id="d547b-107">Do</span></span>    | <span data-ttu-id="d547b-108">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="d547b-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="d547b-109">ユーザーに分かりやすいナビゲーション オプションが表示されるようにする。</span><span class="sxs-lookup"><span data-stu-id="d547b-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="d547b-110">標準的ではない UI を使用してナビゲーション プロセスを複雑にしない。</span><span class="sxs-lookup"><span data-stu-id="d547b-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="d547b-111">可能な場合には以下のコンポーネントを利用して、ユーザーがアドイン内でナビゲートできるようにする。</span><span class="sxs-lookup"><span data-stu-id="d547b-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="d547b-112">ユーザーが、アドインにおける現在の場所またはコンテキストを理解しにくいという状況を避ける。</span><span class="sxs-lookup"><span data-stu-id="d547b-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>

## <a name="command-bar"></a><span data-ttu-id="d547b-113">コマンド バー</span><span class="sxs-lookup"><span data-stu-id="d547b-113">Command Bar</span></span>

<span data-ttu-id="d547b-114">CommandBar は、作業ウィンドウ内の領域で、上にあるウィンドウ、パネル、または親地域の内容を操作するコマンドを格納します。</span><span class="sxs-lookup"><span data-stu-id="d547b-114">The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="d547b-115">オプション機能には、ハンバーガー メニューのアクセス ポイント、検索、およびサイド コマンドが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d547b-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Office デスクトップアプリケーションの作業ウィンドウ内のコマンドバーを示す図](../images/add-in-command-bar.png)

## <a name="tab-bar"></a><span data-ttu-id="d547b-118">タブ バー</span><span class="sxs-lookup"><span data-stu-id="d547b-118">Tab Bar</span></span>

<span data-ttu-id="d547b-119">タブバーは、テキストとアイコンが縦に並んだボタンを使用してナビゲーションを表示します。</span><span class="sxs-lookup"><span data-stu-id="d547b-119">The tab bar shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="d547b-120">タブ バーを使用すると、短くてわかりやすいタイトルのタブが使用されたナビゲーションを表示できます。</span><span class="sxs-lookup"><span data-stu-id="d547b-120">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Office デスクトップアプリケーションの作業ウィンドウ内のタブバーを示す図](../images/add-in-tab-bar.png)

## <a name="back-button"></a><span data-ttu-id="d547b-123">[戻る] ボタン</span><span class="sxs-lookup"><span data-stu-id="d547b-123">Back Button</span></span>

<span data-ttu-id="d547b-124">[戻る] ボタンを使用すると、ユーザーはドリルダウンナビゲーションアクションから回復することができます。</span><span class="sxs-lookup"><span data-stu-id="d547b-124">The back button allows users to recover from a drill-down navigational action.</span></span> <span data-ttu-id="d547b-125">このパターンは、ユーザーが順序のある一連の手順に従えるようにするのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="d547b-125">This pattern helps ensure users follow an ordered series of steps.</span></span>

![Office デスクトップアプリケーション作業ウィンドウ内の [戻る] ボタンを示す図](../images/add-in-back-button.png)
