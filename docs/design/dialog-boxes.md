---
title: Office アドインのダイアログ ボックス
description: Office アドインでのダイアログのビジュアルデザインのベストプラクティスについて説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2f3b25fac7f12494e6b5a1e0a32e72baa345e978
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717195"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="58374-103">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="58374-103">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="58374-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="58374-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="58374-106">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="58374-106">*Figure 1. Typical layout for a dialog box*</span></span>

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="58374-108">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="58374-108">Best practices</span></span>

|<span data-ttu-id="58374-109">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="58374-109">**Do**</span></span>|<span data-ttu-id="58374-110">**使用不可**</span><span class="sxs-lookup"><span data-stu-id="58374-110">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="58374-111">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="58374-111">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="58374-112">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="58374-112">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="58374-113">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="58374-113">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="58374-114">実装</span><span class="sxs-lookup"><span data-stu-id="58374-114">Implementation</span></span>

<span data-ttu-id="58374-115">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58374-115">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="58374-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="58374-116">See also</span></span>

- [<span data-ttu-id="58374-117">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="58374-117">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="58374-118">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="58374-118">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
