---
title: Office アドインのダイアログ ボックス
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 3638006c30515a1fcae93ccfdbd43e0e92005c37
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691105"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="fef87-102">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="fef87-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="fef87-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="fef87-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="fef87-105">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="fef87-105">*Figure 1. Typical layout for a dialog box*</span></span>

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="fef87-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="fef87-107">Best practices</span></span>

|<span data-ttu-id="fef87-108">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="fef87-108">**Do**</span></span>|<span data-ttu-id="fef87-109">**使用不可**</span><span class="sxs-lookup"><span data-stu-id="fef87-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="fef87-110">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="fef87-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="fef87-111">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="fef87-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="fef87-112">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="fef87-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="fef87-113">実装</span><span class="sxs-lookup"><span data-stu-id="fef87-113">Implementation</span></span>

<span data-ttu-id="fef87-114">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fef87-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="fef87-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="fef87-115">See also</span></span>

- [<span data-ttu-id="fef87-116">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fef87-116">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [<span data-ttu-id="fef87-117">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="fef87-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)


