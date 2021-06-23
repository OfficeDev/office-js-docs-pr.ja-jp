---
title: Office アドインのダイアログ ボックス
description: アドインのダイアログの視覚的な設計に関するベスト プラクティスOffice説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d674b747effa57b8a75b79f98f5ff78ccc8a92a4
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076337"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="096db-103">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="096db-103">Dialog boxes in Office Add-ins</span></span>

<span data-ttu-id="096db-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="096db-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="096db-106">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="096db-106">*Figure 1. Typical layout for a dialog box*</span></span>

![アプリケーションに表示されるダイアログ ボックスの一般的なOfficeです。](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="096db-108">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="096db-108">Best practices</span></span>

|<span data-ttu-id="096db-109">するべきこと</span><span class="sxs-lookup"><span data-stu-id="096db-109">Do</span></span>|<span data-ttu-id="096db-110">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="096db-110">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="096db-111">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="096db-111">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="096db-112">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="096db-112">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="096db-113">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="096db-113">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="096db-114">実装</span><span class="sxs-lookup"><span data-stu-id="096db-114">Implementation</span></span>

<span data-ttu-id="096db-115">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="096db-115">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="096db-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="096db-116">See also</span></span>

- [<span data-ttu-id="096db-117">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="096db-117">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="096db-118">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="096db-118">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
