---
title: Office アドインのダイアログ ボックス
description: ''
ms.date: 2/28/2019
localization_priority: Priority
ms.openlocfilehash: 1710d609910cc3c15143605570f97d013a104194
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359220"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="cd5b8-102">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="cd5b8-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="cd5b8-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="cd5b8-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="cd5b8-105">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="cd5b8-105">*Figure 1. Typical layout for a dialog box*</span></span>

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="cd5b8-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="cd5b8-107">Best practices</span></span>

|<span data-ttu-id="cd5b8-108">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="cd5b8-108">**Do**</span></span>|<span data-ttu-id="cd5b8-109">**使用不可**</span><span class="sxs-lookup"><span data-stu-id="cd5b8-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="cd5b8-110">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="cd5b8-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="cd5b8-111">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="cd5b8-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="cd5b8-112">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="cd5b8-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="cd5b8-113">実装</span><span class="sxs-lookup"><span data-stu-id="cd5b8-113">Implementation</span></span>

<span data-ttu-id="cd5b8-114">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cd5b8-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="cd5b8-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="cd5b8-115">See also</span></span>

- [<span data-ttu-id="cd5b8-116">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="cd5b8-116">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [<span data-ttu-id="cd5b8-117">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="cd5b8-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)


