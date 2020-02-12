---
title: Office アドインのダイアログ ボックス
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 6728e9032ba00c2e2ebcaa339f72700bc4dacca5
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950384"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="05088-102">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="05088-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="05088-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="05088-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="05088-105">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="05088-105">*Figure 1. Typical layout for a dialog box*</span></span>

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="05088-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="05088-107">Best practices</span></span>

|<span data-ttu-id="05088-108">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="05088-108">**Do**</span></span>|<span data-ttu-id="05088-109">**使用不可**</span><span class="sxs-lookup"><span data-stu-id="05088-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="05088-110">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="05088-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="05088-111">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="05088-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="05088-112">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="05088-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="05088-113">実装</span><span class="sxs-lookup"><span data-stu-id="05088-113">Implementation</span></span>

<span data-ttu-id="05088-114">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05088-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="05088-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="05088-115">See also</span></span>

- [<span data-ttu-id="05088-116">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="05088-116">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="05088-117">Office アドインの UX 設計パターン</span><span class="sxs-lookup"><span data-stu-id="05088-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
