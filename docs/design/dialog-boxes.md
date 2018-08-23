---
title: Office アドインのダイアログ ボックス
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: af47fd338872d3ecfce06145783fcc9ff314f7bc
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/23/2018
ms.locfileid: "19437144"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="c4088-102">Office アドインのダイアログ ボックス</span><span class="sxs-lookup"><span data-stu-id="c4088-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="c4088-p101">ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。</span><span class="sxs-lookup"><span data-stu-id="c4088-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="c4088-105">*図 1. ダイアログ ボックスの一般的なレイアウト*</span><span class="sxs-lookup"><span data-stu-id="c4088-105">*Figure 1. Typical layout for a dialog box*</span></span>

![ダイアログ ボックスの一般的なレイアウトを表示する画像の例](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="c4088-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="c4088-107">Best practices</span></span>

|<span data-ttu-id="c4088-108">**するべきこと**</span><span class="sxs-lookup"><span data-stu-id="c4088-108">**Do**</span></span>|<span data-ttu-id="c4088-109">**使用不可**</span><span class="sxs-lookup"><span data-stu-id="c4088-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="c4088-110">アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c4088-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="c4088-111">タイトルには会社名を追加しません。</span><span class="sxs-lookup"><span data-stu-id="c4088-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="c4088-112">シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</span><span class="sxs-lookup"><span data-stu-id="c4088-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="c4088-113">実装</span><span class="sxs-lookup"><span data-stu-id="c4088-113">Implementation</span></span>

<span data-ttu-id="c4088-114">ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c4088-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="c4088-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="c4088-115">See also</span></span>

- [<span data-ttu-id="c4088-116">UX パターンのサンプル</span><span class="sxs-lookup"><span data-stu-id="c4088-116">UX Pattern Sample</span></span>](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [<span data-ttu-id="c4088-117">GitHub の開発リソース</span><span class="sxs-lookup"><span data-stu-id="c4088-117">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="c4088-118">Dialog オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c4088-118">Dialog object</span></span>](https://dev.office.com/reference/add-ins/shared/officeui.dialog)


