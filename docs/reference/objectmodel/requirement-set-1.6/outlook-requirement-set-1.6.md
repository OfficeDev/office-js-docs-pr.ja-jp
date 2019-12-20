---
title: Outlook アドイン API 要件セット 1.6
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 22702448b82a108c401f9f81d3b8a321e14ead63
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814662"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="24473-102">Outlook アドイン API 要件セット 1.6</span><span class="sxs-lookup"><span data-stu-id="24473-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="24473-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="24473-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="24473-104">このドキュメントは、最新の要件セット以外の[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="24473-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="24473-105">1.6 の新機能</span><span class="sxs-lookup"><span data-stu-id="24473-105">What's new in 1.6?</span></span>

<span data-ttu-id="24473-106">要件セット 1.6 には、[要件セット 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="24473-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="24473-107">次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="24473-107">It added the following features.</span></span>

- <span data-ttu-id="24473-108">ユーザーがアドインを有効にするために選択したエンティティまたは RegEx 一致を取得する、文脈アドインのための新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="24473-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="24473-109">新しいメッセージ フォームを開く新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="24473-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="24473-110">アドインがユーザーのメールボックスのアカウントの種類を決定するための機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="24473-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="24473-111">変更ログ</span><span class="sxs-lookup"><span data-stu-id="24473-111">Change log</span></span>

- <span data-ttu-id="24473-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) が追加されました: ユーザーが選択した強調表示された一致内で見つかったエンティティを取得する新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="24473-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="24473-113">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="24473-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="24473-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) が追加されました: マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返す新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="24473-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="24473-115">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="24473-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="24473-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods) が追加されました: 新しいメッセージ フォームを表示する新しい関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="24473-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="24473-117">[Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties) が追加されました: ユーザーのアカウントの種類を示す新しいメンバーをユーザー プロファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="24473-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="24473-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="24473-118">See also</span></span>

- [<span data-ttu-id="24473-119">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="24473-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="24473-120">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="24473-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="24473-121">概要</span><span class="sxs-lookup"><span data-stu-id="24473-121">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="24473-122">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="24473-122">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
