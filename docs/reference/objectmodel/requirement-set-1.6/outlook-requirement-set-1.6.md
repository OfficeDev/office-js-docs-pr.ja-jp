---
title: Outlook アドイン API 要件セット 1.6
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 759286abe16e5bc32d6fd9ba1bb0e344d9d0a91c
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597034"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="4ae2a-102">Outlook アドイン API 要件セット 1.6</span><span class="sxs-lookup"><span data-stu-id="4ae2a-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="4ae2a-103">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="4ae2a-104">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="4ae2a-105">1.6 の新機能</span><span class="sxs-lookup"><span data-stu-id="4ae2a-105">What's new in 1.6?</span></span>

<span data-ttu-id="4ae2a-106">要件セット 1.6 には、[要件セット 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="4ae2a-107">次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-107">It added the following features.</span></span>

- <span data-ttu-id="4ae2a-108">ユーザーがアドインを有効にするために選択したエンティティまたは RegEx 一致を取得する、文脈アドインのための新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="4ae2a-109">新しいメッセージ フォームを開く新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="4ae2a-110">アドインがユーザーのメールボックスのアカウントの種類を決定するための機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="4ae2a-111">変更ログ</span><span class="sxs-lookup"><span data-stu-id="4ae2a-111">Change log</span></span>

- <span data-ttu-id="4ae2a-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) が追加されました: ユーザーが選択した強調表示された一致内で見つかったエンティティを取得する新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4ae2a-113">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="4ae2a-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) が追加されました: マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返す新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="4ae2a-115">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="4ae2a-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods) が追加されました: 新しいメッセージ フォームを表示する新しい関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="4ae2a-117">[Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype) が追加されました: ユーザーのアカウントの種類を示す新しいメンバーをユーザー プロファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="4ae2a-117">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="4ae2a-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="4ae2a-118">See also</span></span>

- [<span data-ttu-id="4ae2a-119">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="4ae2a-119">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="4ae2a-120">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="4ae2a-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="4ae2a-121">概要</span><span class="sxs-lookup"><span data-stu-id="4ae2a-121">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="4ae2a-122">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="4ae2a-122">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
