---
title: Outlook アドイン API 要件セット 1.6
description: Outlook アドインおよび Office JavaScript Api for the Mailbox API 1.6 の一部として導入された機能と Api。
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: c1ce30ef1dd717a5d19ef9d71cf737e342cd660f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717636"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="bfb92-103">Outlook アドイン API 要件セット 1.6</span><span class="sxs-lookup"><span data-stu-id="bfb92-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="bfb92-104">Office JavaScript API の Outlook アドイン API サブセットには、Outlook アドインで使用できるオブジェクト、メソッド、プロパティ、およびイベントが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bfb92-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb92-105">このドキュメントは、最新の要件セット以外の[要件セット](../../requirement-sets/outlook-api-requirement-sets.md)のためのものです。</span><span class="sxs-lookup"><span data-stu-id="bfb92-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="bfb92-106">1.6 の新機能</span><span class="sxs-lookup"><span data-stu-id="bfb92-106">What's new in 1.6?</span></span>

<span data-ttu-id="bfb92-107">要件セット 1.6 には、[要件セット 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="bfb92-107">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="bfb92-108">次の機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb92-108">It added the following features.</span></span>

- <span data-ttu-id="bfb92-109">ユーザーがアドインを有効にするために選択したエンティティまたは RegEx 一致を取得する、文脈アドインのための新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb92-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="bfb92-110">新しいメッセージ フォームを開く新しい API が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb92-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="bfb92-111">アドインがユーザーのメールボックスのアカウントの種類を決定するための機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb92-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="bfb92-112">変更ログ</span><span class="sxs-lookup"><span data-stu-id="bfb92-112">Change log</span></span>

- <span data-ttu-id="bfb92-113">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) が追加されました: ユーザーが選択した強調表示された一致内で見つかったエンティティを取得する新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb92-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="bfb92-114">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="bfb92-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="bfb92-115">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) が追加されました: マニフェスト XML ファイルで定義した正規表現と一致する、強調表示された一致内の文字列値を返す新機能を追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb92-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="bfb92-116">強調表示された一致は、コンテキスト アドインに適用されます。</span><span class="sxs-lookup"><span data-stu-id="bfb92-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="bfb92-117">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods) が追加されました: 新しいメッセージ フォームを表示する新しい関数を追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb92-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="bfb92-118">[Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype) が追加されました: ユーザーのアカウントの種類を示す新しいメンバーをユーザー プロファイルに追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb92-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="bfb92-119">関連項目</span><span class="sxs-lookup"><span data-stu-id="bfb92-119">See also</span></span>

- [<span data-ttu-id="bfb92-120">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="bfb92-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="bfb92-121">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="bfb92-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="bfb92-122">概要</span><span class="sxs-lookup"><span data-stu-id="bfb92-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="bfb92-123">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="bfb92-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
