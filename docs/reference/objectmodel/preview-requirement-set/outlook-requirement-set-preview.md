---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 07/24/2019
localization_priority: Priority
ms.openlocfilehash: 2ff1873afb0e0800c3056ae8de4033c56f357b2f
ms.sourcegitcommit: 5e90a90175909e0f4f392f5c98bd1273f444fe49
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/24/2019
ms.locfileid: "35851568"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f0f2a-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="f0f2a-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="f0f2a-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="f0f2a-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="f0f2a-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f0f2a-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="f0f2a-107">この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="f0f2a-108">また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="f0f2a-109">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f0f2a-110">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="f0f2a-110">Features in preview</span></span>

<span data-ttu-id="f0f2a-111">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="f0f2a-112">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="f0f2a-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="f0f2a-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="f0f2a-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="f0f2a-114">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="f0f2a-115">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-115">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="f0f2a-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f0f2a-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="f0f2a-117">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="f0f2a-118">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-118">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="f0f2a-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="f0f2a-120">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="f0f2a-121">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="f0f2a-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="f0f2a-123">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="f0f2a-124">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="f0f2a-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="f0f2a-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="f0f2a-126">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="f0f2a-127">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-127">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="f0f2a-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="f0f2a-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="f0f2a-129">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="f0f2a-130">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-130">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f0f2a-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="f0f2a-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f0f2a-132">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="f0f2a-133">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-133">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="f0f2a-134">送信のブロック</span><span class="sxs-lookup"><span data-stu-id="f0f2a-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="f0f2a-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="f0f2a-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="f0f2a-136">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="f0f2a-137">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="f0f2a-138">**使用できる場所**: Outlook on the web (クラシック)、Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-138">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="categories"></a><span data-ttu-id="f0f2a-139">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="f0f2a-139">Categories</span></span>

<span data-ttu-id="f0f2a-140">Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="f0f2a-141">ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="f0f2a-142">その後、アイテムに 1 つ以上のカテゴリを適用できます。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="f0f2a-143">この機能は Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="f0f2a-144">Categories</span><span class="sxs-lookup"><span data-stu-id="f0f2a-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="f0f2a-145">項目カテゴリを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="f0f2a-146">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-146">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="f0f2a-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="f0f2a-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="f0f2a-148">カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="f0f2a-149">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-149">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="f0f2a-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="f0f2a-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="f0f2a-151">メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f0f2a-152">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-152">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="f0f2a-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="f0f2a-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="f0f2a-154">メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f0f2a-155">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-155">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="f0f2a-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="f0f2a-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="f0f2a-157">アイテムのカテゴリのセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="f0f2a-158">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-158">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="f0f2a-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="f0f2a-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="f0f2a-160">カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="f0f2a-161">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-161">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="f0f2a-162">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="f0f2a-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="f0f2a-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="f0f2a-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="f0f2a-164">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="f0f2a-165">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-165">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="f0f2a-166">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-166">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="f0f2a-167">保存済みの予定またはメッセージ アイテムの ID を取得する新しいメソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="f0f2a-168">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-168">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="f0f2a-169">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-169">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="f0f2a-170">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-170">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="f0f2a-171">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-171">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="f0f2a-172">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="f0f2a-172">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="f0f2a-173">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-173">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="f0f2a-174">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-174">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="f0f2a-175">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="f0f2a-175">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="f0f2a-176">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-176">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="f0f2a-177">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-177">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="f0f2a-178">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-178">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="f0f2a-179">強化された場所</span><span class="sxs-lookup"><span data-stu-id="f0f2a-179">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="f0f2a-180">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f0f2a-180">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="f0f2a-181">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-181">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f0f2a-182">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-182">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="f0f2a-183">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="f0f2a-183">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="f0f2a-184">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-184">Added a new object that represents a location.</span></span> <span data-ttu-id="f0f2a-185">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-185">Read only.</span></span>

<span data-ttu-id="f0f2a-186">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-186">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="f0f2a-187">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="f0f2a-187">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="f0f2a-188">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-188">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="f0f2a-189">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-189">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="f0f2a-190">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f0f2a-190">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="f0f2a-191">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-191">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f0f2a-192">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-192">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="f0f2a-193">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="f0f2a-193">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="f0f2a-194">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-194">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="f0f2a-195">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-195">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f0f2a-196">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="f0f2a-196">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f0f2a-197">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-197">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="f0f2a-198">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-198">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f0f2a-199">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="f0f2a-199">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="f0f2a-200">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-200">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="f0f2a-201">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-201">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f0f2a-202">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="f0f2a-202">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="f0f2a-203">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="f0f2a-203">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="f0f2a-204">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f0f2a-204">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="f0f2a-205">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-205">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="f0f2a-206">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-206">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="f0f2a-207">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="f0f2a-207">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="f0f2a-208">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-208">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="f0f2a-209">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="f0f2a-209">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="f0f2a-210">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="f0f2a-210">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="f0f2a-211">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f0f2a-211">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="f0f2a-212">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-212">Added ability to get Office theme.</span></span>

<span data-ttu-id="f0f2a-213">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="f0f2a-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f0f2a-214">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f0f2a-214">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f0f2a-215">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-215">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f0f2a-216">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="f0f2a-216">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="f0f2a-217">SSO</span><span class="sxs-lookup"><span data-stu-id="f0f2a-217">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="f0f2a-218">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f0f2a-218">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="f0f2a-219">Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0f2a-219">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="f0f2a-220">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="f0f2a-220">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="f0f2a-221">関連項目</span><span class="sxs-lookup"><span data-stu-id="f0f2a-221">See also</span></span>

- [<span data-ttu-id="f0f2a-222">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="f0f2a-222">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="f0f2a-223">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="f0f2a-223">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f0f2a-224">概要</span><span class="sxs-lookup"><span data-stu-id="f0f2a-224">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="f0f2a-225">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="f0f2a-225">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
