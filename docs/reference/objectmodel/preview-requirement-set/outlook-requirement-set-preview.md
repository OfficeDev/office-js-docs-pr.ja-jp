---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: aa3f46c505e8c87508699f6e84194272ee4d13bb
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696457"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="c8c20-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="c8c20-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="c8c20-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c8c20-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c8c20-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="c8c20-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="c8c20-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="c8c20-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="c8c20-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="c8c20-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="c8c20-107">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c8c20-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="c8c20-108">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="c8c20-108">Features in preview</span></span>

<span data-ttu-id="c8c20-109">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="c8c20-109">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="c8c20-110">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="c8c20-110">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="c8c20-111">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="c8c20-111">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="c8c20-112">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-112">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="c8c20-113">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="c8c20-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c8c20-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="c8c20-115">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-115">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="c8c20-116">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-116">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="c8c20-117">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-117">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="c8c20-118">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-118">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="c8c20-119">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-119">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="c8c20-120">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-120">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="c8c20-121">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-121">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="c8c20-122">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-122">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="c8c20-123">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="c8c20-123">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="c8c20-124">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-124">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="c8c20-125">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="c8c20-126">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="c8c20-126">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="c8c20-127">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-127">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="c8c20-128">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="c8c20-129">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="c8c20-129">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="c8c20-130">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-130">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="c8c20-131">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="block-on-send"></a><span data-ttu-id="c8c20-132">送信のブロック</span><span class="sxs-lookup"><span data-stu-id="c8c20-132">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="c8c20-133">Event.completed</span><span class="sxs-lookup"><span data-stu-id="c8c20-133">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="c8c20-134">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-134">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="c8c20-135">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="c8c20-135">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="c8c20-136">**使用できる場所**: Outlook on the web (クラシック)、Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-136">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="categories"></a><span data-ttu-id="c8c20-137">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="c8c20-137">Categories</span></span>

<span data-ttu-id="c8c20-138">Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="c8c20-138">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="c8c20-139">ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。</span><span class="sxs-lookup"><span data-stu-id="c8c20-139">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="c8c20-140">その後、アイテムに 1 つ以上のカテゴリを適用できます。</span><span class="sxs-lookup"><span data-stu-id="c8c20-140">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="c8c20-141">この機能は Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c8c20-141">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="c8c20-142">Categories</span><span class="sxs-lookup"><span data-stu-id="c8c20-142">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="c8c20-143">項目カテゴリを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-143">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="c8c20-144">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="c8c20-145">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="c8c20-145">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="c8c20-146">カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-146">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="c8c20-147">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-147">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="c8c20-148">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="c8c20-148">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="c8c20-149">メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-149">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="c8c20-150">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="c8c20-151">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="c8c20-151">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="c8c20-152">メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-152">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="c8c20-153">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="c8c20-154">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="c8c20-154">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="c8c20-155">アイテムのカテゴリのセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-155">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="c8c20-156">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="c8c20-157">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="c8c20-157">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="c8c20-158">カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-158">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="c8c20-159">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="delegate-access"></a><span data-ttu-id="c8c20-160">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="c8c20-160">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="c8c20-161">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="c8c20-161">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="c8c20-162">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-162">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="c8c20-163">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="c8c20-164">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-164">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="c8c20-165">保存済みの予定またはメッセージ アイテムの ID を取得する新しいメソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-165">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="c8c20-166">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-166">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="c8c20-167">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-167">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="c8c20-168">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-168">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="c8c20-169">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="c8c20-170">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="c8c20-170">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="c8c20-171">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-171">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="c8c20-172">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-172">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="c8c20-173">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="c8c20-173">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="c8c20-174">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-174">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="c8c20-175">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="c8c20-175">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="c8c20-176">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="enhanced-location"></a><span data-ttu-id="c8c20-177">強化された場所</span><span class="sxs-lookup"><span data-stu-id="c8c20-177">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="c8c20-178">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c8c20-178">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="c8c20-179">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-179">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="c8c20-180">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="c8c20-181">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="c8c20-181">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="c8c20-182">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-182">Added a new object that represents a location.</span></span> <span data-ttu-id="c8c20-183">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="c8c20-183">Read only.</span></span>

<span data-ttu-id="c8c20-184">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="c8c20-185">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="c8c20-185">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="c8c20-186">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-186">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="c8c20-187">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="c8c20-188">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="c8c20-188">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="c8c20-189">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-189">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="c8c20-190">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-190">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="c8c20-191">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="c8c20-191">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="c8c20-192">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-192">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="c8c20-193">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-193">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="c8c20-194">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="c8c20-194">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="c8c20-195">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-195">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="c8c20-196">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-196">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="c8c20-197">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="c8c20-197">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="c8c20-198">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-198">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="c8c20-199">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-199">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="c8c20-200">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="c8c20-200">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

<br>

---

### <a name="internet-headers"></a><span data-ttu-id="c8c20-201">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="c8c20-201">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="c8c20-202">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c8c20-202">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="c8c20-203">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-203">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="c8c20-204">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-204">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="c8c20-205">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="c8c20-205">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="c8c20-206">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-206">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="c8c20-207">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="c8c20-207">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="office-theme"></a><span data-ttu-id="c8c20-208">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="c8c20-208">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="c8c20-209">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="c8c20-209">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="c8c20-210">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-210">Added ability to get Office theme.</span></span>

<span data-ttu-id="c8c20-211">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="c8c20-211">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="c8c20-212">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="c8c20-212">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="c8c20-213">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-213">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="c8c20-214">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="c8c20-214">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="c8c20-215">SSO</span><span class="sxs-lookup"><span data-stu-id="c8c20-215">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="c8c20-216">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c8c20-216">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="c8c20-217">Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="c8c20-217">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="c8c20-218">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="c8c20-218">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="c8c20-219">関連項目</span><span class="sxs-lookup"><span data-stu-id="c8c20-219">See also</span></span>

- [<span data-ttu-id="c8c20-220">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="c8c20-220">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="c8c20-221">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="c8c20-221">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c8c20-222">概要</span><span class="sxs-lookup"><span data-stu-id="c8c20-222">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="c8c20-223">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="c8c20-223">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
