---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: b563d6cfc279a18a6a61f39c33a5ab42e1bd6984
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395709"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="17789-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="17789-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="17789-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="17789-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="17789-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="17789-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="17789-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="17789-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="17789-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="17789-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="17789-107">この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="17789-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span>
>
> <span data-ttu-id="17789-108">プレビュー API を使用するには:</span><span class="sxs-lookup"><span data-stu-id="17789-108">To use preview APIs:</span></span>
>
> - <span data-ttu-id="17789-109">CDN (\*\* の\*\*ベータhttps://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ライブラリを参照する必要があります。</span><span class="sxs-lookup"><span data-stu-id="17789-109">You must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span>
> - <span data-ttu-id="17789-110">最新の Office ビルドにアクセスするには、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="17789-110">You may also need to join the [Office Insider program](https://products.office.com/office-insider) for access to more recent Office builds.</span></span>

<span data-ttu-id="17789-111">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="17789-111">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="17789-112">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="17789-112">Features in preview</span></span>

<span data-ttu-id="17789-113">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="17789-113">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="17789-114">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="17789-114">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="17789-115">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="17789-115">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="17789-116">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-116">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="17789-117">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-117">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="17789-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="17789-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="17789-119">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-119">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="17789-120">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="17789-121">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="17789-121">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="17789-122">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-122">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="17789-123">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-123">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="17789-124">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="17789-124">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="17789-125">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-125">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="17789-126">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-126">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="17789-127">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="17789-127">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="17789-128">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-128">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="17789-129">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-129">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="17789-130">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="17789-130">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="17789-131">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-131">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="17789-132">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="17789-133">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="17789-133">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="17789-134">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-134">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="17789-135">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="17789-136">送信のブロック</span><span class="sxs-lookup"><span data-stu-id="17789-136">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="17789-137">Event.completed</span><span class="sxs-lookup"><span data-stu-id="17789-137">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="17789-138">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-138">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="17789-139">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="17789-139">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="17789-140">**使用できる場所**: Outlook on the web (クラシック)、Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-140">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="categories"></a><span data-ttu-id="17789-141">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="17789-141">Categories</span></span>

<span data-ttu-id="17789-142">Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="17789-142">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="17789-143">ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。</span><span class="sxs-lookup"><span data-stu-id="17789-143">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="17789-144">その後、アイテムに 1 つ以上のカテゴリを適用できます。</span><span class="sxs-lookup"><span data-stu-id="17789-144">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="17789-145">この機能は Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="17789-145">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="17789-146">Categories</span><span class="sxs-lookup"><span data-stu-id="17789-146">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="17789-147">項目カテゴリを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-147">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="17789-148">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="17789-149">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="17789-149">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="17789-150">カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-150">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="17789-151">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-151">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="17789-152">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="17789-152">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="17789-153">メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-153">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="17789-154">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-154">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="17789-155">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="17789-155">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="17789-156">メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-156">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="17789-157">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-157">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="17789-158">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="17789-158">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="17789-159">アイテムのカテゴリのセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-159">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="17789-160">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="17789-161">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="17789-161">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="17789-162">カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-162">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="17789-163">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="17789-164">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="17789-164">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="17789-165">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="17789-165">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="17789-166">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-166">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="17789-167">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="17789-168">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="17789-168">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="17789-169">保存済みの予定またはメッセージ アイテムの ID を取得する新しいメソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-169">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="17789-170">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="17789-171">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="17789-171">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="17789-172">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-172">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="17789-173">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="17789-174">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="17789-174">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="17789-175">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-175">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="17789-176">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="17789-177">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="17789-177">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="17789-178">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-178">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="17789-179">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="17789-179">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="17789-180">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="17789-181">強化された場所</span><span class="sxs-lookup"><span data-stu-id="17789-181">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="17789-182">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="17789-182">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="17789-183">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-183">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="17789-184">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="17789-185">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="17789-185">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="17789-186">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-186">Added a new object that represents a location.</span></span> <span data-ttu-id="17789-187">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="17789-187">Read only.</span></span>

<span data-ttu-id="17789-188">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-188">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="17789-189">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="17789-189">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="17789-190">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-190">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="17789-191">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-191">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="17789-192">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="17789-192">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="17789-193">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-193">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="17789-194">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-194">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="17789-195">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="17789-195">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="17789-196">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-196">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="17789-197">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-197">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="17789-198">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="17789-198">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="17789-199">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-199">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="17789-200">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (モダン)、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-200">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="17789-201">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="17789-201">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="17789-202">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="17789-202">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="17789-203">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-203">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="17789-204">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="17789-204">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="17789-205">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="17789-205">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="17789-206">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="17789-206">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="17789-207">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-207">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="17789-208">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-208">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="17789-209">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="17789-209">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="17789-210">メッセージ アイテムのカスタム インターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-210">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="17789-211">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac</span><span class="sxs-lookup"><span data-stu-id="17789-211">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="17789-212">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="17789-212">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="17789-213">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="17789-213">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="17789-214">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-214">Added ability to get Office theme.</span></span>

<span data-ttu-id="17789-215">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="17789-215">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="17789-216">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="17789-216">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="17789-217">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-217">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="17789-218">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="17789-218">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="17789-219">SSO</span><span class="sxs-lookup"><span data-stu-id="17789-219">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="17789-220">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="17789-220">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="17789-221">Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="17789-221">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="17789-222">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (モダン)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="17789-222">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="17789-223">関連項目</span><span class="sxs-lookup"><span data-stu-id="17789-223">See also</span></span>

- [<span data-ttu-id="17789-224">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="17789-224">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="17789-225">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="17789-225">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="17789-226">概要</span><span class="sxs-lookup"><span data-stu-id="17789-226">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="17789-227">要求セットとサポートされているクライアント</span><span class="sxs-lookup"><span data-stu-id="17789-227">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
