---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b46fada2fa69f3526c929a0289341f7dab5b58b8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128476"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="de157-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="de157-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="de157-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="de157-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="de157-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="de157-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="de157-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="de157-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="de157-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="de157-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="de157-107">この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="de157-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="de157-108">また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="de157-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="de157-109">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="de157-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="de157-110">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="de157-110">Features in preview</span></span>

<span data-ttu-id="de157-111">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="de157-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="de157-112">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="de157-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="de157-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="de157-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="de157-114">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="de157-115">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-115">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="de157-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="de157-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="de157-117">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="de157-118">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-118">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="de157-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="de157-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="de157-120">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="de157-121">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-121">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="de157-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="de157-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="de157-123">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="de157-124">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-124">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="de157-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="de157-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="de157-126">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="de157-127">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-127">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="de157-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="de157-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="de157-129">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="de157-130">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-130">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="de157-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="de157-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="de157-132">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="de157-133">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-133">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="de157-134">送信のブロック</span><span class="sxs-lookup"><span data-stu-id="de157-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="de157-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="de157-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="de157-136">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="de157-137">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="de157-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="de157-138">**使用できる場所**: Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="de157-138">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="categories"></a><span data-ttu-id="de157-139">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="de157-139">Categories</span></span>

<span data-ttu-id="de157-140">Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="de157-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="de157-141">ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。</span><span class="sxs-lookup"><span data-stu-id="de157-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="de157-142">その後、アイテムに 1 つ以上のカテゴリを適用できます。</span><span class="sxs-lookup"><span data-stu-id="de157-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="de157-143">この機能は Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="de157-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="de157-144">Categories</span><span class="sxs-lookup"><span data-stu-id="de157-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="de157-145">項目カテゴリを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="de157-146">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="de157-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="de157-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="de157-148">カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="de157-149">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="de157-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="de157-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="de157-151">メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="de157-152">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="de157-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="de157-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="de157-154">メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="de157-155">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="de157-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="de157-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="de157-157">アイテムのカテゴリのセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="de157-158">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="de157-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="de157-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="de157-160">カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="de157-161">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="de157-162">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="de157-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="de157-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="de157-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="de157-164">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="de157-165">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-165">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="de157-166">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="de157-166">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="de157-167">保存済みの予定またはメッセージ アイテムの ID を取得する新しいメソッドが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="de157-168">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-168">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="de157-169">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="de157-169">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="de157-170">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-170">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="de157-171">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-171">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="de157-172">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="de157-172">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="de157-173">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-173">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="de157-174">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-174">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="de157-175">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="de157-175">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="de157-176">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-176">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="de157-177">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="de157-177">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="de157-178">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-178">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="de157-179">強化された場所</span><span class="sxs-lookup"><span data-stu-id="de157-179">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="de157-180">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="de157-180">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="de157-181">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-181">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="de157-182">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-182">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="de157-183">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="de157-183">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="de157-184">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-184">Added a new object that represents a location.</span></span> <span data-ttu-id="de157-185">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="de157-185">Read only.</span></span>

<span data-ttu-id="de157-186">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-186">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="de157-187">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="de157-187">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="de157-188">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-188">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="de157-189">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-189">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="de157-190">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="de157-190">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="de157-191">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-191">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="de157-192">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-192">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="de157-193">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="de157-193">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="de157-194">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-194">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="de157-195">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-195">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="de157-196">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="de157-196">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="de157-197">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-197">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="de157-198">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-198">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="de157-199">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="de157-199">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="de157-200">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="de157-200">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="de157-201">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-201">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="de157-202">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="de157-202">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="de157-203">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="de157-203">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="de157-204">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="de157-204">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="de157-205">メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-205">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="de157-206">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="de157-207">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="de157-207">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="de157-208">メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-208">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="de157-209">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-209">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="de157-210">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="de157-210">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="de157-211">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="de157-211">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="de157-212">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-212">Added ability to get Office theme.</span></span>

<span data-ttu-id="de157-213">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="de157-214">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="de157-214">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="de157-215">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-215">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="de157-216">**使用できる場所**: Outlook on Windows (Office 365 サブスクリプションに接続している場合)</span><span class="sxs-lookup"><span data-stu-id="de157-216">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="de157-217">SSO</span><span class="sxs-lookup"><span data-stu-id="de157-217">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="de157-218">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="de157-218">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="de157-219">Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="de157-219">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="de157-220">**使用できる場所**: Office 365 サブスクリプションに接続している Outlook on Windows、Office 365 サブスクリプションに接続している Outlook on Mac、Outlook on the web (新規)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="de157-220">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (new), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="de157-221">関連項目</span><span class="sxs-lookup"><span data-stu-id="de157-221">See also</span></span>

- [<span data-ttu-id="de157-222">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="de157-222">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="de157-223">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="de157-223">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="de157-224">作業の開始</span><span class="sxs-lookup"><span data-stu-id="de157-224">Get started</span></span>](/outlook/add-ins/quick-start)
