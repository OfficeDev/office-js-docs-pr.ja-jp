---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: e4627699edad801ab4a3a5a65e6307d40d1b4ac9
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952356"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="f397d-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="f397d-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="f397d-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f397d-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="f397d-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="f397d-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="f397d-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="f397d-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="f397d-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="f397d-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="f397d-107">この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f397d-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="f397d-108">また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="f397d-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="f397d-109">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f397d-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="f397d-110">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="f397d-110">Features in preview</span></span>

<span data-ttu-id="f397d-111">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="f397d-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="f397d-112">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f397d-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="f397d-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="f397d-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="f397d-114">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-114">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="f397d-115">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="f397d-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="f397d-116">**使用できる場所**: Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="f397d-116">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="attachments"></a><span data-ttu-id="f397d-117">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="f397d-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="f397d-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="f397d-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="f397d-119">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-119">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="f397d-120">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-120">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="f397d-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f397d-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="f397d-122">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-122">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="f397d-123">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-123">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="f397d-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f397d-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="f397d-125">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-125">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="f397d-126">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-126">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="f397d-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f397d-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="f397d-128">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-128">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="f397d-129">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-129">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="f397d-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="f397d-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="f397d-131">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-131">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="f397d-132">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-132">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="f397d-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="f397d-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="f397d-134">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-134">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="f397d-135">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-135">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f397d-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="f397d-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f397d-137">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="f397d-138">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-138">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="categories"></a><span data-ttu-id="f397d-139">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="f397d-139">Categories</span></span>

<span data-ttu-id="f397d-140">Outlook では、ユーザーはカテゴリを使用してメッセージと予定を色分けしてグループ化できます。</span><span class="sxs-lookup"><span data-stu-id="f397d-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="f397d-141">ユーザーは自分のメールボックスのマスター リストにカテゴリを定義します。</span><span class="sxs-lookup"><span data-stu-id="f397d-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="f397d-142">その後、アイテムに 1 つ以上のカテゴリを適用できます。</span><span class="sxs-lookup"><span data-stu-id="f397d-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="f397d-143">この機能は、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f397d-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="f397d-144">カテゴリ</span><span class="sxs-lookup"><span data-stu-id="f397d-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="f397d-145">項目カテゴリを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="f397d-146">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="f397d-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="f397d-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="f397d-148">カテゴリの詳細 (名前とそれに関連付けられた色) を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="f397d-149">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="f397d-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="f397d-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="f397d-151">メールボックスのカテゴリ マスター リストを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f397d-152">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="f397d-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="f397d-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="f397d-154">メールボックスのカテゴリ マスター リストを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="f397d-155">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="f397d-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="f397d-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="f397d-157">アイテムのカテゴリのセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="f397d-158">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="f397d-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="f397d-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="f397d-160">カテゴリに関連付ける使用可能な色を指定する新しい列挙が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="f397d-161">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="f397d-162">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="f397d-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="f397d-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="f397d-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="f397d-164">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="f397d-165">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-165">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="f397d-166">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f397d-166">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="f397d-167">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="f397d-168">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-168">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="f397d-169">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="f397d-169">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="f397d-170">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-170">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="f397d-171">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-171">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="f397d-172">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="f397d-172">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="f397d-173">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-173">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="f397d-174">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="f397d-174">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="f397d-175">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-175">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="f397d-176">強化された場所</span><span class="sxs-lookup"><span data-stu-id="f397d-176">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="f397d-177">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f397d-177">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="f397d-178">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-178">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f397d-179">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-179">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="f397d-180">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="f397d-180">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="f397d-181">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-181">Added a new object that represents a location.</span></span> <span data-ttu-id="f397d-182">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f397d-182">Read only.</span></span>

<span data-ttu-id="f397d-183">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-183">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="f397d-184">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="f397d-184">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="f397d-185">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-185">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="f397d-186">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-186">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="f397d-187">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f397d-187">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="f397d-188">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-188">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="f397d-189">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-189">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="f397d-190">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="f397d-190">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="f397d-191">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-191">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="f397d-192">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-192">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f397d-193">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="f397d-193">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f397d-194">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-194">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="f397d-195">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-195">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="f397d-196">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="f397d-196">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="f397d-197">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="f397d-197">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="f397d-198">アドインが[操作可能メッセージによってアクティブ化](/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-198">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="f397d-199">**使用できる場所**: Office 365 に接続している Outlook for Windows、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="f397d-199">**Available in**: Outlook for Windows (Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="f397d-200">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="f397d-200">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="f397d-201">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f397d-201">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="f397d-202">メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-202">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="f397d-203">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-203">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="f397d-204">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="f397d-204">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="f397d-205">メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-205">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="f397d-206">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="f397d-207">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="f397d-207">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="f397d-208">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="f397d-208">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="f397d-209">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-209">Added ability to get Office theme.</span></span>

<span data-ttu-id="f397d-210">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-210">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="f397d-211">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="f397d-211">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="f397d-212">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-212">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="f397d-213">**使用できる場所**: Office 365 に接続している Outlook for Windows</span><span class="sxs-lookup"><span data-stu-id="f397d-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="f397d-214">SSO</span><span class="sxs-lookup"><span data-stu-id="f397d-214">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="f397d-215">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f397d-215">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="f397d-216">Microsoft Graph API の[アクセス トークンの取得](/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="f397d-216">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="f397d-217">**使用できる場所**: Office 365 に接続している Outlook for Windows、Office 365 に接続している Outlook for Mac、Outlook.com と Office 365 に接続されている Outlook on the web、Outlook on the web (クラシック) </span><span class="sxs-lookup"><span data-stu-id="f397d-217">**Available in**: Outlook on Windows (connected to Office 365), Outlook for Mac (connected to Office 365), Outlook on the web (Outlook.com and connected to Office 365), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="f397d-218">関連項目</span><span class="sxs-lookup"><span data-stu-id="f397d-218">See also</span></span>

- [<span data-ttu-id="f397d-219">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="f397d-219">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="f397d-220">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="f397d-220">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f397d-221">作業の開始</span><span class="sxs-lookup"><span data-stu-id="f397d-221">Get started</span></span>](/outlook/add-ins/quick-start)
