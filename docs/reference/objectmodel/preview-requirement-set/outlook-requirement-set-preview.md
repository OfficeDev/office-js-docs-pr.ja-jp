---
title: Outlook アドイン API 要件セットのプレビュー
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512861"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="1d4d6-102">Outlook アドイン API 要件セットのプレビュー</span><span class="sxs-lookup"><span data-stu-id="1d4d6-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="1d4d6-103">JavaScript API for Office の Outlook アドイン API サブセットには、Outlook アドインで利用できるオブジェクト、メソッド、プロパティ、イベントが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1d4d6-104">このドキュメントは、[要件セット](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)の**プレビュー**用です。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="1d4d6-105">この要件セットはまだ完全には実装されていないため、このサポートはクライアントによって正確に報告されません。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="1d4d6-106">アドイン マニフェストでこの要件を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="1d4d6-107">この要件のセットに導入されているメソッドとプロパティは、使用前に可用性を個別にテストする必要があります。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="1d4d6-108">また、場合によっては [Office Insider プログラム](https://products.office.com/office-insider)に参加する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="1d4d6-109">要件セットのプレビューには、[要件セット 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) のすべての機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="1d4d6-110">プレビューの機能</span><span class="sxs-lookup"><span data-stu-id="1d4d6-110">Features in preview</span></span>

<span data-ttu-id="1d4d6-111">次の機能はプレビュー段階です。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="1d4d6-112">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="1d4d6-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="1d4d6-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="1d4d6-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="1d4d6-114">1 つの有効な値 `allowEvent` を持つディクショナリである、新しいオプション パラメーター `options` が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-114">Event.completed - A new optional parameter , which is a dictionary with one valid value .</span></span> <span data-ttu-id="1d4d6-115">この値は、イベントの実行をキャンセルするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="1d4d6-116">**使用できる場所**: Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-116">**Available in**: Outlook on the web (Classic)</span></span>

### <a name="attachments"></a><span data-ttu-id="1d4d6-117">添付ファイル</span><span class="sxs-lookup"><span data-stu-id="1d4d6-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="1d4d6-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="1d4d6-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="1d4d6-119">添付ファイルのコンテンツを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-119">AttachmentContent - Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="1d4d6-120">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-120">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="1d4d6-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="1d4d6-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="1d4d6-122">メッセージまたは予定に base 64 エンコード文字列として表されるファイルを添付する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-122">Office.context.mailbox.item.addFileAttachmentFromBase64Async - Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="1d4d6-123">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-123">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="1d4d6-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="1d4d6-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

<span data-ttu-id="1d4d6-125">特定の添付ファイルのコンテンツを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-125">Office.context.mailbox.item.getAttachmentContentAsync - Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="1d4d6-126">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-126">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[<span data-ttu-id="1d4d6-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="1d4d6-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

<span data-ttu-id="1d4d6-128">新規作成モードでアイテムの添付ファイルを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-128">Office.context.mailbox.item.getAttachmentsAsync - Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="1d4d6-129">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-129">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="1d4d6-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="1d4d6-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="1d4d6-131">添付ファイルのコンテンツに適用されるフォーマットを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-131">Office.MailboxEnums.AttachmentContentFormat - Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="1d4d6-132">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-132">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="1d4d6-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="1d4d6-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="1d4d6-134">アイテムから添付ファイルが追加されたか、または削除されたかどうかを特定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-134">Office.MailboxEnums.AttachmentStatus - Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="1d4d6-135">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-135">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="1d4d6-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="1d4d6-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1d4d6-137">`AttachmentsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="1d4d6-138">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-138">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="delegate-access"></a><span data-ttu-id="1d4d6-139">代理人アクセス</span><span class="sxs-lookup"><span data-stu-id="1d4d6-139">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="1d4d6-140">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="1d4d6-140">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="1d4d6-141">共有フォルダー、予定表、メールボックスの中の予定やメッセージ アイテムのプロパティを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-141">SharedProperties - Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="1d4d6-142">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-142">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="1d4d6-143">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1d4d6-143">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="1d4d6-144">予定やメッセージ アイテムの sharedProperties を表すオブジェクトを取得する新しい方法が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-144">Office.context.mailbox.item.getSharedPropertiesAsync - Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="1d4d6-145">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-145">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="1d4d6-146">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="1d4d6-146">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="1d4d6-147">代理人のアクセス権を指定する新しいビット フラグ列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-147">Office.MailboxEnums.DelegatePermissions - Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="1d4d6-148">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-148">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="1d4d6-149">SupportsSharedFolders マニフェスト要素</span><span class="sxs-lookup"><span data-stu-id="1d4d6-149">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="1d4d6-150">[DesktopFormFactor](../../manifest/desktopformfactor.md) マニフェスト要素に子要素が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-150">SupportsSharedFolders manifest element - Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="1d4d6-151">代理人のシナリオでアドインが使用できるかどうかを定義します。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-151">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="1d4d6-152">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-152">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="enhanced-location"></a><span data-ttu-id="1d4d6-153">強化された場所</span><span class="sxs-lookup"><span data-stu-id="1d4d6-153">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="1d4d6-154">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="1d4d6-154">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="1d4d6-155">予定の場所のセットを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-155">EnhancedLocation - Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="1d4d6-156">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-156">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="1d4d6-157">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="1d4d6-157">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="1d4d6-158">場所を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-158">LocationDetails - Added a new object that represents a location.</span></span> <span data-ttu-id="1d4d6-159">読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-159">Read only.</span></span>

<span data-ttu-id="1d4d6-160">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-160">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="1d4d6-161">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="1d4d6-161">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="1d4d6-162">場所の ID を表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-162">LocationIdentifier - Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="1d4d6-163">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-163">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="1d4d6-164">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="1d4d6-164">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

<span data-ttu-id="1d4d6-165">予定の場所のセットを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-165">Office.context.mailbox.item.enhancedLocation - Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="1d4d6-166">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-166">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="1d4d6-167">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="1d4d6-167">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="1d4d6-168">予定の場所の種類を指定する新しい列挙型が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-168">Office.MailboxEnums.LocationType - Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="1d4d6-169">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-169">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="1d4d6-170">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="1d4d6-170">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1d4d6-171">`EnhancedLocationsChanged` イベントが `Item` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-171">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="1d4d6-172">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-172">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="1d4d6-173">操作可能なメッセージとの統合</span><span class="sxs-lookup"><span data-stu-id="1d4d6-173">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="1d4d6-174">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="1d4d6-174">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="1d4d6-175">アドインが[操作可能メッセージによってアクティブ化](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)されるときに渡される初期化データを返す新しい関数が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-175">Office.context.mailbox.item.getInitializationContextAsync - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="1d4d6-176">**使用できる場所**: Office 2019 for Windows (Office 365 サブスクリプション)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-176">**Available in**: Office 2019 for Windows (Office 365 subscription), Outlook on the web (Classic)</span></span>

### <a name="internet-headers"></a><span data-ttu-id="1d4d6-177">インターネット ヘッダー</span><span class="sxs-lookup"><span data-stu-id="1d4d6-177">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="1d4d6-178">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="1d4d6-178">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="1d4d6-179">メッセージ アイテムのインターネット ヘッダーを表す新しいオブジェクトが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-179">InternetHeaders - Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="1d4d6-180">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-180">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="1d4d6-181">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="1d4d6-181">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

<span data-ttu-id="1d4d6-182">メッセージ アイテムのインターネット ヘッダーを表す新しいプロパティが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-182">Office.context.mailbox.item.internetHeaders - Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="1d4d6-183">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-183">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="office-theme"></a><span data-ttu-id="1d4d6-184">Office テーマ</span><span class="sxs-lookup"><span data-stu-id="1d4d6-184">Office Theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="1d4d6-185">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="1d4d6-185">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="1d4d6-186">Office テーマを取得する機能が追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="1d4d6-187">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-187">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="1d4d6-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="1d4d6-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1d4d6-189">`OfficeThemeChanged` イベントが `Mailbox` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="1d4d6-190">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-190">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="sso"></a><span data-ttu-id="1d4d6-191">SSO</span><span class="sxs-lookup"><span data-stu-id="1d4d6-191">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="1d4d6-192">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="1d4d6-192">Office.context.auth.getAccessTokenAsync</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="1d4d6-193">Microsoft Graph API の[アクセス トークンの取得](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)をアドインに対して許可する、`getAccessTokenAsync` へのアクセスが追加されました。</span><span class="sxs-lookup"><span data-stu-id="1d4d6-193">Office.context.auth.getAccessTokenAsync - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="1d4d6-194">**使用できる場所**: Windows 用 Outlook 2019 (Office 365 サブスクリプション)、Mac 用 Outlook 2019、Outlook on the web (Office 365 および Outlook.com)、Outlook on the web (クラシック)</span><span class="sxs-lookup"><span data-stu-id="1d4d6-194">**Available in**: Outlook 2019 for Windows (Office 365 subscription), Outlook 2019 for Mac, Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="1d4d6-195">関連項目</span><span class="sxs-lookup"><span data-stu-id="1d4d6-195">See also</span></span>

- [<span data-ttu-id="1d4d6-196">Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="1d4d6-196">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="1d4d6-197">Outlook アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="1d4d6-197">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1d4d6-198">作業の開始</span><span class="sxs-lookup"><span data-stu-id="1d4d6-198">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
