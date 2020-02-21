---
title: Outlook アドインのアクセス許可を理解する
description: Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは Restricted、ReadItem、ReadWriteItem、ReadWriteMailbox です。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 58d21a33034475b8c33b8449ece24c9dafc84e2b
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166447"
---
# <a name="understanding-outlook-add-in-permissions"></a><span data-ttu-id="b91d1-103">Outlook アドインのアクセス許可を理解する</span><span class="sxs-lookup"><span data-stu-id="b91d1-103">Understanding Outlook add-in permissions</span></span>

<span data-ttu-id="b91d1-p101">Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは **Restricted**、**ReadItem**、**ReadWriteItem**、**ReadWriteMailbox** です。これらのレベルのアクセス許可は累積されます。**Restricted** は最低レベルであり、それぞれの上位レベルには、下位レベルのアクセス許可がすべて含まれます。**ReadWriteMailbox** にはサポートされるアクセス許可がすべて含まれます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p101">Outlook add-ins specify the required permission level in their manifest. The available levels are **Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**. These levels of permissions are cumulative: **Restricted** is the lowest level, and each higher level includes the permissions of all the lower levels. **ReadWriteMailbox** includes all the supported permissions.</span></span>

<span data-ttu-id="b91d1-p102">メール アドインが要求するアクセス許可を、[AppSource](https://appsource.microsoft.com) からメール アドインをインストールする前に表示できます。Exchange 管理センターで、インストールしたアドインに必要なアクセス許可を表示することもできます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p102">You can see the permissions requested by a mail add-in before installing it from [AppSource](https://appsource.microsoft.com). You can also see the required permissions of installed add-ins in the Exchange Admin Center.</span></span>

## <a name="restricted-permission"></a><span data-ttu-id="b91d1-110">制限付きアクセス許可</span><span class="sxs-lookup"><span data-stu-id="b91d1-110">Restricted permission</span></span>

<span data-ttu-id="b91d1-p103">
  \*\*Restricted\*\* アクセス許可は、最も基本的なアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの [Permissions](../reference/manifest/permissions.md) 要素内で \*\*Restricted\*\* を指定します。メール アドインがマニフェストで特定のアクセス許可を要求していない場合、Outlook は既定でこのアクセス許可をそのアドインに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p103">The **Restricted** permission is the most basic level of permission. Specify **Restricted** in the [Permissions](../reference/manifest/permissions.md) element in the manifest to request this permission. Outlook assigns this permission to a mail add-in by default if the add-in does not request a specific permission in its manifest.</span></span>

### <a name="can-do"></a><span data-ttu-id="b91d1-114">できること</span><span class="sxs-lookup"><span data-stu-id="b91d1-114">Can do</span></span>

- <span data-ttu-id="b91d1-115">アイテムの件名または本文から [特定のエンティティのみを取得](match-strings-in-an-item-as-well-known-entities.md) (電話番号、アドレス、URL)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-115">[Get only specific entities](match-strings-in-an-item-as-well-known-entities.md) (phone number, address, URL) from the item's subject or body.</span></span>

- <span data-ttu-id="b91d1-116">閲覧フォームまたは新規作成フォームの現在のアイテムが特定のアイテムの種類であることを要求する [ItemIs アクティブ化ルール](activation-rules.md#itemis-rule)を指定、または、選択したアイテムでサポートされる既知のエンティティ (電話番号、アドレス、URL) の小さなサブセットに一致する [ItemHasKnownEntity ルール](match-strings-in-an-item-as-well-known-entities.md)を指定。</span><span class="sxs-lookup"><span data-stu-id="b91d1-116">Specify an [ItemIs activation rule](activation-rules.md#itemis-rule) that requires the current item in a read or compose form to be a specific item type, or [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) that matches any of a smaller subset of supported well-known entities (phone number, address, URL) in the selected item.</span></span>

- <span data-ttu-id="b91d1-117">ユーザーまたはアイテムに関する特定の情報に関連**しない**プロパティとメソッドへのアクセス (これを実行するメンバーのリストは、次のセクションを参照)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-117">Access any properties and methods that do **not** pertain to specific information about the user or item (see the next section for the list of members that do).</span></span>

### <a name="cant-do"></a><span data-ttu-id="b91d1-118">できないこと</span><span class="sxs-lookup"><span data-stu-id="b91d1-118">Can't do</span></span>

- <span data-ttu-id="b91d1-119">連絡先、電子メール アドレス、会議の提案、タスクの提案のエンティティで [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用。</span><span class="sxs-lookup"><span data-stu-id="b91d1-119">Use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule on the contact, email address, meeting suggestion, or task suggestion entitiy.</span></span>

- <span data-ttu-id="b91d1-120">[ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) ルールまたは [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ルールを使用。</span><span class="sxs-lookup"><span data-stu-id="b91d1-120">Use the [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) or [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule.</span></span>

- <span data-ttu-id="b91d1-p104">ユーザーまたはアイテムの情報に関連する次のリストに示すメンバーへのアクセス。このリストのメンバーにアクセスしようとすると、**null** が返され、Outlook がメール アドインにアクセス許可の引き上げを要求していることを伝えるエラー メッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p104">Access the members in the following list that pertain to the information of the user or item. Attempting to access members in this list will return **null** and result in an error message which states that Outlook requires the mail add-in to have elevated permission.</span></span>

    - [<span data-ttu-id="b91d1-123">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-123">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-124">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-124">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-125">item.attachments</span><span class="sxs-lookup"><span data-stu-id="b91d1-125">item.attachments</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-126">item.bcc</span><span class="sxs-lookup"><span data-stu-id="b91d1-126">item.bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-127">item.body</span><span class="sxs-lookup"><span data-stu-id="b91d1-127">item.body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-128">item.cc</span><span class="sxs-lookup"><span data-stu-id="b91d1-128">item.cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-129">item.from</span><span class="sxs-lookup"><span data-stu-id="b91d1-129">item.from</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-130">item.getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="b91d1-130">item.getRegExMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-131">item.getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b91d1-131">item.getRegExMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-132">item.optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="b91d1-132">item.optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-133">item.organizer</span><span class="sxs-lookup"><span data-stu-id="b91d1-133">item.organizer</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-134">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-134">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-135">item.requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="b91d1-135">item.requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-136">item.sender</span><span class="sxs-lookup"><span data-stu-id="b91d1-136">item.sender</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-137">item.to</span><span class="sxs-lookup"><span data-stu-id="b91d1-137">item.to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [<span data-ttu-id="b91d1-138">mailbox.getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-138">mailbox.getCallbackTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="b91d1-139">mailbox.getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-139">mailbox.getUserIdentityTokenAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="b91d1-140">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-140">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="b91d1-141">mailbox.userProfile</span><span class="sxs-lookup"><span data-stu-id="b91d1-141">mailbox.userProfile</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - <span data-ttu-id="b91d1-142">[Body](/javascript/api/outlook/office.body) およびその子メンバーすべて</span><span class="sxs-lookup"><span data-stu-id="b91d1-142">[Body](/javascript/api/outlook/office.body) and all its child members</span></span>
    - <span data-ttu-id="b91d1-143">[Location](/javascript/api/outlook/office.location) およびその子メンバーすべて</span><span class="sxs-lookup"><span data-stu-id="b91d1-143">[Location](/javascript/api/outlook/office.location) and all its child members</span></span>
    - <span data-ttu-id="b91d1-144">[Recipients](/javascript/api/outlook/office.recipients) およびその子メンバーすべて</span><span class="sxs-lookup"><span data-stu-id="b91d1-144">[Recipients](/javascript/api/outlook/office.recipients) and all its child members</span></span>
    - <span data-ttu-id="b91d1-145">[Subject](/javascript/api/outlook/office.subject) およびその子メンバーすべて</span><span class="sxs-lookup"><span data-stu-id="b91d1-145">[Subject](/javascript/api/outlook/office.subject) and all its child members</span></span>
    - <span data-ttu-id="b91d1-146">[Time](/javascript/api/outlook/office.time) およびその子メンバーすべて</span><span class="sxs-lookup"><span data-stu-id="b91d1-146">[Time](/javascript/api/outlook/office.time) and all its child members</span></span>

## <a name="readitem-permission"></a><span data-ttu-id="b91d1-147">ReadItem アクセス許可</span><span class="sxs-lookup"><span data-stu-id="b91d1-147">ReadItem permission</span></span>

<span data-ttu-id="b91d1-p105">**ReadItem** アクセス許可は、アクセス許可モデルの中でその次に位置するアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの **Permissions** 要素内で **ReadItem** を指定します。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p105">The **ReadItem** permission is the next level of permission in the permissions model. Specify **ReadItem** in the **Permissions** element in the manifest to request this permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="b91d1-150">できること</span><span class="sxs-lookup"><span data-stu-id="b91d1-150">Can do</span></span>

- <span data-ttu-id="b91d1-151">閲覧フォームまたは [新規作成フォーム](item-data.md)の現在のアイテムの [すべてのプロパティの読み取り](get-and-set-item-data-in-a-compose-form.md)。たとえば、閲覧フォームの [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) および新規作成フォームの [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-151">[Read all the properties](item-data.md) of the current item in a read or [compose form](get-and-set-item-data-in-a-compose-form.md), for example, [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) in a read form and [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-) in a compose form.</span></span>

- <span data-ttu-id="b91d1-152">Exchange Web Services (EWS) または [Outlook REST API](use-rest-api.md) で[アイテムの添付ファイルを取得する](get-attachments-of-an-outlook-item.md)か、アイテム全体を取得するためのコールバック トークンを取得。</span><span class="sxs-lookup"><span data-stu-id="b91d1-152">[Get a callback token to get item attachments](get-attachments-of-an-outlook-item.md) or the full item with Exchange Web Services (EWS) or [Outlook REST APIs](use-rest-api.md).</span></span>

- <span data-ttu-id="b91d1-153">そのアイテムのアドインが設定する[カスタム プロパティの書き込み](/javascript/api/outlook/office.CustomProperties)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-153">[Write custom properties](/javascript/api/outlook/office.CustomProperties) set by the add-in on that item.</span></span>

- <span data-ttu-id="b91d1-154">アイテムの件名または本文から、サブセットだけでなく、[存在する既知のエンティティをすべて取得する](match-strings-in-an-item-as-well-known-entities.md)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-154">[Get all existing well-known entities](match-strings-in-an-item-as-well-known-entities.md), not just a subset, from the item's subject or body.</span></span>

- <span data-ttu-id="b91d1-p106">
  [ItemHasKnownEntity](activation-rules.md#itemhasknownentity-rule) ルールの [既知のエンティティ](../reference/manifest/rule.md#itemhasknownentity-rule)、または [ItemHasRegularExpressionMatch](activation-rules.md#itemhasregularexpressionmatch-rule) ルールの [正規表現](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule)をすべて使用します。次の例は、スキーマ v1.1 に従っています。選択されたメッセージの件名または本文に既知のエンティティが 1 つ以上ある場合にアクティブ化されるルールを示しています。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p106">Use all the [well-known entities](activation-rules.md#itemhasknownentity-rule) in [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rules, or [regular expressions](activation-rules.md#itemhasregularexpressionmatch-rule) in [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules. The following example follows schema v1.1. It shows a rule that activates the add-in if one or more of the well-known entities are found in the subject or body of the selected message:</span></span>

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a><span data-ttu-id="b91d1-158">できないこと</span><span class="sxs-lookup"><span data-stu-id="b91d1-158">Can't do</span></span>

- <span data-ttu-id="b91d1-159">**mailbox.getCallbackTokenAsync** によって提供されるトークンを次の目的に使用すること。</span><span class="sxs-lookup"><span data-stu-id="b91d1-159">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="b91d1-160">Outlook REST API を使用した現在のアイテムの更新または削除、またはユーザーのメールボックスにあるその他アイテムへのアクセス。</span><span class="sxs-lookup"><span data-stu-id="b91d1-160">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="b91d1-161">Outlook REST API を使用した現在の予定表イベント アイテムの取得。</span><span class="sxs-lookup"><span data-stu-id="b91d1-161">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="b91d1-162">次のいずれかの API を使用すること。</span><span class="sxs-lookup"><span data-stu-id="b91d1-162">Use any of the following APIs:</span></span>
    - [<span data-ttu-id="b91d1-163">mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-163">mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [<span data-ttu-id="b91d1-164">item.addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-164">item.addFileAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-165">item.addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-165">item.addItemAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-166">item.bcc.addAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-166">item.bcc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-167">item.bcc.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-167">item.bcc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-168">item.body.prependAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-168">item.body.prependAsync</span></span>](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [<span data-ttu-id="b91d1-169">item.body.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-169">item.body.setAsync</span></span>](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [<span data-ttu-id="b91d1-170">item.body.setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-170">item.body.setSelectedDataAsync</span></span>](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [<span data-ttu-id="b91d1-171">item.cc.addAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-171">item.cc.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-172">item.cc.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-172">item.cc.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-173">item.end.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-173">item.end.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="b91d1-174">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-174">item.location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [<span data-ttu-id="b91d1-175">item.optionalAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-175">item.optionalAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-176">item.optionalAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-176">item.optionalAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-177">item.removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-177">item.removeAttachmentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [<span data-ttu-id="b91d1-178">item.requiredAttendees.addAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-178">item.requiredAttendees.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-179">item.requiredAttendees.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-179">item.requiredAttendees.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-180">item.start.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-180">item.start.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [<span data-ttu-id="b91d1-181">item.subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-181">item.subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [<span data-ttu-id="b91d1-182">item.to.addAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-182">item.to.addAsync</span></span>](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [<span data-ttu-id="b91d1-183">item.to.setAsync</span><span class="sxs-lookup"><span data-stu-id="b91d1-183">item.to.setAsync</span></span>](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a><span data-ttu-id="b91d1-184">ReadWriteItem アクセス許可</span><span class="sxs-lookup"><span data-stu-id="b91d1-184">ReadWriteItem permission</span></span>

<span data-ttu-id="b91d1-p107">マニフェストの **Permissions** 要素に **ReadWriteItem** を指定すると、このアクセス許可を要求できます。作成フォームでアクティブになり、書き込みメソッド (**Message.to.addAsync** または **Message.to.setAsync**) を使用するメール アドインは、このレベル以上のアクセス許可を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p107">Specify **ReadWriteItem** in the **Permissions** element in the manifest to request this permission. Mail add-ins activated in compose forms that use write methods (**Message.to.addAsync** or **Message.to.setAsync**) must use at least this level of permission.</span></span>

### <a name="can-do"></a><span data-ttu-id="b91d1-187">できること</span><span class="sxs-lookup"><span data-stu-id="b91d1-187">Can do</span></span>

- <span data-ttu-id="b91d1-188">Outlook で閲覧または新規作成されているアイテムの[すべてのアイテム レベルのプロパティを読み書き](item-data.md)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-188">[Read and write all item-level properties](item-data.md) of the item that is being viewed or composed in Outlook.</span></span>

- <span data-ttu-id="b91d1-189">そのアイテムで[添付ファイルを追加または削除](add-and-remove-attachments-to-an-item-in-a-compose-form.md)。</span><span class="sxs-lookup"><span data-stu-id="b91d1-189">[Add or remove attachments](add-and-remove-attachments-to-an-item-in-a-compose-form.md) of that item.</span></span>

- <span data-ttu-id="b91d1-190">JavaScript API for Office の中でメール アドインに適用される、**Mailbox.makeEWSRequestAsync** を除く他のすべてのメンバーの使用。</span><span class="sxs-lookup"><span data-stu-id="b91d1-190">Use all other members of the JavaScript API for Office that are applicable to mail add-ins, except **Mailbox.makeEWSRequestAsync**.</span></span>

### <a name="cant-do"></a><span data-ttu-id="b91d1-191">できないこと</span><span class="sxs-lookup"><span data-stu-id="b91d1-191">Can't do</span></span>

- <span data-ttu-id="b91d1-192">**mailbox.getCallbackTokenAsync** によって提供されるトークンを次の目的に使用すること。</span><span class="sxs-lookup"><span data-stu-id="b91d1-192">Use the token provided by **mailbox.getCallbackTokenAsync** to:</span></span>
    - <span data-ttu-id="b91d1-193">Outlook REST API を使用した現在のアイテムの更新または削除、またはユーザーのメールボックスにあるその他アイテムへのアクセス。</span><span class="sxs-lookup"><span data-stu-id="b91d1-193">Update or delete the current item using the Outlook REST API or access any other items in the user's mailbox.</span></span>
    - <span data-ttu-id="b91d1-194">Outlook REST API を使用した現在の予定表イベント アイテムの取得。</span><span class="sxs-lookup"><span data-stu-id="b91d1-194">Get the current calendar event item using the Outlook REST API.</span></span>

- <span data-ttu-id="b91d1-195">**Mailbox.makeEWSRequestAsync** の使用。</span><span class="sxs-lookup"><span data-stu-id="b91d1-195">Use **mailbox.makeEWSRequestAsync**.</span></span>

## <a name="readwritemailbox-permission"></a><span data-ttu-id="b91d1-196">ReadWriteMailbox アクセス許可</span><span class="sxs-lookup"><span data-stu-id="b91d1-196">ReadWriteMailbox permission</span></span>

<span data-ttu-id="b91d1-p108">**ReadWriteMailbox** アクセス許可は、最高のアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの **Permissions** 要素内で **ReadWriteMailbox** を指定します。</span><span class="sxs-lookup"><span data-stu-id="b91d1-p108">The **ReadWriteMailbox** permission is the highest level of permission. Specify **ReadWriteMailbox** in the **Permissions** element in the manifest to request this permission.</span></span>

<span data-ttu-id="b91d1-199">**ReadWriteItem** アクセス許可がサポートする内容に加え、**mailbox.getCallbackTokenAsync** が提供するトークンでは、Exchange Web Services (EWS) 操作または Outlook REST API を使用して以下を行うためのアクセス権が提供されます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-199">In addition to what the **ReadWriteItem** permission supports, the token provided by **mailbox.getCallbackTokenAsync** provides access to use Exchange Web Services (EWS) operations or Outlook REST APIs to do the following:</span></span>

- <span data-ttu-id="b91d1-200">ユーザーのメール ボックスのアイテムのすべてのプロパティの読み取りと書き込み。</span><span class="sxs-lookup"><span data-stu-id="b91d1-200">Read and write all properties of any item in the user's mailbox.</span></span>
- <span data-ttu-id="b91d1-201">そのメール ボックスのフォルダーまたはアイテムの作成、読み取り、書き込み。</span><span class="sxs-lookup"><span data-stu-id="b91d1-201">Create, read, and write to any folder or item in that mailbox.</span></span>
- <span data-ttu-id="b91d1-202">そのメール ボックスからのアイテムの送信。</span><span class="sxs-lookup"><span data-stu-id="b91d1-202">Send an item from that mailbox</span></span>

<span data-ttu-id="b91d1-203">**mailbox.makeEWSRequestAsync** を使用して、次の EWS の操作にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-203">Through **mailbox.makeEWSRequestAsync**, you can access the following EWS operations:</span></span>

- [<span data-ttu-id="b91d1-204">CopyItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-204">CopyItem</span></span>](/exchange/client-developer/web-service-reference/copyitem-operation)
- [<span data-ttu-id="b91d1-205">CreateFolder</span><span class="sxs-lookup"><span data-stu-id="b91d1-205">CreateFolder</span></span>](/exchange/client-developer/web-service-reference/createfolder-operation)
- [<span data-ttu-id="b91d1-206">CreateItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-206">CreateItem</span></span>](/exchange/client-developer/web-service-reference/createitem-operation)
- [<span data-ttu-id="b91d1-207">FindConversation</span><span class="sxs-lookup"><span data-stu-id="b91d1-207">FindConversation</span></span>](/exchange/client-developer/web-service-reference/findconversation-operation)
- [<span data-ttu-id="b91d1-208">FindFolder</span><span class="sxs-lookup"><span data-stu-id="b91d1-208">FindFolder</span></span>](/exchange/client-developer/web-service-reference/findfolder-operation)
- [<span data-ttu-id="b91d1-209">FindItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-209">FindItem</span></span>](/exchange/client-developer/web-service-reference/finditem-operation)
- [<span data-ttu-id="b91d1-210">GetConversationItems</span><span class="sxs-lookup"><span data-stu-id="b91d1-210">GetConversationItems</span></span>](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [<span data-ttu-id="b91d1-211">GetFolder</span><span class="sxs-lookup"><span data-stu-id="b91d1-211">GetFolder</span></span>](/exchange/client-developer/web-service-reference/getfolder-operation)
- [<span data-ttu-id="b91d1-212">GetItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-212">GetItem</span></span>](/exchange/client-developer/web-service-reference/getitem-operation)
- [<span data-ttu-id="b91d1-213">MarkAsJunk</span><span class="sxs-lookup"><span data-stu-id="b91d1-213">MarkAsJunk</span></span>](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [<span data-ttu-id="b91d1-214">MoveItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-214">MoveItem</span></span>](/exchange/client-developer/web-service-reference/moveitem-operation)
- [<span data-ttu-id="b91d1-215">SendItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-215">SendItem</span></span>](/exchange/client-developer/web-service-reference/senditem-operation)
- [<span data-ttu-id="b91d1-216">UpdateFolder</span><span class="sxs-lookup"><span data-stu-id="b91d1-216">UpdateFolder</span></span>](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [<span data-ttu-id="b91d1-217">UpdateItem</span><span class="sxs-lookup"><span data-stu-id="b91d1-217">UpdateItem</span></span>](/exchange/client-developer/web-service-reference/updateitem-operation)

<span data-ttu-id="b91d1-218">サポートされていない操作を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="b91d1-218">Attempting to use an unsupported operation will result in an error response.</span></span>

## <a name="see-also"></a><span data-ttu-id="b91d1-219">関連項目</span><span class="sxs-lookup"><span data-stu-id="b91d1-219">See also</span></span>

- [<span data-ttu-id="b91d1-220">Outlook アドインに関するプライバシー、アクセス許可、セキュリティ</span><span class="sxs-lookup"><span data-stu-id="b91d1-220">Privacy, permissions, and security for Outlook add-ins</span></span>](../develop/privacy-and-security.md)
- [<span data-ttu-id="b91d1-221">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="b91d1-221">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
