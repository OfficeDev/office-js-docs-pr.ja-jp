---
title: Outlook アドインのアクティブ化ルール
description: Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349022"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="2502d-103">Outlook コンテキスト アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="2502d-p101">Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。これは、1.1 マニフェストのスキーマを使用するすべてのアドインについて同様です。ユーザーは、Outlook UI からアドインを選び、現在のアイテムに、そのアドインを起動することができます。</span><span class="sxs-lookup"><span data-stu-id="2502d-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="2502d-107">次の図は、閲覧ウィンドウにあるアドイン バーでアクティブ化されたメッセージ用の Outlook アドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="2502d-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![アクティブ化された読み取りメール アプリを表示するアプリ バー。](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="2502d-109">マニフェストでのアクティブ化ルールの指定</span><span class="sxs-lookup"><span data-stu-id="2502d-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="2502d-110">特定のOutlookをアクティブ化するには、次のいずれかの要素を使用して、アドイン マニフェストでアクティブ化ルールを指定 `Rule` します。</span><span class="sxs-lookup"><span data-stu-id="2502d-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements.</span></span>

- <span data-ttu-id="2502d-111">[Rule 要素 (MailApp complexType)](../reference/manifest/rule.md) - 個別のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="2502d-112">[Rule 要素 (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - 論理演算子を使用して複数のルールを結合します。</span><span class="sxs-lookup"><span data-stu-id="2502d-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>


 > [!NOTE]
 > <span data-ttu-id="2502d-113">個々 `Rule` のルールを指定するために使用する要素は、抽象 [Rule](../reference/manifest/rule.md) 複合型です。</span><span class="sxs-lookup"><span data-stu-id="2502d-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="2502d-114">次の各種類のルールは、この抽象複合型 `Rule` を拡張します。</span><span class="sxs-lookup"><span data-stu-id="2502d-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="2502d-115">したがって、マニフェストで個別のルールを指定するときは、[xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性を使用してルールの以下の型の 1 つをさらに定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2502d-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="2502d-116">たとえば、次のルールは [ItemIs ルールを定義](../reference/manifest/rule.md#itemis-rule) します。</span><span class="sxs-lookup"><span data-stu-id="2502d-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule.</span></span>
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > <span data-ttu-id="2502d-117">属性はマニフェスト v1.1 のアクティブ化ルールに適用されますが `FormType` `VersionOverrides` 、v1.0 では定義されていません。</span><span class="sxs-lookup"><span data-stu-id="2502d-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="2502d-118">したがって [、ItemIs](../reference/manifest/rule.md#itemis-rule) がノードで使用されている場合は使用 `VersionOverrides` できません。</span><span class="sxs-lookup"><span data-stu-id="2502d-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="2502d-p105">次の表は、使用できるルールの種類を示しています。詳細については、この表の後の説明と、「[閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)」の該当記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2502d-p105">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="2502d-121">**ルール名**</span><span class="sxs-lookup"><span data-stu-id="2502d-121">**Rule name**</span></span>|<span data-ttu-id="2502d-122">**該当するフォーム**</span><span class="sxs-lookup"><span data-stu-id="2502d-122">**Applicable forms**</span></span>|<span data-ttu-id="2502d-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="2502d-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="2502d-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="2502d-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="2502d-125">読み取り、作成</span><span class="sxs-lookup"><span data-stu-id="2502d-125">Read, Compose</span></span>|<span data-ttu-id="2502d-p106">現在選択されているアイテムは指定された種類のアイテム (メッセージまたは予定) かどうかを調べます。また、アイテム クラス、フォームの種類、さらにはオプションでアイテム メッセージ クラスも調べることができます。</span><span class="sxs-lookup"><span data-stu-id="2502d-p106">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="2502d-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="2502d-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="2502d-129">読み取り</span><span class="sxs-lookup"><span data-stu-id="2502d-129">Read</span></span>|<span data-ttu-id="2502d-130">選択されているアイテムに添付ファイルが含まれるかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="2502d-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="2502d-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="2502d-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="2502d-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="2502d-132">Read</span></span>|<span data-ttu-id="2502d-p107">選択されているアイテムに 1 つ以上の一般的なエンティティが含まれるかどうかを調べます。詳細: 「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」。</span><span class="sxs-lookup"><span data-stu-id="2502d-p107">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="2502d-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="2502d-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="2502d-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="2502d-136">Read</span></span>|<span data-ttu-id="2502d-137">選択されているアイテムの送信者のメール アドレス、件名、本文に正規表現と一致するものが含まれるかどうかを調べます。詳細: [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="2502d-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="2502d-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="2502d-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="2502d-139">読み取り、作成</span><span class="sxs-lookup"><span data-stu-id="2502d-139">Read, Compose</span></span>|<span data-ttu-id="2502d-140">複数のルールを組み合わせて、より複雑なルールを作成できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="2502d-141">ItemIs ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-141">ItemIs rule</span></span>

<span data-ttu-id="2502d-142">**ItemIs** 複合型は、現在のアイテムがアイテムの種類と一致している場合 (また、オプションとしてルールに明記されている場合はアイテムのメッセージ クラスとも一致している場合) に **true** と評価されるルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="2502d-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="2502d-143">ItemIs ルールの属性で、次のいずれかの `ItemType` アイテムの種類 **を指定** します。</span><span class="sxs-lookup"><span data-stu-id="2502d-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="2502d-144">マニフェストでは、複数の **ItemIs** ルールを指定できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="2502d-145">ItemType simpleType では、Outlook アドインをサポートしている Outlook アイテムの種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="2502d-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="2502d-146">**値**</span><span class="sxs-lookup"><span data-stu-id="2502d-146">**Value**</span></span>|<span data-ttu-id="2502d-147">**説明**</span><span class="sxs-lookup"><span data-stu-id="2502d-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="2502d-148">**Appointment**</span><span class="sxs-lookup"><span data-stu-id="2502d-148">**Appointment**</span></span>|<span data-ttu-id="2502d-149">Outlook の予定表内のアイテムを指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-149">Specifies an item in an Outlook calendar.</span></span> <span data-ttu-id="2502d-150">このアイテムには、開催者と出席者を持つ応答済みの会議アイテムと、開催者と出席者を持たない、単なる予定表上のアイテムである予定が含まれます。</span><span class="sxs-lookup"><span data-stu-id="2502d-150">This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.</span></span> <span data-ttu-id="2502d-151">これは Outlook の IPM.Appointment メッセージ クラスに対応します。</span><span class="sxs-lookup"><span data-stu-id="2502d-151">This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="2502d-152">**メッセージ**</span><span class="sxs-lookup"><span data-stu-id="2502d-152">**Message**</span></span>|<span data-ttu-id="2502d-153">通常受信トレイで受信される次のいずれかの項目を指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-153">Specifies one of the following items received in typically the Inbox.</span></span> <ul><li><p><span data-ttu-id="2502d-p110">電子メール メッセージ。これは Outlook の IPM.Note メッセージ クラスに対応します。</span><span class="sxs-lookup"><span data-stu-id="2502d-p110">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="2502d-156">会議出席依頼、返信、または取り消し。</span><span class="sxs-lookup"><span data-stu-id="2502d-156">A meeting request, response, or cancellation.</span></span> <span data-ttu-id="2502d-157">これは、次のメッセージ クラスに対応Outlook。</span><span class="sxs-lookup"><span data-stu-id="2502d-157">This corresponds to the following message classes in Outlook.</span></span></p><p><span data-ttu-id="2502d-158">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="2502d-158">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="2502d-159">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="2502d-159">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="2502d-160">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="2502d-160">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="2502d-161">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="2502d-161">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="2502d-162">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="2502d-162">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="2502d-163">この属性を使用して、アドインをアクティブにするモード (読み取りまたは作成 `FormType` ) を指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-163">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="2502d-164">ItemIs `FormType` 属性はスキーマ v1.1 以降で定義されますが `VersionOverrides` 、v1.0 では定義されません。</span><span class="sxs-lookup"><span data-stu-id="2502d-164">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="2502d-165">アドイン コマンドを定義 `FormType` するときに属性を含めない。</span><span class="sxs-lookup"><span data-stu-id="2502d-165">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="2502d-166">アドインがアクティブ化された後は、 [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) プロパティを使用して Outlook で現在選択されているアイテムを取得し、 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティを使用して現在のアイテムの種類を取得できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-166">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="2502d-167">必要に応じて、属性を使用してアイテムのメッセージ クラスを指定し、属性を使用して、アイテムが指定されたクラスのサブクラスである場合にルールを true にするかどうかを `ItemClass` `IncludeSubClasses` 指定できます。 </span><span class="sxs-lookup"><span data-stu-id="2502d-167">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="2502d-168">メッセージ クラスの詳細については、「[Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="2502d-168">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="2502d-169">次の例は、ユーザーがメッセージを読み取っているときに、アドイン バー Outlookアドインを表示できる **ItemIs** ルールです。</span><span class="sxs-lookup"><span data-stu-id="2502d-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message.</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="2502d-170">次の例は、ユーザーがメッセージまたは予約を閲覧するときに Outlook のアドイン バーにアドインを表示する **ItemIs** ルールを示しています。</span><span class="sxs-lookup"><span data-stu-id="2502d-170">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="2502d-171">ItemHasAttachment ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-171">ItemHasAttachment rule</span></span>


<span data-ttu-id="2502d-172">複合 `ItemHasAttachment` 型は、選択したアイテムに添付ファイルが含まれている場合にチェックするルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="2502d-172">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="2502d-173">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-173">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="2502d-174">アイテムをアドインで使用できる前に、サーバーはアドインを調べて、件名と本文に既知のエンティティの 1 つである可能性があるテキストが含まれているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="2502d-174">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities.</span></span> <span data-ttu-id="2502d-175">これらのエンティティが見つかった場合は、そのアイテムの or メソッドを使用してアクセスする既知のエンティティのコレクション `getEntities` `getEntitiesByType` に配置されます。</span><span class="sxs-lookup"><span data-stu-id="2502d-175">If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="2502d-176">指定した種類のエンティティがアイテムに存在する場合にアドインを表示するルールを `ItemHasKnownEntity` 使用して指定できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-176">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item.</span></span> <span data-ttu-id="2502d-177">ルールの属性には、次の既知の `EntityType` エンティティを指定 `ItemHasKnownEntity` できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-177">You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule.</span></span>

- <span data-ttu-id="2502d-178">Address</span><span class="sxs-lookup"><span data-stu-id="2502d-178">Address</span></span>
- <span data-ttu-id="2502d-179">Contact</span><span class="sxs-lookup"><span data-stu-id="2502d-179">Contact</span></span>
- <span data-ttu-id="2502d-180">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="2502d-180">EmailAddress</span></span>
- <span data-ttu-id="2502d-181">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2502d-181">MeetingSuggestion</span></span>
- <span data-ttu-id="2502d-182">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2502d-182">PhoneNumber</span></span>
- <span data-ttu-id="2502d-183">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2502d-183">TaskSuggestion</span></span>
- <span data-ttu-id="2502d-184">URL</span><span class="sxs-lookup"><span data-stu-id="2502d-184">URL</span></span>

<span data-ttu-id="2502d-185">必要に応じて、属性に正規表現を含め、現在の正規表現と一致するエンティティの場合にのみアドイン `RegularExpression` が表示されます。</span><span class="sxs-lookup"><span data-stu-id="2502d-185">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present.</span></span> <span data-ttu-id="2502d-186">ルールで指定された正規表現に一致する文字列を取得するには、現在選択されているアイテムアイテムに `ItemHasKnownEntity` `getRegExMatches` or `getFilteredEntitiesByName` メソッドOutlookできます。</span><span class="sxs-lookup"><span data-stu-id="2502d-186">To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="2502d-187">次の例は、指定された既知のエンティティの 1 つがメッセージに存在する場合にアドインを表示する要素の `Rule` コレクションを示しています。</span><span class="sxs-lookup"><span data-stu-id="2502d-187">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="2502d-188">次の例は、"contoso" という単語を含む URL がメッセージ内に存在する場合にアドインをアクティブ化する属性を持つルール `ItemHasKnownEntity` `RegularExpression` を示しています。</span><span class="sxs-lookup"><span data-stu-id="2502d-188">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="2502d-189">アクティブ化ルールのエンティティの詳細については、「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2502d-189">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="2502d-190">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-190">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="2502d-191">複合型は、正規表現を使用してアイテムの指定されたプロパティの内容と一致 `ItemHasRegularExpressionMatch` するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="2502d-191">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item.</span></span> <span data-ttu-id="2502d-192">正規表現に一致するテキストがアイテムの指定プロパティ内に見つかった場合に、Outlook はアドイン バーをアクティブ化してそのアドインを表示します。</span><span class="sxs-lookup"><span data-stu-id="2502d-192">If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in.</span></span> <span data-ttu-id="2502d-193">現在選択されているアイテムを表すオブジェクトの or メソッドを使用して、指定した正規表現の `getRegExMatches` `getRegExMatchesByName` 一致を取得できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-193">You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="2502d-194">次の例は、選択したアイテムの本文に大文字と小文字を無視して、"apple"、"banana"、または "ココナッツ" が含まれている場合にアドインをアクティブ化する例 `ItemHasRegularExpressionMatch` を示しています。</span><span class="sxs-lookup"><span data-stu-id="2502d-194">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="2502d-195">ルールの使用の詳細については `ItemHasRegularExpressionMatch` [、「Use regular expression activation rules to show a Outlookアドイン」を参照してください](use-regular-expressions-to-show-an-outlook-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="2502d-195">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="2502d-196">RuleCollection ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-196">RuleCollection rule</span></span>


<span data-ttu-id="2502d-197">複合 `RuleCollection` 型は、複数のルールを 1 つのルールに結合します。</span><span class="sxs-lookup"><span data-stu-id="2502d-197">The `RuleCollection` complex type combines multiple rules into a single rule.</span></span> <span data-ttu-id="2502d-198">属性を使用して、コレクション内のルールを論理 OR または論理 AND と組み合わせるかどうかを指定 `Mode` できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-198">You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="2502d-p118">論理 AND を指定する場合、アドインは、コレクション内で指定されているすべてのルールにアイテムが一致する場合にのみ表示されます。論理 OR を指定する場合は、コレクションで指定されているルールのいずれか 1 つにでもアイテムが一致すれば、アドインは表示されます。</span><span class="sxs-lookup"><span data-stu-id="2502d-p118">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="2502d-201">ルールを組み `RuleCollection` 合わせて複雑なルールを形成できます。</span><span class="sxs-lookup"><span data-stu-id="2502d-201">You can combine `RuleCollection` rules to form complex rules.</span></span> <span data-ttu-id="2502d-202">次に示す例では、件名や本文に住所が含まれるメッセージまたは予定表のアイテムをユーザーが表示したときに、アドインがアクティブ化されます。</span><span class="sxs-lookup"><span data-stu-id="2502d-202">The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="2502d-203">次の例では、ユーザーがメッセージを新規作成するときか、件名か本文に住所が含まれる予定を表示するときに、アドインがアクティブ化されます。</span><span class="sxs-lookup"><span data-stu-id="2502d-203">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="2502d-204">ルールと正規表現の制約事項</span><span class="sxs-lookup"><span data-stu-id="2502d-204">Limits for rules and regular expressions</span></span>


<span data-ttu-id="2502d-205">アドインを十分にOutlookするには、ライセンス認証と API の使用ガイドラインに従う必要があります。</span><span class="sxs-lookup"><span data-stu-id="2502d-205">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines.</span></span> <span data-ttu-id="2502d-206">次の表に、正規表現とルールの一般的な制限を示しますが、アプリケーションごとに特定のルールがあります。</span><span class="sxs-lookup"><span data-stu-id="2502d-206">The following table shows general limits for regular expressions and rules but there are specific rules for different applications.</span></span> <span data-ttu-id="2502d-207">詳細については、「ライセンス認証の制限」および[「JavaScript API for Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) Outlookアドインのライセンス認証のトラブルシューティング」を参照[してください](troubleshoot-outlook-add-in-activation.md)。</span><span class="sxs-lookup"><span data-stu-id="2502d-207">For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="2502d-208">**アドインの要素**</span><span class="sxs-lookup"><span data-stu-id="2502d-208">**Add-in element**</span></span>|<span data-ttu-id="2502d-209">**ガイドライン**</span><span class="sxs-lookup"><span data-stu-id="2502d-209">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="2502d-210">マニフェストのサイズ</span><span class="sxs-lookup"><span data-stu-id="2502d-210">Manifest Size</span></span>|<span data-ttu-id="2502d-211">256 KB 未満。</span><span class="sxs-lookup"><span data-stu-id="2502d-211">No larger than 256 KB.</span></span>|
|<span data-ttu-id="2502d-212">ルール</span><span class="sxs-lookup"><span data-stu-id="2502d-212">Rules</span></span>|<span data-ttu-id="2502d-213">15 ルール未満。</span><span class="sxs-lookup"><span data-stu-id="2502d-213">No more than 15 rules.</span></span>|
|<span data-ttu-id="2502d-214">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="2502d-214">ItemHasKnownEntity</span></span>|<span data-ttu-id="2502d-215">Outlook リッチ クライアントでは、本文の最初の 1 MB にルールを適用し、残りの部分には適用しません。</span><span class="sxs-lookup"><span data-stu-id="2502d-215">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="2502d-216">正規表現</span><span class="sxs-lookup"><span data-stu-id="2502d-216">Regular Expressions</span></span>|<span data-ttu-id="2502d-217">すべてのアプリケーションの ItemHasKnownEntity ルールまたは ItemHasRegularExpressionMatch ルールOutlookします。</span><span class="sxs-lookup"><span data-stu-id="2502d-217">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="2502d-p121">Outlook アドインのアクティブ化ルールで指定する正規表現は 5 個までにしてください。その制約数を超えるアドインをインストールすることはできません。</span><span class="sxs-lookup"><span data-stu-id="2502d-p121">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="2502d-220">予期される結果が <b>getRegExMatches</b> メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-220">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="2502d-221">正規表現で先読みアサーションは指定しますが、後読み `(?<=text)` および否定の後読み `(?<!text)` アサーションは指定しません。</span><span class="sxs-lookup"><span data-stu-id="2502d-221">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="2502d-222">一致数が次の表の制限を超えない正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="2502d-222">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="2502d-223">正規表現の長さ制限</span><span class="sxs-lookup"><span data-stu-id="2502d-223">Limit on length of a regex match</span></span></th><th><span data-ttu-id="2502d-224">Outlook リッチ クライアント</span><span class="sxs-lookup"><span data-stu-id="2502d-224">Outlook rich clients</span></span></th><th><span data-ttu-id="2502d-225">iOS および Android 用の Outlook</span><span class="sxs-lookup"><span data-stu-id="2502d-225">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="2502d-226">アイテムの本文がテキスト形式の場合</span><span class="sxs-lookup"><span data-stu-id="2502d-226">Item body is plain text</span></span></td><td><span data-ttu-id="2502d-227">1.5 KB</span><span class="sxs-lookup"><span data-stu-id="2502d-227">1.5 KB</span></span></td><td><span data-ttu-id="2502d-228">3 KB</span><span class="sxs-lookup"><span data-stu-id="2502d-228">3 KB</span></span></td></tr><tr><td><span data-ttu-id="2502d-229">アイテムの本文が HTML の場合</span><span class="sxs-lookup"><span data-stu-id="2502d-229">Item body it HTML</span></span></td><td><span data-ttu-id="2502d-230">3 KB</span><span class="sxs-lookup"><span data-stu-id="2502d-230">3 KB</span></span></td><td><span data-ttu-id="2502d-231">3 KB</span><span class="sxs-lookup"><span data-stu-id="2502d-231">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="2502d-232">関連項目</span><span class="sxs-lookup"><span data-stu-id="2502d-232">See also</span></span>

- [<span data-ttu-id="2502d-233">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="2502d-233">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="2502d-234">Outlook アドインのアクティブ化と JavaScript API の制限</span><span class="sxs-lookup"><span data-stu-id="2502d-234">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="2502d-235">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="2502d-235">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="2502d-236">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="2502d-236">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
