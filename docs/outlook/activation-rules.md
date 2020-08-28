---
title: Outlook アドインのアクティブ化ルール
description: Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 7a3ed48f77146a25725d46b3e06296cb0eb5616a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294053"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a><span data-ttu-id="5146a-103">Outlook コンテキスト アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-103">Activation rules for contextual Outlook add-ins</span></span>

<span data-ttu-id="5146a-p101">Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。これは、1.1 マニフェストのスキーマを使用するすべてのアドインについて同様です。ユーザーは、Outlook UI からアドインを選び、現在のアイテムに、そのアドインを起動することができます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p101">Outlook activates some types of add-ins if the message or appointment that the user is reading or composing satisfies the activation rules of the add-in. This is true for all add-ins that use the 1.1 manifest schema. The user can then choose the add-in from the Outlook UI to start it for the current item.</span></span>

<span data-ttu-id="5146a-107">次の図は、閲覧ウィンドウにあるアドイン バーでアクティブ化されたメッセージ用の Outlook アドインを示しています。</span><span class="sxs-lookup"><span data-stu-id="5146a-107">The following figure shows Outlook add-ins activated in the add-in bar for the message in the Reading Pane.</span></span>

![メール読み取りアプリがアクティブ化されたことを示すアプリ バー](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a><span data-ttu-id="5146a-109">マニフェストでのアクティブ化ルールの指定</span><span class="sxs-lookup"><span data-stu-id="5146a-109">Specify activation rules in a manifest</span></span>


<span data-ttu-id="5146a-110">Outlook で特定の条件に応じてアドインをアクティブ化するには、次のいずれかの要素を使用して、アドインマニフェストでアクティブ化ルールを指定し `Rule` ます。</span><span class="sxs-lookup"><span data-stu-id="5146a-110">To have Outlook activate an add-in for specific conditions, specify activation rules in the add-in manifest by using one of the following `Rule` elements:</span></span>

- <span data-ttu-id="5146a-111">[Rule 要素 (MailApp complexType)](../reference/manifest/rule.md) - 個別のルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-111">[Rule element (MailApp complexType)](../reference/manifest/rule.md) - Specifies an individual rule.</span></span>
- <span data-ttu-id="5146a-112">[Rule 要素 (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - 論理演算子を使用して複数のルールを結合します。</span><span class="sxs-lookup"><span data-stu-id="5146a-112">[Rule element (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - Combines multiple rules using logical operations.</span></span>
    

 > [!NOTE]
 > <span data-ttu-id="5146a-113">`Rule`個別のルールを指定するために使用する要素は、抽象[ルール](../reference/manifest/rule.md)複合型です。</span><span class="sxs-lookup"><span data-stu-id="5146a-113">The `Rule` element that you use to specify an individual rule is of the abstract [Rule](../reference/manifest/rule.md) complex type.</span></span> <span data-ttu-id="5146a-114">次のルールの各型は、この抽象 `Rule` 複合型を拡張します。</span><span class="sxs-lookup"><span data-stu-id="5146a-114">Each of the following types of rules extends this abstract `Rule` complex type.</span></span> <span data-ttu-id="5146a-115">したがって、マニフェストで個別のルールを指定するときは、[xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性を使用してルールの以下の型の 1 つをさらに定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5146a-115">So when you specify an individual rule in a manifest, you must use the [xsi:type](https://www.w3.org/TR/xmlschema-1/) attribute to further define one of the following types of rules.</span></span>
 > 
 > <span data-ttu-id="5146a-116">たとえば、次のルールは [ItemIs](../reference/manifest/rule.md#itemis-rule) ルールを定義します。`<Rule xsi:type="ItemIs" ItemType="Message" />`</span><span class="sxs-lookup"><span data-stu-id="5146a-116">For example, the following rule defines an [ItemIs](../reference/manifest/rule.md#itemis-rule) rule: `<Rule xsi:type="ItemIs" ItemType="Message" />`</span></span>
 > 
 > <span data-ttu-id="5146a-117">この属性は、マニフェスト v1.1 の `FormType` アクティブ化ルールに適用されますが、v2.0 では定義されていません `VersionOverrides` 。</span><span class="sxs-lookup"><span data-stu-id="5146a-117">The `FormType` attribute applies to activation rules in the manifest v1.1 but is not defined in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="5146a-118">そのため、ノードで [Itemis](../reference/manifest/rule.md#itemis-rule) が使用されている場合は使用できません `VersionOverrides` 。</span><span class="sxs-lookup"><span data-stu-id="5146a-118">So it can't be used when [ItemIs](../reference/manifest/rule.md#itemis-rule) is used in the `VersionOverrides` node.</span></span>

<span data-ttu-id="5146a-p104">次の表は、使用できるルールの種類を示しています。詳細については、この表の後の説明と、「[閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)」の該当記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5146a-p104">The following table lists the types of rules that are available. You can find more information following the table and in the specified articles under [Create Outlook add-ins for read forms](read-scenario.md).</span></span>

<br/>

|<span data-ttu-id="5146a-121">**ルール名**</span><span class="sxs-lookup"><span data-stu-id="5146a-121">**Rule name**</span></span>|<span data-ttu-id="5146a-122">**該当するフォーム**</span><span class="sxs-lookup"><span data-stu-id="5146a-122">**Applicable forms**</span></span>|<span data-ttu-id="5146a-123">**説明**</span><span class="sxs-lookup"><span data-stu-id="5146a-123">**Description**</span></span>|
|:-----|:-----|:-----|
|[<span data-ttu-id="5146a-124">ItemIs</span><span class="sxs-lookup"><span data-stu-id="5146a-124">ItemIs</span></span>](#itemis-rule)|<span data-ttu-id="5146a-125">読み取り、作成</span><span class="sxs-lookup"><span data-stu-id="5146a-125">Read, Compose</span></span>|<span data-ttu-id="5146a-p105">現在選択されているアイテムは指定された種類のアイテム (メッセージまたは予定) かどうかを調べます。また、アイテム クラス、フォームの種類、さらにはオプションでアイテム メッセージ クラスも調べることができます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p105">Checks to see whether the current item is of the specified type (message or appointment). Can also check the item class and form type.and optionally, item message class.</span></span>|
|[<span data-ttu-id="5146a-128">ItemHasAttachment</span><span class="sxs-lookup"><span data-stu-id="5146a-128">ItemHasAttachment</span></span>](#itemhasattachment-rule)|<span data-ttu-id="5146a-129">読み取り</span><span class="sxs-lookup"><span data-stu-id="5146a-129">Read</span></span>|<span data-ttu-id="5146a-130">選択されているアイテムに添付ファイルが含まれるかどうかを調べます。</span><span class="sxs-lookup"><span data-stu-id="5146a-130">Checks to see whether the selected item contains an attachment.</span></span>|
|[<span data-ttu-id="5146a-131">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="5146a-131">ItemHasKnownEntity</span></span>](#itemhasknownentity-rule)|<span data-ttu-id="5146a-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="5146a-132">Read</span></span>|<span data-ttu-id="5146a-p106">選択されているアイテムに 1 つ以上の一般的なエンティティが含まれるかどうかを調べます。詳細: 「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」。</span><span class="sxs-lookup"><span data-stu-id="5146a-p106">Checks to see whether the selected item contains one or more well-known entities. More information: [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>|
|[<span data-ttu-id="5146a-135">ItemHasRegularExpressionMatch</span><span class="sxs-lookup"><span data-stu-id="5146a-135">ItemHasRegularExpressionMatch</span></span>](#itemhasregularexpressionmatch-rule)|<span data-ttu-id="5146a-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="5146a-136">Read</span></span>|<span data-ttu-id="5146a-137">選択されているアイテムの送信者のメール アドレス、件名、本文に正規表現と一致するものが含まれるかどうかを調べます。詳細: [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="5146a-137">Checks to see whether the sender's email address, the subject, and/or the body of the selected item contains a match to a regular expression.More information: [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>|
|[<span data-ttu-id="5146a-138">RuleCollection</span><span class="sxs-lookup"><span data-stu-id="5146a-138">RuleCollection</span></span>](#rulecollection-rule)|<span data-ttu-id="5146a-139">読み取り、作成</span><span class="sxs-lookup"><span data-stu-id="5146a-139">Read, Compose</span></span>|<span data-ttu-id="5146a-140">複数のルールを組み合わせて、より複雑なルールを作成できます。</span><span class="sxs-lookup"><span data-stu-id="5146a-140">Combines a set of rules so that you can form more complex rules.</span></span>|

## <a name="itemis-rule"></a><span data-ttu-id="5146a-141">ItemIs ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-141">ItemIs rule</span></span>

<span data-ttu-id="5146a-142">**ItemIs** 複合型は、現在のアイテムがアイテムの種類と一致している場合 (また、オプションとしてルールに明記されている場合はアイテムのメッセージ クラスとも一致している場合) に **true** と評価されるルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="5146a-142">The **ItemIs** complex type defines a rule that evaluates to **true** if the current item matches the item type, and optionally the item message class if it's stated in the rule.</span></span>

<span data-ttu-id="5146a-143">`ItemType` **Itemis**ルールの属性に、次のいずれかのアイテムの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-143">Specify one of the following item types in the `ItemType` attribute of an **ItemIs** rule.</span></span> <span data-ttu-id="5146a-144">マニフェストでは、複数の **ItemIs** ルールを指定できます。</span><span class="sxs-lookup"><span data-stu-id="5146a-144">You can specify more than one **ItemIs** rule in a manifest.</span></span> <span data-ttu-id="5146a-145">ItemType simpleType では、Outlook アドインをサポートしている Outlook アイテムの種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="5146a-145">The ItemType simpleType defines the types of Outlook items that support Outlook add-ins.</span></span>

<br/>

|<span data-ttu-id="5146a-146">**値**</span><span class="sxs-lookup"><span data-stu-id="5146a-146">**Value**</span></span>|<span data-ttu-id="5146a-147">**説明**</span><span class="sxs-lookup"><span data-stu-id="5146a-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5146a-148">**Appointment**</span><span class="sxs-lookup"><span data-stu-id="5146a-148">**Appointment**</span></span>|<span data-ttu-id="5146a-p108">Outlook の予定表内のアイテムを指定します。このアイテムには、開催者と出席者を持つ応答済みの会議アイテムと、開催者と出席者を持たない、単なる予定表上のアイテムである予定が含まれます。これは Outlook の IPM.Appointment メッセージ クラスに対応します。</span><span class="sxs-lookup"><span data-stu-id="5146a-p108">Specifies an item in an Outlook calendar. This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that does not have an organizer or attendee and is simply an item on the calendar.This corresponds to the IPM.Appointment message class in Outlook.</span></span>|
|<span data-ttu-id="5146a-151">**Message**</span><span class="sxs-lookup"><span data-stu-id="5146a-151">**Message**</span></span>|<span data-ttu-id="5146a-152">通常は受信トレイで受信される次のアイテムのいずれかを指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-152">Specifies one of the following items received in typically the Inbox:</span></span> <ul><li><p><span data-ttu-id="5146a-p109">電子メール メッセージ。これは Outlook の IPM.Note メッセージ クラスに対応します。</span><span class="sxs-lookup"><span data-stu-id="5146a-p109">An email message. This corresponds to the IPM.Note message class in Outlook.</span></span></p></li><li><p><span data-ttu-id="5146a-p110">会議出席依頼、返信、または取り消し。Outlook の次のメッセージ クラスに対応します。</span><span class="sxs-lookup"><span data-stu-id="5146a-p110">A meeting request, response, or cancellation. This corresponds to the following  message classes in Outlook:</span></span></p><p><span data-ttu-id="5146a-157">IPM.Schedule.Meeting.Request</span><span class="sxs-lookup"><span data-stu-id="5146a-157">IPM.Schedule.Meeting.Request</span></span></p><p><span data-ttu-id="5146a-158">IPM.Schedule.Meeting.Neg</span><span class="sxs-lookup"><span data-stu-id="5146a-158">IPM.Schedule.Meeting.Neg</span></span></p><p><span data-ttu-id="5146a-159">IPM.Schedule.Meeting.Pos</span><span class="sxs-lookup"><span data-stu-id="5146a-159">IPM.Schedule.Meeting.Pos</span></span></p><p><span data-ttu-id="5146a-160">IPM.Schedule.Meeting.Tent</span><span class="sxs-lookup"><span data-stu-id="5146a-160">IPM.Schedule.Meeting.Tent</span></span></p><p><span data-ttu-id="5146a-161">IPM.Schedule.Meeting.Canceled</span><span class="sxs-lookup"><span data-stu-id="5146a-161">IPM.Schedule.Meeting.Canceled</span></span></p></li></ul>|

<span data-ttu-id="5146a-162">この属性を使用して、 `FormType` アドインをアクティブ化するモード (読み取りまたは新規作成) を指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-162">The `FormType` attribute is used to specify the mode (read or compose) in which the add-in should activate.</span></span>


 > [!NOTE]
 > <span data-ttu-id="5146a-163">ItemIs 属性は、スキーマ v1.1 以降では定義されていますが、v1.0 では定義されてい `FormType` ません `VersionOverrides` 。</span><span class="sxs-lookup"><span data-stu-id="5146a-163">The ItemIs `FormType` attribute is defined in schema v1.1 and later but not in `VersionOverrides` v1.0.</span></span> <span data-ttu-id="5146a-164">`FormType`アドインコマンドを定義するときに、属性を含めないでください。</span><span class="sxs-lookup"><span data-stu-id="5146a-164">Do not include the `FormType` attribute when defining add-in commands.</span></span>

<span data-ttu-id="5146a-165">アドインがアクティブ化された後は、 [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) プロパティを使用して Outlook で現在選択されているアイテムを取得し、 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティを使用して現在のアイテムの種類を取得できます。</span><span class="sxs-lookup"><span data-stu-id="5146a-165">After an add-in is activated, you can use the [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) property to obtain the currently selected item in Outlook, and the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to obtain the type of the current item.</span></span>

<span data-ttu-id="5146a-166">必要に応じて、属性を使用して `ItemClass` アイテムのメッセージクラスを指定し、その `IncludeSubClasses` アイテムが指定したクラスのサブクラスである場合にルールを **true** にする必要があるかどうかを指定する属性を指定できます。</span><span class="sxs-lookup"><span data-stu-id="5146a-166">You can optionally use the `ItemClass` attribute to specify the message class of the item, and the `IncludeSubClasses` attribute to specify whether the rule should be **true** when the item is a subclass of the specified class.</span></span>

<span data-ttu-id="5146a-167">メッセージ クラスの詳細については、「[Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5146a-167">For more information about message classes, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes).</span></span>

<span data-ttu-id="5146a-168">次の例は、ユーザーがメッセージを読むときに Outlook のアドイン バーにアドインを表示する **ItemIs** ルールを示しています。</span><span class="sxs-lookup"><span data-stu-id="5146a-168">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message:</span></span>

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

<span data-ttu-id="5146a-169">次の例は、ユーザーがメッセージまたは予約を閲覧するときに Outlook のアドイン バーにアドインを表示する **ItemIs** ルールを示しています。</span><span class="sxs-lookup"><span data-stu-id="5146a-169">The following example is an **ItemIs** rule that lets users see the add-in in the Outlook add-in bar when the user is reading a message or appointment.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a><span data-ttu-id="5146a-170">ItemHasAttachment ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-170">ItemHasAttachment rule</span></span>


<span data-ttu-id="5146a-171">`ItemHasAttachment`複合型は、選択されているアイテムに添付ファイルが含まれているかどうかを確認するルールを定義します。</span><span class="sxs-lookup"><span data-stu-id="5146a-171">The `ItemHasAttachment` complex type defines a rule that checks if the selected item contains an attachment.</span></span>

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a><span data-ttu-id="5146a-172">ItemHasKnownEntity ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-172">ItemHasKnownEntity rule</span></span>

<span data-ttu-id="5146a-p112">アイテムがアドインで使用可能になる前に、サーバーはそれを調べて、件名と本文に既知のエンティティのいずれかであると考えられるテキストが含まれているかどうかを判断します。これらのエンティティのいずれかが見つかった場合は、その `getEntities` アイテムのまたはメソッドを使用してアクセスする既知のエンティティのコレクションに配置され `getEntitiesByType` ます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p112">Before an item is made available to an add-in, the server examines it to determine whether the subject and body contain any text that is likely to be one of the known entities. If any of these entities are found, it is placed in a collection of known entities that you access by using the `getEntities` or `getEntitiesByType` method of that item.</span></span>

<span data-ttu-id="5146a-p113">指定した `ItemHasKnownEntity` 型のエンティティがアイテム内に存在する場合にアドインを表示するルールを指定できます。次の既知のエンティティを `EntityType` ルールの属性に指定でき `ItemHasKnownEntity` ます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p113">You can specify a rule by using `ItemHasKnownEntity` that shows your add-in when an entity of the specified type is present in the item. You can specify the following known entities in the `EntityType` attribute of an `ItemHasKnownEntity` rule:</span></span>

- <span data-ttu-id="5146a-177">Address</span><span class="sxs-lookup"><span data-stu-id="5146a-177">Address</span></span>
- <span data-ttu-id="5146a-178">Contact</span><span class="sxs-lookup"><span data-stu-id="5146a-178">Contact</span></span>
- <span data-ttu-id="5146a-179">EmailAddress</span><span class="sxs-lookup"><span data-stu-id="5146a-179">EmailAddress</span></span>
- <span data-ttu-id="5146a-180">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="5146a-180">MeetingSuggestion</span></span>
- <span data-ttu-id="5146a-181">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="5146a-181">PhoneNumber</span></span>
- <span data-ttu-id="5146a-182">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="5146a-182">TaskSuggestion</span></span>
- <span data-ttu-id="5146a-183">URL</span><span class="sxs-lookup"><span data-stu-id="5146a-183">URL</span></span>
    
<span data-ttu-id="5146a-p114">必要に応じて、属性に正規表現を含めることができ `RegularExpression` ます。これにより、現在の正規表現に一致するエンティティがある場合にアドインが表示されるようになります。ルールで指定された正規表現に一致するものを取得するに `ItemHasKnownEntity` `getRegExMatches` は、現在選択されている Outlook アイテムに対して or メソッドを使用でき `getFilteredEntitiesByName` ます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p114">You can optionally include a regular expression in the `RegularExpression` attribute so that your add-in is only shown when an entity that matches the regular expression in present. To obtain matches to regular expressions specified in `ItemHasKnownEntity` rules, you can use the `getRegExMatches` or `getFilteredEntitiesByName` method for the currently selected Outlook item.</span></span>

<span data-ttu-id="5146a-186">次の例は、 `Rule` 指定された既知のエンティティのいずれかがメッセージに存在するときにアドインを表示する要素のコレクションを示しています。</span><span class="sxs-lookup"><span data-stu-id="5146a-186">The following example shows a collection of `Rule` elements that show the add-in when one of the specified well-known entities is present in the message.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

<span data-ttu-id="5146a-187">次の例は、 `ItemHasKnownEntity` `RegularExpression` "contoso" という単語を含む URL がメッセージ内に存在するときにアドインをアクティブ化する属性を持つルールを示しています。</span><span class="sxs-lookup"><span data-stu-id="5146a-187">The following example shows an `ItemHasKnownEntity` rule with a `RegularExpression` attribute that activates the add-in when a URL that contains the word "contoso" is present in a message.</span></span>


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

<span data-ttu-id="5146a-188">アクティブ化ルールのエンティティの詳細については、「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5146a-188">For more information about entities in activation rules, see [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>


## <a name="itemhasregularexpressionmatch-rule"></a><span data-ttu-id="5146a-189">ItemHasRegularExpressionMatch ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-189">ItemHasRegularExpressionMatch rule</span></span>

<span data-ttu-id="5146a-p115">`ItemHasRegularExpressionMatch`複合型は、アイテムの指定されたプロパティの内容と一致するように正規表現を使用するルールを定義します。正規表現と一致するテキストがアイテムの指定されたプロパティに含まれている場合、Outlook はアドインバーをアクティブにして、アドインを表示します。`getRegExMatches`現在選択されているアイテムを表すオブジェクトのまたはメソッドを使用して、指定した `getRegExMatchesByName` 正規表現に一致するものを取得できます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p115">The `ItemHasRegularExpressionMatch` complex type defines a rule that uses a regular expression to match the contents of the specified property of an item. If text that matches the regular expression is found in the specified property of the item, Outlook activates the add-in bar and displays the add-in. You can use the `getRegExMatches` or `getRegExMatchesByName` method of the object that represents the currently selected item to obtain matches for the specified regular expression.</span></span>

<span data-ttu-id="5146a-193">次の例は、 `ItemHasRegularExpressionMatch` 選択したアイテムの本文に "apple"、"banana"、または "coconut" が含まれている場合にアドインをアクティブにする方法を示しています。大文字小文字は無視されます。</span><span class="sxs-lookup"><span data-stu-id="5146a-193">The following example shows an `ItemHasRegularExpressionMatch` that activates the add-in when the body of the selected item contains "apple", "banana", or "coconut", ignoring case.</span></span>

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" pPropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

<span data-ttu-id="5146a-194">ルールの使用の詳細について `ItemHasRegularExpressionMatch` は、「 [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5146a-194">For more information about using the `ItemHasRegularExpressionMatch` rule, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>


## <a name="rulecollection-rule"></a><span data-ttu-id="5146a-195">RuleCollection ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-195">RuleCollection rule</span></span>


<span data-ttu-id="5146a-p116">`RuleCollection`複合型は、複数のルールを1つのルールに結合します。コレクション内のルールを論理 OR または論理と組み合わせて使用するかどうかを指定でき `Mode` ます。属性を使用します。</span><span class="sxs-lookup"><span data-stu-id="5146a-p116">The `RuleCollection` complex type combines multiple rules into a single rule. You can specify whether the rules in the collection should be combined with a logical OR or a logical AND by using the `Mode` attribute.</span></span>

<span data-ttu-id="5146a-p117">論理 AND を指定する場合、アドインは、コレクション内で指定されているすべてのルールにアイテムが一致する場合にのみ表示されます。論理 OR を指定する場合は、コレクションで指定されているルールのいずれか 1 つにでもアイテムが一致すれば、アドインは表示されます。</span><span class="sxs-lookup"><span data-stu-id="5146a-p117">When a logical AND is specified, an item must match all the specified rules in the collection to show the add-in. When a logical OR is specified, an item that matches any of the specified rules in the collection will show the add-in.</span></span>

<span data-ttu-id="5146a-p118">ルールを結合して `RuleCollection` 複雑なルールを形成できます。次の例では、ユーザーが予定またはメッセージアイテムを表示していて、アイテムの件名または本文に住所が含まれている場合にアドインをアクティブにします。</span><span class="sxs-lookup"><span data-stu-id="5146a-p118">You can combine `RuleCollection` rules to form complex rules. The following example activates the add-in when the user is viewing an appointment or message item and the subject or body of the item contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<span data-ttu-id="5146a-202">次の例では、ユーザーがメッセージを新規作成するときか、件名か本文に住所が含まれる予定を表示するときに、アドインがアクティブ化されます。</span><span class="sxs-lookup"><span data-stu-id="5146a-202">The following example activates the add-in when the user is composing a message, or when the user is viewing an appointment and the subject or body of the appointment contains an address.</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a><span data-ttu-id="5146a-203">ルールと正規表現の制約事項</span><span class="sxs-lookup"><span data-stu-id="5146a-203">Limits for rules and regular expressions</span></span>


<span data-ttu-id="5146a-p119">Outlook アドインでの満足感を得るには、アクティブ化と API の使用に関するガイドラインに従う必要があります。次の表に、正規表現とルールの一般的な制限を示しますが、アプリケーションごとに固有のルールがあります。詳細については、「 [outlook アドインのアクティブ化と JAVASCRIPT API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 」および「 [outlook アドインのアクティブ化のトラブルシューティング](troubleshoot-outlook-add-in-activation.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5146a-p119">To provide a satisfactory experience with Outlook add-ins, you should adhere to the activation and API usage guidelines. The following table shows general limits for regular expressions and rules but there are specific rules for different applications. For more information, see [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) and [Troubleshoot Outlook add-in activation](troubleshoot-outlook-add-in-activation.md).</span></span>

<br/>

|<span data-ttu-id="5146a-207">**アドインの要素**</span><span class="sxs-lookup"><span data-stu-id="5146a-207">**Add-in element**</span></span>|<span data-ttu-id="5146a-208">**ガイドライン**</span><span class="sxs-lookup"><span data-stu-id="5146a-208">**Guidelines**</span></span>|
|:-----|:-----|
|<span data-ttu-id="5146a-209">マニフェストのサイズ</span><span class="sxs-lookup"><span data-stu-id="5146a-209">Manifest Size</span></span>|<span data-ttu-id="5146a-210">256 KB 未満。</span><span class="sxs-lookup"><span data-stu-id="5146a-210">No larger than 256 KB.</span></span>|
|<span data-ttu-id="5146a-211">ルール</span><span class="sxs-lookup"><span data-stu-id="5146a-211">Rules</span></span>|<span data-ttu-id="5146a-212">15 ルール未満。</span><span class="sxs-lookup"><span data-stu-id="5146a-212">No more than 15 rules.</span></span>|
|<span data-ttu-id="5146a-213">ItemHasKnownEntity</span><span class="sxs-lookup"><span data-stu-id="5146a-213">ItemHasKnownEntity</span></span>|<span data-ttu-id="5146a-214">Outlook リッチ クライアントでは、本文の最初の 1 MB にルールを適用し、残りの部分には適用しません。</span><span class="sxs-lookup"><span data-stu-id="5146a-214">An Outlook rich client will apply the rule against the first 1 MB of the body, and not to the rest of the body.</span></span>|
|<span data-ttu-id="5146a-215">正規表現</span><span class="sxs-lookup"><span data-stu-id="5146a-215">Regular Expressions</span></span>|<span data-ttu-id="5146a-216">すべての Outlook アプリケーションの ItemHasKnownEntity または ItemHasRegularExpressionMatch ルールの場合:</span><span class="sxs-lookup"><span data-stu-id="5146a-216">For ItemHasKnownEntity or ItemHasRegularExpressionMatch rules for all Outlook applications:</span></span><br><ul><li><span data-ttu-id="5146a-p120">Outlook アドインのアクティブ化ルールで指定する正規表現は 5 個までにしてください。その制約数を超えるアドインをインストールすることはできません。</span><span class="sxs-lookup"><span data-stu-id="5146a-p120">Specify no more than 5 regular expressions in activation rules for an Outlook add-in. You cannot install an add-in if you exceed that limit.</span></span></li><li><span data-ttu-id="5146a-219">予期される結果が <b>getRegExMatches</b> メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-219">Specify regular expressions whose anticipated results are returned by the <b>getRegExMatches</b> method call within the first 50 matches.</span></span> </li><li><span data-ttu-id="5146a-220">正規表現で先読みアサーションは指定しますが、後読み `(?<=text)` および否定の後読み `(?<!text)` アサーションは指定しません。</span><span class="sxs-lookup"><span data-stu-id="5146a-220">Specify look-ahead assertions in regular expressions, but not look-behind, `(?<=text)`, and negative look-behind `(?<!text)`.</span></span></li><li><span data-ttu-id="5146a-221">一致数が次の表の制限を超えない正規表現を指定します。</span><span class="sxs-lookup"><span data-stu-id="5146a-221">Specify regular expressions whose match does not exceed the limits in the table below.</span></span><br/><br/><table><tr><th><span data-ttu-id="5146a-222">正規表現の長さ制限</span><span class="sxs-lookup"><span data-stu-id="5146a-222">Limit on length of a regex match</span></span></th><th><span data-ttu-id="5146a-223">Outlook リッチ クライアント</span><span class="sxs-lookup"><span data-stu-id="5146a-223">Outlook rich clients</span></span></th><th><span data-ttu-id="5146a-224">iOS および Android 用の Outlook</span><span class="sxs-lookup"><span data-stu-id="5146a-224">Outlook on iOS and Android</span></span></th></tr><tr><td><span data-ttu-id="5146a-225">アイテムの本文がテキスト形式の場合</span><span class="sxs-lookup"><span data-stu-id="5146a-225">Item body is plain text</span></span></td><td><span data-ttu-id="5146a-226">1.5 KB</span><span class="sxs-lookup"><span data-stu-id="5146a-226">1.5 KB</span></span></td><td><span data-ttu-id="5146a-227">3 KB</span><span class="sxs-lookup"><span data-stu-id="5146a-227">3 KB</span></span></td></tr><tr><td><span data-ttu-id="5146a-228">アイテムの本文が HTML の場合</span><span class="sxs-lookup"><span data-stu-id="5146a-228">Item body it HTML</span></span></td><td><span data-ttu-id="5146a-229">3 KB</span><span class="sxs-lookup"><span data-stu-id="5146a-229">3 KB</span></span></td><td><span data-ttu-id="5146a-230">3 KB</span><span class="sxs-lookup"><span data-stu-id="5146a-230">3 KB</span></span></td></tr></table>|

## <a name="see-also"></a><span data-ttu-id="5146a-231">関連項目</span><span class="sxs-lookup"><span data-stu-id="5146a-231">See also</span></span>

- [<span data-ttu-id="5146a-232">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="5146a-232">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="5146a-233">Outlook アドインのアクティブ化と JavaScript API の制限</span><span class="sxs-lookup"><span data-stu-id="5146a-233">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="5146a-234">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="5146a-234">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="5146a-235">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="5146a-235">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)
    
