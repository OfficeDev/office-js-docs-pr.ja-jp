---
title: Outlook アドインで既知のエンティティとして文字列を照合する
description: JavaScript API Office使用すると、特定の既知のエンティティに一致する文字列を取得して、さらに処理できます。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 8d4b78259b771d29244641d9e3ca867018b763ef
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348497"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a><span data-ttu-id="465f1-103">Outlook アイテム内の文字列を既知のエンティティとして照合する</span><span class="sxs-lookup"><span data-stu-id="465f1-103">Match strings in an Outlook item as well-known entities</span></span>

<span data-ttu-id="465f1-p101">メッセージおよび会議出席依頼のアイテムを送信する前に、Exchange Server によりアイテムの内容が解析され、件名と本文から、メール アドレス、電話番号、URL など、Exchange にとっての既知のエンティティに似た文字列が特定され、スタンプが付けられます。メッセージと会議出席依頼は、Exchange Server によって、既知のエンティティにスタンプが付けられた状態で、Outlook の受信トレイに配信されます。</span><span class="sxs-lookup"><span data-stu-id="465f1-p101">Before sending a message or meeting request item, Exchange Server parses the contents of the item, identifies and stamps certain strings in the subject and body that resemble entities well-known to Exchange, for example, email addresses, phone numbers, and URLs. Messages and meeting requests are delivered by Exchange Server in an Outlook Inbox with well-known entities stamped.</span></span> 

<span data-ttu-id="465f1-106">JavaScript API Office使用すると、特定の既知のエンティティと一致するこれらの文字列を取得して、さらに処理することができます。</span><span class="sxs-lookup"><span data-stu-id="465f1-106">Using the Office JavaScript API, you can get these strings that match specific well-known entities for further processing.</span></span> <span data-ttu-id="465f1-107">さらに、既知のエンティティをアドイン マニフェスト内のルールで指定して、ユーザーがそのエンティティと一致するものを含んだアイテムを表示したときに、Outlook がアドインをアクティブにするように設定することもできます。</span><span class="sxs-lookup"><span data-stu-id="465f1-107">You can also specify a well-known entity in a rule in the add-in manifest so that Outlook can activate your add-in when the user is viewing an item that contains matches for that entity.</span></span> <span data-ttu-id="465f1-108">その後で、エンティティと一致するものを抽出してアクションを実行することができます。</span><span class="sxs-lookup"><span data-stu-id="465f1-108">You can then extract and take action on matches for the entity.</span></span> 

<span data-ttu-id="465f1-109">選択されたメッセージや予定からこれらのインスタンスを特定したり抽出したりできるので便利です。</span><span class="sxs-lookup"><span data-stu-id="465f1-109">Being able to identify or extract such instances from a selected message or appointment is convenient.</span></span> <span data-ttu-id="465f1-110">たとえば、Outlook のアドインとして電話番号の逆引き検索サービスを作成できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-110">For example, you can build a reverse phone look-up service as an Outlook add-in.</span></span> <span data-ttu-id="465f1-111">このアドインは、アイテムの件名や本文から電話番号に似た文字列を抽出して逆引き検索を行い、各電話番号の登録所有者を表示させることができます。</span><span class="sxs-lookup"><span data-stu-id="465f1-111">The add-in can extract strings in the item subject or body that resemble a phone number, do a reverse lookup, and display the registered owner of each phone number.</span></span>

<span data-ttu-id="465f1-112">このトピックでは既知のエンティティ、既知のエンティティに基づくアクティブ化ルールの例、およびアクティブ化ルール内でエンティティが使用されているかどうかに関係なく、一致するエンティティを抽出する方法を紹介します。</span><span class="sxs-lookup"><span data-stu-id="465f1-112">This topic introduces these well-known entities, shows examples of activation rules based on well-known entities, and how to extract entity matches independently of having used entities in activation rules.</span></span>


## <a name="support-for-well-known-entities"></a><span data-ttu-id="465f1-113">既知のエンティティに対するサポート</span><span class="sxs-lookup"><span data-stu-id="465f1-113">Support for well-known entities</span></span>

<span data-ttu-id="465f1-p104">Exchange Server は、ユーザーがメッセージや会議出席依頼アイテムを送信した後、それが受信者に配信される前に、アイテム内の既知のエンティティにスタンプを付けます。そのため、Exchange 内のトランスポートを通過したアイテムだけにスタンプが付けられ、Outlook はユーザーがそのようなアイテムを表示中にそれらのスタンプに基づいてアドインをアクティブにすることができます。しかし、ユーザーがアイテムを作成している間や、送信済みアイテム フォルダー内のアイテムを表示しているときは、そのアイテムがまだトランスポートを通過していないため、Outlook は既知のエンティティに基づいてアドインをアクティブにすることができません。</span><span class="sxs-lookup"><span data-stu-id="465f1-p104">Exchange Server stamps well-known entities in a message or meeting request item after the sender sends the item and before Exchange delivers the item to the recipient. Therefore, only items that have gone through transport in Exchange are stamped, and Outlook can activate add-ins based on these stamps when the user is viewing such items. On the contrary, when the user is composing an item or viewing an item that is in the Sent Items folder, because the item has not gone through transport, Outlook cannot activate add-ins based on well-known entities.</span></span> 

<span data-ttu-id="465f1-p105">同様に、作成中または送信済みアイテム フォルダー内のアイテムはトランスポートを通過しておらず、スタンプが付けられていないため、既知のエンティティを抽出できません。アクティブ化をサポートしているアイテムの種類の詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="465f1-p105">Similarly, you cannot extract well-known entities in items that are being composed or in the Sent Items folder, as these items have not gone through transport and are not stamped. For additional information about the kinds of items that support activation, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

<span data-ttu-id="465f1-p106">次の表は、Exchange Server と Outlook でサポートされ、認識されるエンティティ (つまり、「既知のエンティティ」) と、各エンティティのインスタンスのオブジェクト タイプを一覧にしたものです。これらのエンティティの 1 つとしての文字列の自然言語認識は、大量のデータに対してトレーニングを行った学習モデルに基づきます。したがって、認識は決定論的ではありません。認識に関する条件の詳細については、「 [既知のエンティティを使用するためのヒント](#tips-for-using-well-known-entities)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="465f1-p106">The following table lists the entities that Exchange Server and Outlook support and recognize (hence the name "well-known entities"), and the object type of an instance of each entity. The natural language recognition of a string as one of these entities is based on a learning model that has been trained on a large amount of data. Therefore, the recognition is non-deterministic. See [Tips for using well-known entities](#tips-for-using-well-known-entities) for more information about conditions for recognition.</span></span>

<span data-ttu-id="465f1-123">**表 1.サポートされるエンティティとその型**</span><span class="sxs-lookup"><span data-stu-id="465f1-123">**Table 1. Supported entities and their types**</span></span>

|<span data-ttu-id="465f1-124">エンティティの型</span><span class="sxs-lookup"><span data-stu-id="465f1-124">Entity type</span></span>|<span data-ttu-id="465f1-125">認識に関する条件</span><span class="sxs-lookup"><span data-stu-id="465f1-125">Conditions for recognition</span></span>|<span data-ttu-id="465f1-126">オブジェクトの種類</span><span class="sxs-lookup"><span data-stu-id="465f1-126">Object type</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="465f1-127">**住所**</span><span class="sxs-lookup"><span data-stu-id="465f1-127">**Address**</span></span>|<span data-ttu-id="465f1-p107">米国の住所。次はその例です。1234 Main Street, Redmond, WA 07722.通常、住所が認識されるには、米国の住所の構造に従う必要があり、ほとんどには番地、住所、都市名、州名、郵便番号の要素が存在します。住所は 1 行または複数行で指定できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-p107">United States street addresses; for example: 1234 Main Street, Redmond, WA 07722. Generally, for an address to be recognized, it should follow the structure of a United States postal address, with most of the elements of a street number, street name, city, state, and zip code present. The address can be specified in one or multiple lines.</span></span>|<span data-ttu-id="465f1-131">JavaScript **String** オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-131">JavaScript **String** object</span></span>|
|<span data-ttu-id="465f1-132">**連絡先**</span><span class="sxs-lookup"><span data-stu-id="465f1-132">**Contact**</span></span>|<span data-ttu-id="465f1-133">自然言語の認識による、人に関する情報の参照。</span><span class="sxs-lookup"><span data-stu-id="465f1-133">A reference to a person's information as recognized in natural language.</span></span> <span data-ttu-id="465f1-134">連絡先の認識は、状況によりさまざまな方法で行われます。</span><span class="sxs-lookup"><span data-stu-id="465f1-134">The recognition of a contact depends on the context.</span></span> <span data-ttu-id="465f1-135">たとえば、メッセージの最後にある署名や、人の名前の近くに現れる電話番号、住所、メール アドレス、URL などの情報です。</span><span class="sxs-lookup"><span data-stu-id="465f1-135">For example, a signature at the end of a message, or a person's name appearing in the vicinity of some of the following information: a phone number, address, email address, and URL.</span></span>|<span data-ttu-id="465f1-136">[Contact](/javascript/api/outlook/office.contact) オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-136">[Contact](/javascript/api/outlook/office.contact) object</span></span>|
|<span data-ttu-id="465f1-137">**EmailAddress**</span><span class="sxs-lookup"><span data-stu-id="465f1-137">**EmailAddress**</span></span>|<span data-ttu-id="465f1-138">SMTP メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="465f1-138">SMTP email addresses.</span></span>|<span data-ttu-id="465f1-139">JavaScript `String` オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-139">JavaScript `String` object</span></span>|
|<span data-ttu-id="465f1-140">**MeetingSuggestion**</span><span class="sxs-lookup"><span data-stu-id="465f1-140">**MeetingSuggestion**</span></span>|<span data-ttu-id="465f1-p109">イベントまたは会議の参照。たとえば、Exchange 2013では次のテキストは会議の提案として認識されます。 _明日、昼食会議を開きましょう。_</span><span class="sxs-lookup"><span data-stu-id="465f1-p109">A reference to an event or meeting. For example, Exchange 2013 would recognize the following text as a meeting suggestion:  _Let's meet tomorrow for lunch._</span></span>|<span data-ttu-id="465f1-143">[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-143">[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) object</span></span>|
|<span data-ttu-id="465f1-144">**PhoneNumber**</span><span class="sxs-lookup"><span data-stu-id="465f1-144">**PhoneNumber**</span></span>|<span data-ttu-id="465f1-145">米国の電話番号。次はその例です。_(235) 555-0110_</span><span class="sxs-lookup"><span data-stu-id="465f1-145">United States telephone numbers; for example:  _(235) 555-0110_</span></span>|<span data-ttu-id="465f1-146">[PhoneNumber](/javascript/api/outlook/office.phonenumber) オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-146">[PhoneNumber](/javascript/api/outlook/office.phonenumber) object</span></span>|
|<span data-ttu-id="465f1-147">**TaskSuggestion**</span><span class="sxs-lookup"><span data-stu-id="465f1-147">**TaskSuggestion**</span></span>|<span data-ttu-id="465f1-p110">電子メールの対応可能な文言。たとえば、_スプレッドシートを更新してください。_</span><span class="sxs-lookup"><span data-stu-id="465f1-p110">Actionable sentences in an email. For example:  _Please update the spreadsheet._</span></span>|<span data-ttu-id="465f1-150">[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-150">[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) object</span></span>|
|<span data-ttu-id="465f1-151">**Url**</span><span class="sxs-lookup"><span data-stu-id="465f1-151">**Url**</span></span>|<span data-ttu-id="465f1-152">ネットワーク ロケーションと Web リソースの識別子を明記した Web アドレス。</span><span class="sxs-lookup"><span data-stu-id="465f1-152">A web address that explicitly specifies the network location and identifier for a web resource.</span></span> <span data-ttu-id="465f1-153">Exchange Serverは、Web アドレスにアクセス プロトコルを必要としないし、リンク テキストに埋め込まれている URL をエンティティのインスタンスとして認識 `Url` しない。</span><span class="sxs-lookup"><span data-stu-id="465f1-153">Exchange Server does not require the access protocol in the web address, and does not recognize URLs that are embedded in link text as instances of the `Url` entity.</span></span> <span data-ttu-id="465f1-154">Exchange Serverは、次の例と一致できます。 `www.youtube.com/user/officevideos``https://www.youtube.com/user/officevideos`</span><span class="sxs-lookup"><span data-stu-id="465f1-154">Exchange Server can match the following examples: `www.youtube.com/user/officevideos` `https://www.youtube.com/user/officevideos`</span></span> |<span data-ttu-id="465f1-155">JavaScript `String` オブジェクト</span><span class="sxs-lookup"><span data-stu-id="465f1-155">JavaScript `String` object</span></span>|

<br/>

<span data-ttu-id="465f1-p112">次の図は、アドインで Exchange Server と Outlook が既知のエンティティをサポートする仕組みと、既知のエンティティを使用してアドインでできる操作について説明しています。エンティティの利用方法について詳しくは、「[アドインでのエンティティの取得](#retrieving-entities-in-your-add-in)」と「[エンティティの存在に基づくアドインのアクティブ化](#activating-an-add-in-based-on-the-existence-of-an-entity)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="465f1-p112">The following figure describes how Exchange Server and Outlook support well-known entities for add-ins, and what add-ins can do with well-known entities. See [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in) and [Activating an add-in based on the existence of an entity](#activating-an-add-in-based-on-the-existence-of-an-entity) for more details on how to use these entities.</span></span>

<span data-ttu-id="465f1-158">**Exchange Server、Outlook、アドインが既知のエンティティをサポートする仕組み**</span><span class="sxs-lookup"><span data-stu-id="465f1-158">**How Exchange Server, Outlook, and add-ins support well-known entities**</span></span>

![メール アプリでの既知のエンティティのサポートと使用。](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a><span data-ttu-id="465f1-160">エンティティを抽出するためのアクセス許可</span><span class="sxs-lookup"><span data-stu-id="465f1-160">Permissions to extract entities</span></span>

<span data-ttu-id="465f1-161">JavaScript コードでエンティティを抽出したり、特定の既知のエンティティの存在に基づいてアドインをアクティブ化したりする場合は、アドイン マニフェストで適切なアクセス許可を要求しておきます。</span><span class="sxs-lookup"><span data-stu-id="465f1-161">To extract entities in your JavaScript code or to have your add-in activated based on the existence of certain well-known entities, make sure you have requested the appropriate permissions in the add-in manifest.</span></span>

<span data-ttu-id="465f1-162">既定の制限付きアクセス許可を指定すると、アドインで `Address` 、 、または `MeetingSuggestion` エンティティを抽出 `TaskSuggestion` できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-162">Specifying the default restricted permission allows your add-in to extract the `Address`, `MeetingSuggestion`, or `TaskSuggestion` entity.</span></span> <span data-ttu-id="465f1-163">その他のエンティティを抽出するには、開封済みアイテム、読み取り/書き込みアイテム、またはメールボックスの読み取り/書き込み許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="465f1-163">To extract any of the other entities, specify read item, read/write item, or read/write mailbox permission.</span></span> <span data-ttu-id="465f1-164">マニフェストでこれを行うには [、Permissions](../reference/manifest/permissions.md)要素を使用し、次の例のように、適切なアクセス許可 &mdash; Restricted、ReadItem、ReadWriteItem、**または ReadWriteMailbox** &mdash; を指定します。</span><span class="sxs-lookup"><span data-stu-id="465f1-164">To do that in the manifest, use the [Permissions](../reference/manifest/permissions.md) element and specify the appropriate permission&mdash;**Restricted**, **ReadItem**, **ReadWriteItem**, or **ReadWriteMailbox**&mdash;as in the following example.</span></span>

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a><span data-ttu-id="465f1-165">アドインでのエンティティの取得</span><span class="sxs-lookup"><span data-stu-id="465f1-165">Retrieving entities in your add-in</span></span>

<span data-ttu-id="465f1-166">ユーザーが表示するアイテムの件名または本文に、Exchange および Outlook が既知のエンティティとして認識できる文字列が含まれている限り、これらのインスタンスはアドインで使用できます。これらは、既知のエンティティに基づいてアドインがアクティブ化されていない場合でも使用できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-166">As long as the subject or body of the item that is being viewed by the user contains strings that Exchange and Outlook can recognize as well-known entities, these instances are available to add-ins. They are available even if an add-in is not activated based on well-known entities.</span></span> <span data-ttu-id="465f1-167">適切なアクセス許可を使用すると、or メソッドを使用して、現在のメッセージまたは予定に存在する既知のエンティティ `getEntities` `getEntitiesByType` を取得できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-167">With the appropriate permission, you can use the `getEntities` or `getEntitiesByType` method to retrieve well-known entities that are present in the current message or appointment.</span></span>

<span data-ttu-id="465f1-168">この `getEntities` メソッドは、アイテム内のすべての既知のエンティティを含む [Entities](/javascript/api/outlook/office.entities) オブジェクトの配列を返します。</span><span class="sxs-lookup"><span data-stu-id="465f1-168">The `getEntities` method returns an array of [Entities](/javascript/api/outlook/office.entities) objects that contains all the well-known entities in the item.</span></span>

<span data-ttu-id="465f1-169">特定の種類のエンティティに興味がある場合は、必要なエンティティの配列のみを返すメソッド `getEntitiesByType` を使用します。</span><span class="sxs-lookup"><span data-stu-id="465f1-169">If you're interested in a particular type of entities, use the `getEntitiesByType`method which returns an array of only the entities you want.</span></span> <span data-ttu-id="465f1-170">[EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 列挙型は抽出可能なすべての既知のエンティティの種類を表します。</span><span class="sxs-lookup"><span data-stu-id="465f1-170">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) enumeration represents all the types of well-known entities you can extract.</span></span>

<span data-ttu-id="465f1-171">呼び出した後、オブジェクトの対応するプロパティを使用して、エンティティの種類のインスタンスの `getEntities` `Entities` 配列を取得できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-171">After calling `getEntities`, you can then use the corresponding property of the `Entities` object to obtain an array of instances of a type of entity.</span></span> <span data-ttu-id="465f1-172">エンティティの型により、配列内のインスタンスは単なる文字列であることも、特定のオブジェクトにマップできることもあります。</span><span class="sxs-lookup"><span data-stu-id="465f1-172">Depending on the type of entity, the instances in the array can be just strings, or can map to specific objects.</span></span> 

<span data-ttu-id="465f1-173">たとえば、前出の図のように、アイテムのアドレスを取得するには、`getEntities().addresses[]` により返される配列にアクセスします。</span><span class="sxs-lookup"><span data-stu-id="465f1-173">As an example seen in the earlier figure, to get addresses in the item, access the array returned by `getEntities().addresses[]`.</span></span> <span data-ttu-id="465f1-174">この `Entities.addresses` プロパティは、郵便番号として認識Outlook文字列の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="465f1-174">The `Entities.addresses` property returns an array of strings that Outlook recognizes as postal addresses.</span></span> <span data-ttu-id="465f1-175">同様に、 `Entities.contacts` このプロパティは、連絡先情報として認識Outlook `Contact` オブジェクトの配列を返します。</span><span class="sxs-lookup"><span data-stu-id="465f1-175">Similarly, the `Entities.contacts` property returns an array of `Contact` objects that Outlook recognizes as contact information.</span></span> <span data-ttu-id="465f1-176">表 1 に、サポートされる各エンティティのインスタンスのオブジェクト型を示します。</span><span class="sxs-lookup"><span data-stu-id="465f1-176">Tables 1 lists the object type of an instance of each supported entity.</span></span>

<span data-ttu-id="465f1-177">以下の例では、メッセージ内で見つかった住所を取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="465f1-177">The following example shows how to retrieve any addresses found in a message.</span></span>

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a><span data-ttu-id="465f1-178">エンティティの存在に基づくアドインのアクティブ化</span><span class="sxs-lookup"><span data-stu-id="465f1-178">Activating an add-in based on the existence of an entity</span></span>

<span data-ttu-id="465f1-179">既知のエンティティを利用するもう 1 つの方法は、現在表示されているアイテムの件名または本文に 1 つまたは複数の種類のエンティティが存在するかどうかに基づいて Outlook にアドインをアクティブ化させる方法です。</span><span class="sxs-lookup"><span data-stu-id="465f1-179">Another way to use well-known entities is to have Outlook activate your add-in based on the existence of one or more types of entities in the subject or body of the currently viewed item.</span></span> <span data-ttu-id="465f1-180">これを行うには、アドイン マニフェスト `ItemHasKnownEntity` でルールを指定します。</span><span class="sxs-lookup"><span data-stu-id="465f1-180">You can do so by specifying an `ItemHasKnownEntity` rule in the add-in manifest.</span></span> <span data-ttu-id="465f1-181">[EntityType 単純](/javascript/api/outlook/office.mailboxenums.entitytype)型は、ルールでサポートされる既知のエンティティの種類を表 `ItemHasKnownEntity` します。</span><span class="sxs-lookup"><span data-stu-id="465f1-181">The [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) simple type represents the different types of well-known entities supported by `ItemHasKnownEntity` rules.</span></span> <span data-ttu-id="465f1-182">アドインがアクティブ化されたら、前のセクション「 [アドインでのエンティティの取得](#retrieving-entities-in-your-add-in)」で説明したように、目的のエンティティのインスタンスを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="465f1-182">After your add-in is activated, you can also retrieve the instances of such entities for your purposes, as described in the previous section [Retrieving entities in your add-in](#retrieving-entities-in-your-add-in).</span></span>

<span data-ttu-id="465f1-183">必要に応じて、ルールに正規表現を適用して、エンティティのインスタンスをさらにフィルター処理し、Outlook がエンティティのインスタンスのサブセットでのみアドインをアクティブ化できます `ItemHasKnownEntity` 。</span><span class="sxs-lookup"><span data-stu-id="465f1-183">You can optionally apply a regular expression in an `ItemHasKnownEntity` rule, so as to further filter instances of an entity and have Outlook activate an add-in only on a subset of the instances of the entity.</span></span> <span data-ttu-id="465f1-184">たとえば、"98" で始まるワシントン州の郵便番号を含むメッセージの中の街路住所エンティティを検出するフィルターを指定できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-184">For example, you can specify a filter for the street address entity in a message that contains a Washington state zip code beginning with "98".</span></span> <span data-ttu-id="465f1-185">エンティティ インスタンスにフィルターを適用するには `RegExFilter` `FilterName` `Rule` [、ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 型の要素の and 属性を使用します。</span><span class="sxs-lookup"><span data-stu-id="465f1-185">To apply a filter on the entity instances, use the `RegExFilter` and `FilterName` attributes in the `Rule` element of the [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) type.</span></span>

<span data-ttu-id="465f1-186">他のアクティブ化ルールと同様に、複数のルールを指定してアドインのルール コレクションを作成できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-186">Similar to other activation rules, you can specify multiple rules to form a rule collection for your add-in.</span></span> <span data-ttu-id="465f1-187">次の例では、ルールとルールの 2 つのルールに "AND" `ItemIs` 操作を適用 `ItemHasKnownEntity` します。</span><span class="sxs-lookup"><span data-stu-id="465f1-187">The following example applies an "AND" operation on 2 rules: an `ItemIs` rule and an `ItemHasKnownEntity` rule.</span></span> <span data-ttu-id="465f1-188">このルール コレクションにより、現在のアイテムがメッセージである場合に、Outlook がそのアイテムの件名または本文から住所を認識すると、アドインがアクティブ化されます。</span><span class="sxs-lookup"><span data-stu-id="465f1-188">This rule collection activates the add-in whenever the current item is a message and Outlook recognizes an address in the subject or body of that item.</span></span>

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

<span data-ttu-id="465f1-189">次の使用例は `getEntitiesByType` 、現在のアイテムを使用して、前のルール コレクションの結果に変数 `addresses` を設定します。</span><span class="sxs-lookup"><span data-stu-id="465f1-189">The following example uses `getEntitiesByType` of the current item to set a variable `addresses` to the results of the preceding rule collection.</span></span>

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

<span data-ttu-id="465f1-190">次のルール例は、現在のアイテムの件名または本文に URL が含まれている場合は常にアドインをアクティブ化し、URL には文字列の場合に関係なく、"youtube" という文字列が含 `ItemHasKnownEntity` まれます。</span><span class="sxs-lookup"><span data-stu-id="465f1-190">The following `ItemHasKnownEntity` rule example activates the add-in whenever there is a URL in the subject or body of the current item, and the URL contains the string "youtube", regardless of the case of the string.</span></span>

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

<span data-ttu-id="465f1-191">次の使用例は、現在のアイテムを使用して変数を設定し、前のルールの正規表現に一致する結果の配列 `getFilteredEntitiesByName(name)` `videos` を取得 `ItemHasKnownEntity` します。</span><span class="sxs-lookup"><span data-stu-id="465f1-191">The following example uses `getFilteredEntitiesByName(name)` of the current item to set a variable `videos` to get an array of results that match the regular expression in the preceding `ItemHasKnownEntity` rule.</span></span>

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a><span data-ttu-id="465f1-192">既知のエンティティを使用するためのヒント</span><span class="sxs-lookup"><span data-stu-id="465f1-192">Tips for using well-known entities</span></span>

<span data-ttu-id="465f1-193">アドインで既知のエンティティを使用する場合に知っておくべきいくつかの事実と制限があります。</span><span class="sxs-lookup"><span data-stu-id="465f1-193">There are a few facts and limits you should be aware of if you use well-known entities in your add-in.</span></span> <span data-ttu-id="465f1-194">ルールを使用するかどうかに関係なく、ユーザーが既知のエンティティの一致を含むアイテムを読み取っているときにアドインがアクティブ化されている限り、以下が適用 `ItemHasKnownEntity` されます。</span><span class="sxs-lookup"><span data-stu-id="465f1-194">The following applies as long as your add-in is activated when the user is reading an item which contains matches of well-known entities, regardless of whether you use an `ItemHasKnownEntity` rule.</span></span>


- <span data-ttu-id="465f1-195">文字列が英語の場合にのみ、既知のエンティティである文字列を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="465f1-195">You can extract strings that are well-known entities only if the strings are in English.</span></span>
    
- <span data-ttu-id="465f1-196">アイテム本文の最初の 2,000 文字から既知のエンティティを抽出できます。2,000 を超える文字からは抽出できません。</span><span class="sxs-lookup"><span data-stu-id="465f1-196">You can extract well-known entities from the first 2,000 characters in the item body, but not beyond that limit.</span></span> <span data-ttu-id="465f1-197">このサイズ制限により機能とパフォーマンスのニーズのバランスが維持されるため、サイズの大きなメッセージと予定の中から既知のエンティティのインスタンスの解析と特定をしても、Exchange Server と Outlook は停止しません。</span><span class="sxs-lookup"><span data-stu-id="465f1-197">This size limit helps balance the need for functionality and performance, so that Exchange Server and Outlook are not bogged down by parsing and identifying instances of well-known entities in large messages and appointments.</span></span> <span data-ttu-id="465f1-198">この制限は、アドインがルールを指定するかどうかに依存しない点に注意 `ItemHasKnownEntity` してください。</span><span class="sxs-lookup"><span data-stu-id="465f1-198">Note that this limit is independent of whether the add-in specifies an `ItemHasKnownEntity` rule.</span></span> <span data-ttu-id="465f1-199">アドインでそのようなルールを使用する場合には、Outlook リッチ クライアントに対する以下の 2 番目の項目のルール処理制限にも注意してください。</span><span class="sxs-lookup"><span data-stu-id="465f1-199">If the add-in does use such a rule, note also the rule processing limit in item 2 below for the Outlook rich clients.</span></span>
    
- <span data-ttu-id="465f1-p123">メールボックスの所有者以外の誰かが計画した会議である予定からエンティティを抽出できます。会議ではないカレンダー アイテムやメールボックスの所有者が計画した会議である予定からエンティティを抽出することはできません。</span><span class="sxs-lookup"><span data-stu-id="465f1-p123">You can extract entities from appointments that are meetings organized by someone other than the mailbox owner. You cannot extract entities from calendar items that are not meetings, or meetings organized by the mailbox owner.</span></span>
    
- <span data-ttu-id="465f1-202">種類のエンティティはメッセージのみから抽出できますが `MeetingSuggestion` 、予定は抽出しません。</span><span class="sxs-lookup"><span data-stu-id="465f1-202">You can extract entities of the `MeetingSuggestion` type from only messages but not appointments.</span></span>
    
- <span data-ttu-id="465f1-203">アイテム本文に明示的に存在する URL を抽出することはできますが、HTML のアイテム本文のハイパーリンク テキストに埋め込まれている URL を抽出することはできません。</span><span class="sxs-lookup"><span data-stu-id="465f1-203">You can extract URLs that exist explicitly in the item body, but not URLs that are embedded in hyperlinked text in HTML item body.</span></span> <span data-ttu-id="465f1-204">明示的な URL と埋め込み URL の両方を取得するには、代わりにルール `ItemHasRegularExpressionMatch` を使用してください。</span><span class="sxs-lookup"><span data-stu-id="465f1-204">Consider using an `ItemHasRegularExpressionMatch` rule instead to get both explicit and embedded URLs.</span></span> <span data-ttu-id="465f1-205">`BodyAsHTML`PropertyName _として指定し_、URL と一致する正規表現を _RegExValue として指定します_。</span><span class="sxs-lookup"><span data-stu-id="465f1-205">Specify `BodyAsHTML` as the _PropertyName_, and a regular expression that matches URLs as the  _RegExValue_.</span></span>
    
- <span data-ttu-id="465f1-206">[送信済みアイテム] フォルダーのアイテムからエンティティを抽出することはできません。</span><span class="sxs-lookup"><span data-stu-id="465f1-206">You cannot extract entities from items in the Sent Items folder.</span></span>
    
<span data-ttu-id="465f1-207">さらに [、ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用する場合は、次のことが適用され、アドインがアクティブ化されると予想されるシナリオに影響を与える可能性があります。</span><span class="sxs-lookup"><span data-stu-id="465f1-207">In addition, the following applies if you use an [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) rule, and may affect the scenarios where you'd otherwise expect your add-in to be activated.</span></span>

- <span data-ttu-id="465f1-208">ルールを使用する場合、マニフェストOutlook既定のロケールに関係なく、エンティティ文字列が英語でのみ一致する必要 `ItemHasKnownEntity` があります。</span><span class="sxs-lookup"><span data-stu-id="465f1-208">When using the `ItemHasKnownEntity` rule, expect Outlook to match entity strings in only English regardless of the default locale specified in the manifest.</span></span>
    
- <span data-ttu-id="465f1-209">アドインが Outlook リッチ クライアントで実行されている場合は、Outlook がアイテム本文の最初のメガバイトにルールを適用し、その制限を超える残りの本文には適用されないと予想します。 `ItemHasKnownEntity`</span><span class="sxs-lookup"><span data-stu-id="465f1-209">When your add-in is running on an Outlook rich client, expect Outlook to apply the `ItemHasKnownEntity` rule to the first megabyte of the item body and not to the rest of the body over that limit.</span></span>
    
- <span data-ttu-id="465f1-210">ルールを使用して、[送信されたアイテム] フォルダー内のアイテムのアドイン `ItemHasKnownEntity` をアクティブ化することはできません。</span><span class="sxs-lookup"><span data-stu-id="465f1-210">You cannot use an `ItemHasKnownEntity` rule to activate an add-in for items in the Sent Items folder.</span></span>
    

## <a name="see-also"></a><span data-ttu-id="465f1-211">関連項目</span><span class="sxs-lookup"><span data-stu-id="465f1-211">See also</span></span>

- [<span data-ttu-id="465f1-212">閲覧フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="465f1-212">Create Outlook add-ins for read forms</span></span>](read-scenario.md)
- [<span data-ttu-id="465f1-213">Outlook アイテムからエンティティ文字列を抽出する</span><span class="sxs-lookup"><span data-stu-id="465f1-213">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)
- [<span data-ttu-id="465f1-214">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="465f1-214">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="465f1-215">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="465f1-215">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="465f1-216">Outlook アドインのアクセス許可を理解する</span><span class="sxs-lookup"><span data-stu-id="465f1-216">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
