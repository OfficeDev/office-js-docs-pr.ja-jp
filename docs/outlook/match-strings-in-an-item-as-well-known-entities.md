---
title: Outlook アドインで既知のエンティティとして文字列を照合する
description: JavaScript API Office使用すると、特定の既知のエンティティに一致する文字列を取得して、さらに処理できます。
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: bc95b53b0b94574f2db3543d9c8b1b3f6eecae0b87f82922a4ad245b8ad84575
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097909"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a>Outlook アイテム内の文字列を既知のエンティティとして照合する

メッセージおよび会議出席依頼のアイテムを送信する前に、Exchange Server によりアイテムの内容が解析され、件名と本文から、メール アドレス、電話番号、URL など、Exchange にとっての既知のエンティティに似た文字列が特定され、スタンプが付けられます。メッセージと会議出席依頼は、Exchange Server によって、既知のエンティティにスタンプが付けられた状態で、Outlook の受信トレイに配信されます。 

JavaScript API Office使用すると、特定の既知のエンティティと一致するこれらの文字列を取得して、さらに処理することができます。 さらに、既知のエンティティをアドイン マニフェスト内のルールで指定して、ユーザーがそのエンティティと一致するものを含んだアイテムを表示したときに、Outlook がアドインをアクティブにするように設定することもできます。 その後で、エンティティと一致するものを抽出してアクションを実行することができます。 

選択されたメッセージや予定からこれらのインスタンスを特定したり抽出したりできるので便利です。 たとえば、Outlook のアドインとして電話番号の逆引き検索サービスを作成できます。 このアドインは、アイテムの件名や本文から電話番号に似た文字列を抽出して逆引き検索を行い、各電話番号の登録所有者を表示させることができます。

このトピックでは既知のエンティティ、既知のエンティティに基づくアクティブ化ルールの例、およびアクティブ化ルール内でエンティティが使用されているかどうかに関係なく、一致するエンティティを抽出する方法を紹介します。


## <a name="support-for-well-known-entities"></a>既知のエンティティに対するサポート

Exchange Server は、ユーザーがメッセージや会議出席依頼アイテムを送信した後、それが受信者に配信される前に、アイテム内の既知のエンティティにスタンプを付けます。そのため、Exchange 内のトランスポートを通過したアイテムだけにスタンプが付けられ、Outlook はユーザーがそのようなアイテムを表示中にそれらのスタンプに基づいてアドインをアクティブにすることができます。しかし、ユーザーがアイテムを作成している間や、送信済みアイテム フォルダー内のアイテムを表示しているときは、そのアイテムがまだトランスポートを通過していないため、Outlook は既知のエンティティに基づいてアドインをアクティブにすることができません。 

同様に、作成中または送信済みアイテム フォルダー内のアイテムはトランスポートを通過しておらず、スタンプが付けられていないため、既知のエンティティを抽出できません。アクティブ化をサポートしているアイテムの種類の詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。

次の表は、Exchange Server と Outlook でサポートされ、認識されるエンティティ (つまり、「既知のエンティティ」) と、各エンティティのインスタンスのオブジェクト タイプを一覧にしたものです。これらのエンティティの 1 つとしての文字列の自然言語認識は、大量のデータに対してトレーニングを行った学習モデルに基づきます。したがって、認識は決定論的ではありません。認識に関する条件の詳細については、「 [既知のエンティティを使用するためのヒント](#tips-for-using-well-known-entities)」を参照してください。

**表 1.サポートされるエンティティとその型**

|エンティティの型|認識に関する条件|オブジェクトの種類|
|:-----|:-----|:-----|
|**住所**|米国の住所。次はその例です。1234 Main Street, Redmond, WA 07722.通常、住所が認識されるには、米国の住所の構造に従う必要があり、ほとんどには番地、住所、都市名、州名、郵便番号の要素が存在します。住所は 1 行または複数行で指定できます。|JavaScript **String** オブジェクト|
|**連絡先**|自然言語の認識による、人に関する情報の参照。 連絡先の認識は、状況によりさまざまな方法で行われます。 たとえば、メッセージの最後にある署名や、人の名前の近くに現れる電話番号、住所、メール アドレス、URL などの情報です。|[Contact](/javascript/api/outlook/office.contact) オブジェクト|
|**EmailAddress**|SMTP メール アドレス。|JavaScript `String` オブジェクト|
|**MeetingSuggestion**|イベントまたは会議の参照。たとえば、Exchange 2013では次のテキストは会議の提案として認識されます。 _明日、昼食会議を開きましょう。_|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) オブジェクト|
|**PhoneNumber**|米国の電話番号。次はその例です。_(235) 555-0110_|[PhoneNumber](/javascript/api/outlook/office.phonenumber) オブジェクト|
|**TaskSuggestion**|電子メールの対応可能な文言。たとえば、_スプレッドシートを更新してください。_|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) オブジェクト|
|**Url**|ネットワーク ロケーションと Web リソースの識別子を明記した Web アドレス。 Exchange Serverは、Web アドレスにアクセス プロトコルを必要としないし、リンク テキストに埋め込まれている URL をエンティティのインスタンスとして認識 `Url` しない。 Exchange Serverは、次の例と一致できます。 `www.youtube.com/user/officevideos``https://www.youtube.com/user/officevideos` |JavaScript `String` オブジェクト|

<br/>

次の図は、アドインで Exchange Server と Outlook が既知のエンティティをサポートする仕組みと、既知のエンティティを使用してアドインでできる操作について説明しています。エンティティの利用方法について詳しくは、「[アドインでのエンティティの取得](#retrieving-entities-in-your-add-in)」と「[エンティティの存在に基づくアドインのアクティブ化](#activating-an-add-in-based-on-the-existence-of-an-entity)」をご覧ください。

**Exchange Server、Outlook、アドインが既知のエンティティをサポートする仕組み**

![メール アプリでの既知のエンティティのサポートと使用。](../images/well-known-entities-info.png)


## <a name="permissions-to-extract-entities"></a>エンティティを抽出するためのアクセス許可

JavaScript コードでエンティティを抽出したり、特定の既知のエンティティの存在に基づいてアドインをアクティブ化したりする場合は、アドイン マニフェストで適切なアクセス許可を要求しておきます。

既定の制限付きアクセス許可を指定すると、アドインで `Address` 、 、または `MeetingSuggestion` エンティティを抽出 `TaskSuggestion` できます。 その他のエンティティを抽出するには、開封済みアイテム、読み取り/書き込みアイテム、またはメールボックスの読み取り/書き込み許可を指定します。 マニフェストでこれを行うには [、Permissions](../reference/manifest/permissions.md)要素を使用し、次の例のように、適切なアクセス許可 &mdash; Restricted、ReadItem、ReadWriteItem、**または ReadWriteMailbox** &mdash; を指定します。

```xml
<Permissions>ReadItem</Permissions>
```


## <a name="retrieving-entities-in-your-add-in"></a>アドインでのエンティティの取得

ユーザーが表示するアイテムの件名または本文に、Exchange および Outlook が既知のエンティティとして認識できる文字列が含まれている限り、これらのインスタンスはアドインで使用できます。これらは、既知のエンティティに基づいてアドインがアクティブ化されていない場合でも使用できます。 適切なアクセス許可を使用すると、or メソッドを使用して、現在のメッセージまたは予定に存在する既知のエンティティ `getEntities` `getEntitiesByType` を取得できます。

この `getEntities` メソッドは、アイテム内のすべての既知のエンティティを含む [Entities](/javascript/api/outlook/office.entities) オブジェクトの配列を返します。

特定の種類のエンティティに興味がある場合は、必要なエンティティの配列のみを返すメソッド `getEntitiesByType` を使用します。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 列挙型は抽出可能なすべての既知のエンティティの種類を表します。

呼び出した後、オブジェクトの対応するプロパティを使用して、エンティティの種類のインスタンスの `getEntities` `Entities` 配列を取得できます。 エンティティの型により、配列内のインスタンスは単なる文字列であることも、特定のオブジェクトにマップできることもあります。 

たとえば、前出の図のように、アイテムのアドレスを取得するには、`getEntities().addresses[]` により返される配列にアクセスします。 この `Entities.addresses` プロパティは、郵便番号として認識Outlook文字列の配列を返します。 同様に、 `Entities.contacts` このプロパティは、連絡先情報として認識Outlook `Contact` オブジェクトの配列を返します。 表 1 に、サポートされる各エンティティのインスタンスのオブジェクト型を示します。

以下の例では、メッセージ内で見つかった住所を取得する方法を示します。

```js
// Get the address entities from the item.
var entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}

```


## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a>エンティティの存在に基づくアドインのアクティブ化

既知のエンティティを利用するもう 1 つの方法は、現在表示されているアイテムの件名または本文に 1 つまたは複数の種類のエンティティが存在するかどうかに基づいて Outlook にアドインをアクティブ化させる方法です。 これを行うには、アドイン マニフェスト `ItemHasKnownEntity` でルールを指定します。 [EntityType 単純](/javascript/api/outlook/office.mailboxenums.entitytype)型は、ルールでサポートされる既知のエンティティの種類を表 `ItemHasKnownEntity` します。 アドインがアクティブ化されたら、前のセクション「 [アドインでのエンティティの取得](#retrieving-entities-in-your-add-in)」で説明したように、目的のエンティティのインスタンスを取得することもできます。

必要に応じて、ルールに正規表現を適用して、エンティティのインスタンスをさらにフィルター処理し、Outlook がエンティティのインスタンスのサブセットでのみアドインをアクティブ化できます `ItemHasKnownEntity` 。 たとえば、"98" で始まるワシントン州の郵便番号を含むメッセージの中の街路住所エンティティを検出するフィルターを指定できます。 エンティティ インスタンスにフィルターを適用するには `RegExFilter` `FilterName` `Rule` [、ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 型の要素の and 属性を使用します。

他のアクティブ化ルールと同様に、複数のルールを指定してアドインのルール コレクションを作成できます。 次の例では、ルールとルールの 2 つのルールに "AND" `ItemIs` 操作を適用 `ItemHasKnownEntity` します。 このルール コレクションにより、現在のアイテムがメッセージである場合に、Outlook がそのアイテムの件名または本文から住所を認識すると、アドインがアクティブ化されます。

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

<br/>

次の使用例は `getEntitiesByType` 、現在のアイテムを使用して、前のルール コレクションの結果に変数 `addresses` を設定します。

```js
var addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

<br/>

次のルール例は、現在のアイテムの件名または本文に URL が含まれている場合は常にアドインをアクティブ化し、URL には文字列の場合に関係なく、"youtube" という文字列が含 `ItemHasKnownEntity` まれます。

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

<br/>

次の使用例は、現在のアイテムを使用して変数を設定し、前のルールの正規表現に一致する結果の配列 `getFilteredEntitiesByName(name)` `videos` を取得 `ItemHasKnownEntity` します。

```js
var videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```


## <a name="tips-for-using-well-known-entities"></a>既知のエンティティを使用するためのヒント

アドインで既知のエンティティを使用する場合に知っておくべきいくつかの事実と制限があります。 ルールを使用するかどうかに関係なく、ユーザーが既知のエンティティの一致を含むアイテムを読み取っているときにアドインがアクティブ化されている限り、以下が適用 `ItemHasKnownEntity` されます。


- 文字列が英語の場合にのみ、既知のエンティティである文字列を抽出できます。
    
- アイテム本文の最初の 2,000 文字から既知のエンティティを抽出できます。2,000 を超える文字からは抽出できません。 このサイズ制限により機能とパフォーマンスのニーズのバランスが維持されるため、サイズの大きなメッセージと予定の中から既知のエンティティのインスタンスの解析と特定をしても、Exchange Server と Outlook は停止しません。 この制限は、アドインがルールを指定するかどうかに依存しない点に注意 `ItemHasKnownEntity` してください。 アドインでそのようなルールを使用する場合には、Outlook リッチ クライアントに対する以下の 2 番目の項目のルール処理制限にも注意してください。
    
- メールボックスの所有者以外の誰かが計画した会議である予定からエンティティを抽出できます。会議ではないカレンダー アイテムやメールボックスの所有者が計画した会議である予定からエンティティを抽出することはできません。
    
- 種類のエンティティはメッセージのみから抽出できますが `MeetingSuggestion` 、予定は抽出しません。
    
- アイテム本文に明示的に存在する URL を抽出することはできますが、HTML のアイテム本文のハイパーリンク テキストに埋め込まれている URL を抽出することはできません。 明示的な URL と埋め込み URL の両方を取得するには、代わりにルール `ItemHasRegularExpressionMatch` を使用してください。 `BodyAsHTML`PropertyName _として指定し_、URL と一致する正規表現を _RegExValue として指定します_。
    
- [送信済みアイテム] フォルダーのアイテムからエンティティを抽出することはできません。
    
さらに [、ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用する場合は、次のことが適用され、アドインがアクティブ化されると予想されるシナリオに影響を与える可能性があります。

- ルールを使用する場合、マニフェストOutlook既定のロケールに関係なく、エンティティ文字列が英語でのみ一致する必要 `ItemHasKnownEntity` があります。
    
- アドインが Outlook リッチ クライアントで実行されている場合は、Outlook がアイテム本文の最初のメガバイトにルールを適用し、その制限を超える残りの本文には適用されないと予想します。 `ItemHasKnownEntity`
    
- ルールを使用して、[送信されたアイテム] フォルダー内のアイテムのアドイン `ItemHasKnownEntity` をアクティブ化することはできません。
    

## <a name="see-also"></a>関連項目

- [閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)
- [Outlook アイテムからエンティティ文字列を抽出する](extract-entity-strings-from-an-item.md)
- [Outlook アドインのアクティブ化ルール](activation-rules.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アドインのアクセス許可を理解する](understanding-outlook-add-in-permissions.md)
