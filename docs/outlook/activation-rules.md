---
title: Outlook アドインのアクティブ化ルール
description: Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: cdcdfbf3961ad9f627ba00f7366f49c77bba435d
ms.sourcegitcommit: fd110305c2be8660ab8a47c1da3e3969bd1ede86
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/23/2020
ms.locfileid: "48214597"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Outlook コンテキスト アドインのアクティブ化ルール

Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。これは、1.1 マニフェストのスキーマを使用するすべてのアドインについて同様です。ユーザーは、Outlook UI からアドインを選び、現在のアイテムに、そのアドインを起動することができます。

次の図は、閲覧ウィンドウにあるアドイン バーでアクティブ化されたメッセージ用の Outlook アドインを示しています。

![メール読み取りアプリがアクティブ化されたことを示すアプリ バー](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>マニフェストでのアクティブ化ルールの指定


Outlook で特定の条件に応じてアドインをアクティブ化するには、次のいずれかの要素を使用して、アドインマニフェストでアクティブ化ルールを指定し `Rule` ます。

- [Rule 要素 (MailApp complexType)](../reference/manifest/rule.md) - 個別のルールを指定します。
- [Rule 要素 (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - 論理演算子を使用して複数のルールを結合します。
    

 > [!NOTE]
 > `Rule`個別のルールを指定するために使用する要素は、抽象[ルール](../reference/manifest/rule.md)複合型です。 次のルールの各型は、この抽象 `Rule` 複合型を拡張します。 したがって、マニフェストで個別のルールを指定するときは、[xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性を使用してルールの以下の型の 1 つをさらに定義する必要があります。
 > 
 > たとえば、次のルールは [ItemIs](../reference/manifest/rule.md#itemis-rule) ルールを定義します。`<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > この属性は、マニフェスト v1.1 の `FormType` アクティブ化ルールに適用されますが、v2.0 では定義されていません `VersionOverrides` 。 そのため、ノードで [Itemis](../reference/manifest/rule.md#itemis-rule) が使用されている場合は使用できません `VersionOverrides` 。

次の表は、使用できるルールの種類を示しています。詳細については、この表の後の説明と、「[閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)」の該当記事を参照してください。

<br/>

|**ルール名**|**該当するフォーム**|**説明**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|読み取り、作成|現在選択されているアイテムは指定された種類のアイテム (メッセージまたは予定) かどうかを調べます。また、アイテム クラス、フォームの種類、さらにはオプションでアイテム メッセージ クラスも調べることができます。|
|[ItemHasAttachment](#itemhasattachment-rule)|読み取り|選択されているアイテムに添付ファイルが含まれるかどうかを調べます。|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|読み取り|選択されているアイテムに 1 つ以上の一般的なエンティティが含まれるかどうかを調べます。詳細: 「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」。|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|読み取り|選択されているアイテムの送信者のメール アドレス、件名、本文に正規表現と一致するものが含まれるかどうかを調べます。詳細: [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)|
|[RuleCollection](#rulecollection-rule)|読み取り、作成|複数のルールを組み合わせて、より複雑なルールを作成できます。|

## <a name="itemis-rule"></a>ItemIs ルール

**ItemIs** 複合型は、現在のアイテムがアイテムの種類と一致している場合 (また、オプションとしてルールに明記されている場合はアイテムのメッセージ クラスとも一致している場合) に **true** と評価されるルールを定義します。

`ItemType` **Itemis**ルールの属性に、次のいずれかのアイテムの種類を指定します。 マニフェストでは、複数の **ItemIs** ルールを指定できます。 ItemType simpleType では、Outlook アドインをサポートしている Outlook アイテムの種類を定義します。

<br/>

|**値**|**説明**|
|:-----|:-----|
|**Appointment**|Outlook の予定表内のアイテムを指定します。このアイテムには、開催者と出席者を持つ応答済みの会議アイテムと、開催者と出席者を持たない、単なる予定表上のアイテムである予定が含まれます。これは Outlook の IPM.Appointment メッセージ クラスに対応します。|
|**Message**|通常は受信トレイで受信される次のアイテムのいずれかを指定します。 <ul><li><p>電子メール メッセージ。これは Outlook の IPM.Note メッセージ クラスに対応します。</p></li><li><p>会議出席依頼、返信、または取り消し。Outlook の次のメッセージ クラスに対応します。</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

この属性を使用して、 `FormType` アドインをアクティブ化するモード (読み取りまたは新規作成) を指定します。


 > [!NOTE]
 > ItemIs 属性は、スキーマ v1.1 以降では定義されていますが、v1.0 では定義されてい `FormType` ません `VersionOverrides` 。 `FormType`アドインコマンドを定義するときに、属性を含めないでください。

アドインがアクティブ化された後は、 [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) プロパティを使用して Outlook で現在選択されているアイテムを取得し、 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティを使用して現在のアイテムの種類を取得できます。

必要に応じて、属性を使用して `ItemClass` アイテムのメッセージクラスを指定し、その `IncludeSubClasses` アイテムが指定したクラスのサブクラスである場合にルールを **true** にする必要があるかどうかを指定する属性を指定できます。

メッセージ クラスの詳細については、「[Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)」をご覧ください。

次の例は、ユーザーがメッセージを読むときに Outlook のアドイン バーにアドインを表示する **ItemIs** ルールを示しています。

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

次の例は、ユーザーがメッセージまたは予約を閲覧するときに Outlook のアドイン バーにアドインを表示する **ItemIs** ルールを示しています。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a>ItemHasAttachment ルール


`ItemHasAttachment`複合型は、選択されているアイテムに添付ファイルが含まれているかどうかを確認するルールを定義します。

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

アイテムがアドインで使用可能になる前に、サーバーはそれを調べて、件名と本文に既知のエンティティのいずれかであると考えられるテキストが含まれているかどうかを判断します。これらのエンティティのいずれかが見つかった場合は、その `getEntities` アイテムのまたはメソッドを使用してアクセスする既知のエンティティのコレクションに配置され `getEntitiesByType` ます。

指定した `ItemHasKnownEntity` 型のエンティティがアイテム内に存在する場合にアドインを表示するルールを指定できます。次の既知のエンティティを `EntityType` ルールの属性に指定でき `ItemHasKnownEntity` ます。

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL
    
必要に応じて、属性に正規表現を含めることができ `RegularExpression` ます。これにより、現在の正規表現に一致するエンティティがある場合にアドインが表示されるようになります。ルールで指定された正規表現に一致するものを取得するに `ItemHasKnownEntity` `getRegExMatches` は、現在選択されている Outlook アイテムに対して or メソッドを使用でき `getFilteredEntitiesByName` ます。

次の例は、 `Rule` 指定された既知のエンティティのいずれかがメッセージに存在するときにアドインを表示する要素のコレクションを示しています。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

次の例は、 `ItemHasKnownEntity` `RegularExpression` "contoso" という単語を含む URL がメッセージ内に存在するときにアドインをアクティブ化する属性を持つルールを示しています。


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

アクティブ化ルールのエンティティの詳細については、「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」を参照してください。


## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

`ItemHasRegularExpressionMatch`複合型は、アイテムの指定されたプロパティの内容と一致するように正規表現を使用するルールを定義します。正規表現と一致するテキストがアイテムの指定されたプロパティに含まれている場合、Outlook はアドインバーをアクティブにして、アドインを表示します。`getRegExMatches`現在選択されているアイテムを表すオブジェクトのまたはメソッドを使用して、指定した `getRegExMatchesByName` 正規表現に一致するものを取得できます。

次の例は、 `ItemHasRegularExpressionMatch` 選択したアイテムの本文に "apple"、"banana"、または "coconut" が含まれている場合にアドインをアクティブにする方法を示しています。大文字小文字は無視されます。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

ルールの使用の詳細について `ItemHasRegularExpressionMatch` は、「 [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)」を参照してください。


## <a name="rulecollection-rule"></a>RuleCollection ルール


`RuleCollection`複合型は、複数のルールを1つのルールに結合します。コレクション内のルールを論理 OR または論理と組み合わせて使用するかどうかを指定でき `Mode` ます。属性を使用します。

論理 AND を指定する場合、アドインは、コレクション内で指定されているすべてのルールにアイテムが一致する場合にのみ表示されます。論理 OR を指定する場合は、コレクションで指定されているルールのいずれか 1 つにでもアイテムが一致すれば、アドインは表示されます。

ルールを結合して `RuleCollection` 複雑なルールを形成できます。次の例では、ユーザーが予定またはメッセージアイテムを表示していて、アイテムの件名または本文に住所が含まれている場合にアドインをアクティブにします。

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

次の例では、ユーザーがメッセージを新規作成するときか、件名か本文に住所が含まれる予定を表示するときに、アドインがアクティブ化されます。

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a>ルールと正規表現の制約事項


Outlook アドインでの満足感を得るには、アクティブ化と API の使用に関するガイドラインに従う必要があります。次の表に、正規表現とルールの一般的な制限を示しますが、アプリケーションごとに固有のルールがあります。詳細については、「 [outlook アドインのアクティブ化と JAVASCRIPT API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) 」および「 [outlook アドインのアクティブ化のトラブルシューティング](troubleshoot-outlook-add-in-activation.md)」を参照してください。

<br/>

|**アドインの要素**|**ガイドライン**|
|:-----|:-----|
|マニフェストのサイズ|256 KB 未満。|
|ルール|15 ルール未満。|
|ItemHasKnownEntity|Outlook リッチ クライアントでは、本文の最初の 1 MB にルールを適用し、残りの部分には適用しません。|
|正規表現|すべての Outlook アプリケーションの ItemHasKnownEntity または ItemHasRegularExpressionMatch ルールの場合:<br><ul><li>Outlook アドインのアクティブ化ルールで指定する正規表現は 5 個までにしてください。その制約数を超えるアドインをインストールすることはできません。</li><li>予期される結果が <b>getRegExMatches</b> メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。 </li><li>正規表現で先読みアサーションは指定しますが、後読み `(?<=text)` および否定の後読み `(?<!text)` アサーションは指定しません。</li><li>一致数が次の表の制限を超えない正規表現を指定します。<br/><br/><table><tr><th>正規表現の長さ制限</th><th>Outlook リッチ クライアント</th><th>iOS および Android 用の Outlook</th></tr><tr><td>アイテムの本文がテキスト形式の場合</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>アイテムの本文が HTML の場合</td><td>3 KB</td><td>3 KB</td></tr></table>|

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
    
