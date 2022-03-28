---
title: Outlook アドインのアクティブ化ルール
description: Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。
ms.date: 12/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6af5003108efede78fa06b220abe3c89d472c6e5
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484114"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Outlook コンテキスト アドインのアクティブ化ルール

Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。これは、1.1 マニフェストのスキーマを使用するすべてのアドインについて同様です。ユーザーは、Outlook UI からアドインを選び、現在のアイテムに、そのアドインを起動することができます。

次の図は、閲覧ウィンドウにあるアドイン バーでアクティブ化されたメッセージ用の Outlook アドインを示しています。

![アクティブ化された読み取りメール アプリを表示するアプリ バー。](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>マニフェストでのアクティブ化ルールの指定


特定のOutlookをアクティブ化するには、次のいずれかの要素を使用してアドイン マニフェストでアクティブ化ルールを指定`Rule`します。

- [Rule 要素 (MailApp complexType)](/javascript/api/manifest/rule) - 個別のルールを指定します。
- [Rule 要素 (RuleCollection complexType)](/javascript/api/manifest/rule#rulecollection) - 論理演算子を使用して複数のルールを結合します。


 > [!NOTE]
 > 個々 `Rule` のルールを指定するために使用する要素は、抽象 [Rule](/javascript/api/manifest/rule) 複合型です。 次の各種類のルールは、この抽象複合型を `Rule` 拡張します。 したがって、マニフェストで個別のルールを指定するときは、[xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性を使用してルールの以下の型の 1 つをさらに定義する必要があります。
 > 
 > たとえば、次のルールは [ItemIs ルールを定義](/javascript/api/manifest/rule#itemis-rule) します。
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > 属性 `FormType` はマニフェスト v1.1 `VersionOverrides` のアクティブ化ルールに適用されますが、v1.0 では定義されていません。 したがって、 [ItemIs](/javascript/api/manifest/rule#itemis-rule) がノードで使用されている場合は使用 `VersionOverrides` できません。

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

ItemIs ルールの属性で、次のいずれかの `ItemType` アイテムの種類 **を指定** します。 マニフェストでは、複数の **ItemIs** ルールを指定できます。 ItemType simpleType では、Outlook アドインをサポートしている Outlook アイテムの種類を定義します。

<br/>

|**値**|**説明**|
|:-----|:-----|
|**Appointment**|Outlook の予定表内のアイテムを指定します。 このアイテムには、開催者と出席者を持つ応答済みの会議アイテムと、開催者と出席者を持たない、単なる予定表上のアイテムである予定が含まれます。 これは Outlook の IPM.Appointment メッセージ クラスに対応します。|
|**メッセージ**|通常受信トレイで受信される次のいずれかの項目を指定します。 <ul><li><p>電子メール メッセージ。これは Outlook の IPM.Note メッセージ クラスに対応します。</p></li><li><p>会議出席依頼、返信、または取り消し。 これは、次のメッセージ クラスに対応Outlook。</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

この `FormType` 属性を使用して、アドインをアクティブにするモード (読み取りまたは作成) を指定します。


 > [!NOTE]
 > ItemIs 属性 `FormType` はスキーマ v1.1 `VersionOverrides` 以降で定義されますが、v1.0 では定義されません。 アドイン コマンドを定義するときに `FormType` 属性を含めない。

アドインがアクティブ化された後は、 [mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) プロパティを使用して Outlook で現在選択されているアイテムを取得し、 [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティを使用して現在のアイテムの種類を取得できます。

必要に応じて`ItemClass``IncludeSubClasses`、属性を使用してアイテムのメッセージ クラスを指定し、属性を使用して、アイテムが指定されたクラスのサブクラスである場合にルールを **true** にするかどうかを指定できます。

メッセージ クラスの詳細については、「[Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)」をご覧ください。

次の例は、ユーザーがメッセージを読んでいるときにアドインを Outlookバーに表示できる **ItemIs** ルールです。

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


複合 `ItemHasAttachment` 型は、選択したアイテムに添付ファイルが含まれている場合にチェックするルールを定義します。

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

アイテムをアドインで使用できる前に、サーバーはアドインを調べて、件名と本文に既知のエンティティの 1 つである可能性があるテキストが含まれているかどうかを判断します。 これらのエンティティが見つかった`getEntities``getEntitiesByType`場合は、そのアイテムの or メソッドを使用してアクセスする既知のエンティティのコレクションに配置されます。

指定した種類のエンティティ `ItemHasKnownEntity` がアイテムに存在する場合にアドインを表示するルールを使用して指定できます。 ルールの属性には、次の既知の `EntityType` エンティティを指定 `ItemHasKnownEntity` できます。

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL

必要に応じて `RegularExpression` 、属性に正規表現を含め、現在の正規表現と一致するエンティティの場合にのみアドインが表示されます。 ルールで指定された正規表現に`ItemHasKnownEntity``getRegExMatches`一致する文字列を取得するには、現在選択されているアイテムアイテムに or `getFilteredEntitiesByName` メソッドOutlookできます。

次の例は、 `Rule` 指定された既知のエンティティの 1 つがメッセージに存在する場合にアドインを表示する要素のコレクションを示しています。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

次の例は、`ItemHasKnownEntity``RegularExpression`"contoso" という単語を含む URL がメッセージ内に存在する場合にアドインをアクティブ化する属性を持つルールを示しています。


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

アクティブ化ルールのエンティティの詳細については、「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」を参照してください。


## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

複合 `ItemHasRegularExpressionMatch` 型は、正規表現を使用してアイテムの指定されたプロパティの内容と一致するルールを定義します。 正規表現に一致するテキストがアイテムの指定プロパティ内に見つかった場合に、Outlook はアドイン バーをアクティブ化してそのアドインを表示します。 現在選択されているアイテム `getRegExMatches` を `getRegExMatchesByName` 表すオブジェクトの or メソッドを使用して、指定した正規表現の一致を取得できます。

次の例は、 `ItemHasRegularExpressionMatch` 選択したアイテムの本文に大文字と小文字を無視して、"apple"、"banana"、または "ココナッツ" が含まれている場合にアドインをアクティブ化する例を示しています。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

ルールの使用の詳細については、「`ItemHasRegularExpressionMatch`[Use regular expression activation rules to show a Outlookアドイン」を参照してください](use-regular-expressions-to-show-an-outlook-add-in.md)。


## <a name="rulecollection-rule"></a>RuleCollection ルール


複合型 `RuleCollection` は、複数のルールを 1 つのルールに結合します。 属性を使用して、コレクション内のルールを論理 OR または論理 AND と組み合わせるかどうかを指定 `Mode` できます。

論理 AND を指定する場合、アドインは、コレクション内で指定されているすべてのルールにアイテムが一致する場合にのみ表示されます。論理 OR を指定する場合は、コレクションで指定されているルールのいずれか 1 つにでもアイテムが一致すれば、アドインは表示されます。

ルールを組み合わせて `RuleCollection` 複雑なルールを形成できます。 次に示す例では、件名や本文に住所が含まれるメッセージまたは予定表のアイテムをユーザーが表示したときに、アドインがアクティブ化されます。

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


このアドインを十分にOutlookするには、ライセンス認証と API の使用ガイドラインに従う必要があります。 次の表に、正規表現とルールの一般的な制限を示しますが、アプリケーションごとに特定のルールがあります。 詳細については、「ライセンス認証の制限」および「[JavaScript API for Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」および「Outlookアドインのライセンス認証[」を参照してください](troubleshoot-outlook-add-in-activation.md)。

<br/>

|**アドインの要素**|**ガイドライン**|
|:-----|:-----|
|マニフェストのサイズ|256 KB 未満。|
|ルール|15 ルール未満。|
|ItemHasKnownEntity|Outlook リッチ クライアントでは、本文の最初の 1 MB にルールを適用し、残りの部分には適用しません。|
|正規表現|すべてのアプリケーションの ItemHasKnownEntity または ItemHasRegularExpressionMatch Outlookします。<br><ul><li>Outlook アドインのアクティブ化ルールで指定する正規表現は 5 個までにしてください。その制約数を超えるアドインをインストールすることはできません。</li><li>予期される結果が <b>getRegExMatches</b> メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。 </li><li>**重要**: 正規表現に一致した文字列に基づいて、テキストが強調表示されます。 ただし、強調表示されたオ`(?!text)``(?<=text)`カレンスは、負の先取り、後ろ向き、負の後ろ向きなど、実際の正規表現アサーションの結果と正確には一致しない場合があります`(?<!text)`。 たとえば、 `under(?!score)` "Like under, under score, and アンダースコア" で正規表現を使用すると、最初の 2 つの文字列ではなく、すべてのオカレンスに対して文字列 "under" が強調表示されます。</li><li>次の表で、一致する制限を超えない正規表現を指定します。<br/><br/><table><tr><th>正規表現の長さ制限</th><th>Outlook リッチ クライアント</th><th>iOS および Android 用の Outlook</th></tr><tr><td>アイテムの本文がテキスト形式の場合</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>アイテムの本文が HTML の場合</td><td>3 KB</td><td>3 KB</td></tr></table>|

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
    
