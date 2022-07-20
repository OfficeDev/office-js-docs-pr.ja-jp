---
title: Outlook アドインのアクティブ化ルール
description: Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。
ms.date: 12/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: af9edf0254156d7bdac13d0553036a614d8c4c39
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889640"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>Outlook コンテキスト アドインのアクティブ化ルール

Outlook では、ユーザーが読み取りや作成をしようとしているメッセージまたは予定が、アドインのアクティブ化のルールに準ずる場合に、ある種類のアドインをアクティブにします。これは、1.1 マニフェストのスキーマを使用するすべてのアドインについて同様です。ユーザーは、Outlook UI からアドインを選び、現在のアイテムに、そのアドインを起動することができます。

次の図は、閲覧ウィンドウにあるアドイン バーでアクティブ化されたメッセージ用の Outlook アドインを示しています。

![アクティブ化された読み取りメール アプリを示すアプリ バー。](../images/read-form-app-bar.png)

## <a name="specify-activation-rules-in-a-manifest"></a>マニフェストでのアクティブ化ルールの指定

Outlook で特定の条件に対してアドインをアクティブ化するには、次 `Rule` のいずれかの要素を使用して、アドイン マニフェストでアクティブ化規則を指定します。

- [Rule 要素 (MailApp complexType)](/javascript/api/manifest/rule) - 個別のルールを指定します。
- [Rule 要素 (RuleCollection complexType)](/javascript/api/manifest/rule#rulecollection) - 論理演算子を使用して複数のルールを結合します。

 > [!NOTE]
 > 個々のルールを指定するために使用する要素は `Rule` 、抽象 [Rule](/javascript/api/manifest/rule) 複合型です。 次の各種類のルールは、この抽象 `Rule` 複合型を拡張します。 したがって、マニフェストで個別のルールを指定するときは、[xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性を使用してルールの以下の型の 1 つをさらに定義する必要があります。
 >
 > たとえば、次のルールは [ItemIs](/javascript/api/manifest/rule#itemis-rule) ルールを定義します。
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 >
 > この属性は `FormType` マニフェスト v1.1 のアクティブ化規則に適用されますが、v1.0 では `VersionOverrides` 定義されていません。 そのため、ノードで `VersionOverrides` [ItemIs が](/javascript/api/manifest/rule#itemis-rule)使用されている場合は使用できません。

次の表は、使用できるルールの種類を示しています。詳細については、この表の後の説明と、「[閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)」の該当記事を参照してください。

|**ルール名**|**該当するフォーム**|**説明**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|読み取り、作成|現在選択されているアイテムは指定された種類のアイテム (メッセージまたは予定) かどうかを調べます。また、アイテム クラス、フォームの種類、さらにはオプションでアイテム メッセージ クラスも調べることができます。|
|[ItemHasAttachment](#itemhasattachment-rule)|読み取り|選択されているアイテムに添付ファイルが含まれるかどうかを調べます。|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|読み取り|選択されているアイテムに 1 つ以上の一般的なエンティティが含まれるかどうかを調べます。詳細: 「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」。|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|読み取り|選択されているアイテムの送信者のメール アドレス、件名、本文に正規表現と一致するものが含まれるかどうかを調べます。詳細: [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)|
|[RuleCollection](#rulecollection-rule)|読み取り、作成|複数のルールを組み合わせて、より複雑なルールを作成できます。|

## <a name="itemis-rule"></a>ItemIs ルール

複合型は `ItemIs` 、現在の項目が項目の種類と一致するかどうかを評価する `true` ルールを定義し、必要に応じてアイテム メッセージ クラスがルールに記載されている場合はアイテム メッセージ クラスを定義します。

ルールの属性で `ItemType` 、次のいずれかの項目の種類を `ItemIs` 指定します。 マニフェストでは、複数 `ItemIs` のルールを指定できます。 ItemType simpleType では、Outlook アドインをサポートしている Outlook アイテムの種類を定義します。

|**値**|**説明**|
|:-----|:-----|
|**Appointment**|Outlook の予定表内のアイテムを指定します。 このアイテムには、開催者と出席者を持つ応答済みの会議アイテムと、開催者と出席者を持たない、単なる予定表上のアイテムである予定が含まれます。 これは Outlook の IPM.Appointment メッセージ クラスに対応します。|
|**Message**|受信トレイで通常受信する次の項目のいずれかを指定します。 <ul><li><p>電子メール メッセージ。 これは Outlook の IPM.Note メッセージ クラスに対応します。</p></li><li><p>会議出席依頼、返信、または取り消し。 これは、Outlook の次のメッセージ クラスに対応します。</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

この `FormType` 属性は、アドインをアクティブにするモード (読み取りまたは作成) を指定するために使用されます。

 > [!NOTE]
 > ItemIs `FormType` 属性はスキーマ v1.1 以降で定義されますが、v1.0 では `VersionOverrides` 定義されません。 アドイン コマンドを定義するときは、 `FormType` 属性を含めないでください。

アドインがアクティブ化された後は、 [mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) プロパティを使用して Outlook で現在選択されているアイテムを取得し、 [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティを使用して現在のアイテムの種類を取得できます。

必要に応じて、この属性を `ItemClass` 使用してアイテムのメッセージ クラスを指定し、属性を `IncludeSubClasses` 使用して、アイテムが指定されたクラスのサブクラスである場合にルールを指定する必要 `true` があるかどうかを指定できます。

メッセージ クラスの詳細については、「[Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)」をご覧ください。

次の例は、 `ItemIs` ユーザーがメッセージを読み取っているときに Outlook アドイン バーにアドインを表示できるようにするルールです。

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

次の例は、 `ItemIs` ユーザーがメッセージまたは予定を読み取っているときに Outlook アドイン バーにアドインを表示できるようにするルールです。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment ルール

複合型は `ItemHasAttachment` 、選択したアイテムに添付ファイルが含まれているかどうかを確認するルールを定義します。

```xml
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

アイテムをアドインで使用できるようになる前に、サーバーは、既知のエンティティの 1 つである可能性のあるテキストが件名と本文に含まれているかどうかを調べます。 これらのエンティティのいずれかが見つかった場合は、そのアイテムのメソッドを`getEntitiesByType`使用して`getEntities`アクセスする既知のエンティティのコレクションに配置されます。

指定した種類のエンティティがアイテムに存在する場合にアドインを表示するルールを使用 `ItemHasKnownEntity` して指定できます。 ルールの属性で、次の `EntityType` 既知のエンティティを `ItemHasKnownEntity` 指定できます。

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL

必要に応じて、正規表現を `RegularExpression` 属性に含めて、現在の正規表現と一致するエンティティの場合にのみアドインを表示できます。 ルールで`ItemHasKnownEntity`指定された正規表現に一致する文字列を取得するには、現在選択されている Outlook アイテムのメソッドを`getFilteredEntitiesByName`使用`getRegExMatches`します。

次の例は、指定した既知の `Rule` エンティティの 1 つがメッセージに存在する場合にアドインを示す要素のコレクションを示しています。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

次の例は、 `ItemHasKnownEntity` "contoso" という単語を `RegularExpression` 含む URL がメッセージ内に存在する場合にアドインをアクティブ化する属性を持つルールを示しています。

```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

アクティブ化ルールのエンティティの詳細については、「[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)」を参照してください。

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

複合型は `ItemHasRegularExpressionMatch` 、正規表現を使用してアイテムの指定されたプロパティの内容と一致するルールを定義します。 正規表現に一致するテキストがアイテムの指定プロパティ内に見つかった場合に、Outlook はアドイン バーをアクティブ化してそのアドインを表示します。 現在選択されている項目を `getRegExMatches` 表すオブジェクトのメソッドを `getRegExMatchesByName` 使用して、指定した正規表現の一致を取得できます。

次の例は、 `ItemHasRegularExpressionMatch` 選択したアイテムの本文に "apple"、"debugger"、または "debugger" が含まれている場合にアドインをアクティブ化する例を示しています。大文字と小文字は無視されます。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

ルールの `ItemHasRegularExpressionMatch` 使用の詳細については、「 [正規表現のアクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)」を参照してください。

## <a name="rulecollection-rule"></a>RuleCollection ルール

複合型は、複数の `RuleCollection` ルールを 1 つのルールに結合します。 属性を使用 `Mode` して、コレクション内のルールを論理 OR または論理 AND と組み合わせる必要があるかどうかを指定できます。

論理 AND を指定する場合、アドインは、コレクション内で指定されているすべてのルールにアイテムが一致する場合にのみ表示されます。論理 OR を指定する場合は、コレクションで指定されているルールのいずれか 1 つにでもアイテムが一致すれば、アドインは表示されます。

ルールを組み合わせて `RuleCollection` 、複雑なルールを形成できます。 次に示す例では、件名や本文に住所が含まれるメッセージまたは予定表のアイテムをユーザーが表示したときに、アドインがアクティブ化されます。

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

Outlook アドインで満足のいくエクスペリエンスを提供するには、アクティブ化と API の使用ガイドラインに従う必要があります。 次の表は、正規表現と規則の一般的な制限を示していますが、アプリケーションごとに固有の規則があります。 詳細については、「 [ライセンス認証の制限」と「Outlook アドイン用 JavaScript API」および「Outlook](limits-for-activation-and-javascript-api-for-outlook-add-ins.md) アドインの [アクティブ化のトラブルシューティング](troubleshoot-outlook-add-in-activation.md)」を参照してください。

|**アドインの要素**|**ガイドライン**|
|:-----|:-----|
|マニフェストのサイズ|256 KB 未満。|
|ルール|15 ルール未満。|
|ItemHasKnownEntity|Outlook リッチ クライアントでは、本文の最初の 1 MB にルールを適用し、残りの部分には適用しません。|
|正規表現|すべての Outlook アプリケーションの ItemHasKnownEntity または ItemHasRegularExpressionMatch ルールの場合:<br><ul><li>Outlook アドインのアクティブ化ルールで指定する正規表現は 5 個までにしてください。その制約数を超えるアドインをインストールすることはできません。</li><li>予期される結果が <b>getRegExMatches</b> メソッド呼び出しによって返されて、それらが最初の 50 件以内に収まるように、正規表現を指定します。 </li><li>**重要**: 文字列は、正規表現に一致した結果の文字列に基づいて強調表示されます。 ただし、強調表示された出現箇所は、負の先読み`(?!text)`、ルックビハインド、負のルックビハ`(?<=text)``(?<!text)`インドなどの実際の正規表現アサーションの結果と完全には一致しない可能性があります。 たとえば、"Like under,under score, and アンダースコア" の正規表現 `under(?!score)` を使用すると、最初の 2 つの文字列だけでなく、すべての出現箇所で文字列 "under" が強調表示されます。</li><li>一致が次の表の制限を超えていない正規表現を指定します。<br/><br/><table><tr><th>正規表現の長さ制限</th><th>Outlook リッチ クライアント</th><th>iOS および Android 用の Outlook</th></tr><tr><td>アイテムの本文がテキスト形式の場合</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>アイテムの本文が HTML の場合</td><td>3 KB</td><td>3 KB</td></tr></table>|

## <a name="see-also"></a>関連項目

- [新規作成フォーム用の Outlook アドインを作成する](compose-scenario.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [正規表現アクティブ化ルールを使用して Outlook アドインを表示する](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
